import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { StyleInfo } from '@genoffice/docx-engine'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  effectiveSizeHalfPoints,
  isEffectivelyMarked,
  toggleStyleAwareMark,
} from '../src/renderer/editor/text-style-resolve'

const editors: Editor[] = []

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.length = 0
})

function createEditor(options?: {
  paragraphStyleId?: string
  characterStyleId?: string
  directSize?: number
  bold?: boolean
  boldOff?: boolean
}): Editor {
  const marks: Array<{ type: string; attrs?: Record<string, unknown> }> = []
  if (options?.bold) marks.push({ type: 'bold' })
  if (options?.characterStyleId || options?.directSize || options?.boldOff) {
    marks.push({
      type: 'docTextStyle',
      attrs: {
        styleId: options.characterStyleId ?? null,
        sizeHalfPoints: options.directSize ?? null,
        boldOff: options.boldOff || null,
      },
    })
  }
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null, styleId: options?.paragraphStyleId ?? null },
          content: [{ type: 'text', text: 'Styled text', marks: marks.length ? marks : undefined }],
        },
      ],
    },
  })
  editor.commands.setTextSelection({ from: 1, to: 12 })
  editors.push(editor)
  return editor
}

function styles(
  ...entries: Array<[string, number] | [string, NonNullable<StyleInfo['display']>]>
): Map<string, StyleInfo> {
  return new Map(
    entries.map(([styleId, display]) => [
      styleId,
      {
        styleId,
        name: styleId,
        type: styleId.startsWith('Char') ? 'character' : 'paragraph',
        display: typeof display === 'number' ? { sizeHalfPoints: display } : display,
      },
    ]),
  )
}

describe('effectiveSizeHalfPoints', () => {
  it('resolves the paragraph style size at the caret', () => {
    const editor = createEditor({ paragraphStyleId: 'Title' })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56]))).toBe(56)
  })

  it('resolves a character style before the paragraph style', () => {
    const editor = createEditor({
      paragraphStyleId: 'Title',
      characterStyleId: 'CharEmphasis',
    })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56], ['CharEmphasis', 28]))).toBe(28)
  })

  it('prefers direct formatting over styles', () => {
    const editor = createEditor({ paragraphStyleId: 'Title', directSize: 24 })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56]))).toBe(24)
  })

  it('falls back to document defaults', () => {
    const editor = createEditor()
    expect(effectiveSizeHalfPoints(editor, undefined, { sizeHalfPoints: 21 })).toBe(21)
  })
})

describe('style-aware bold toggle', () => {
  it('treats heading-style bold as effectively on without a bold mark', () => {
    const editor = createEditor({ paragraphStyleId: 'Heading1' })
    const map = styles(['Heading1', { bold: true, sizeHalfPoints: 32 }])
    expect(editor.isActive('bold')).toBe(false)
    expect(isEffectivelyMarked(editor, 'bold', map)).toBe(true)
  })

  it('unbolds style-inherited text by writing boldOff', () => {
    const editor = createEditor({ paragraphStyleId: 'Heading1' })
    const map = styles(['Heading1', { bold: true }])
    editor.storage.listNumbering.styles = map
    expect(toggleStyleAwareMark(editor, 'bold', map)).toBe(true)
    expect(editor.isActive('bold')).toBe(false)
    expect(editor.getAttributes('docTextStyle').boldOff).toBe(true)
    expect(isEffectivelyMarked(editor, 'bold', map)).toBe(false)
    const span = editor.view.dom.querySelector<HTMLElement>('span[data-doc-style]')
    expect(span?.style.fontWeight).toBe('normal')
  })

  it('re-bolds by clearing boldOff without adding a redundant bold mark', () => {
    const editor = createEditor({ paragraphStyleId: 'Heading1', boldOff: true })
    const map = styles(['Heading1', { bold: true }])
    expect(isEffectivelyMarked(editor, 'bold', map)).toBe(false)
    expect(toggleStyleAwareMark(editor, 'bold', map)).toBe(true)
    expect(editor.getAttributes('docTextStyle').boldOff).toBeFalsy()
    expect(editor.isActive('bold')).toBe(false)
    expect(isEffectivelyMarked(editor, 'bold', map)).toBe(true)
  })

  it('still toggles a plain bold mark on body text', () => {
    const editor = createEditor()
    expect(toggleStyleAwareMark(editor, 'bold')).toBe(true)
    expect(editor.isActive('bold')).toBe(true)
    expect(toggleStyleAwareMark(editor, 'bold')).toBe(true)
    expect(editor.isActive('bold')).toBe(false)
  })

  it('saves style unbold as w:b w:val="0"', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>',
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    editor.storage.listNumbering.styles = parsed.styles
    editor.commands.setTextSelection({ from: 1, to: 6 })
    toggleStyleAwareMark(editor, 'bold', parsed.styles)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = plan.saveBlocks[0]
    expect(saved.kind).toBe('generated')
    if (saved.kind !== 'generated') throw new Error('expected generated block')
    expect(saved.block.runs[0]?.bold).toBe(false)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks[0].runs![0].bold).toBe(false)
  })
})
