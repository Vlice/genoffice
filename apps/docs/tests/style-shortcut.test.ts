/**
 * Issue #126 wave 2: paragraph-style and line-spacing commands backing the
 * ⌥⌘0-3 / ⌘1·2·5 shortcuts. applyParagraphStyle was extracted from the ribbon
 * gallery closure; these pin its Word-like behaviors (node switch + shedding
 * the runs' direct font/size/color).
 *
 * After save, parsed headings carry w:pStyle (styleId=Heading1). TipTap setNode
 * copies that attr, so switching to Heading 2 / Normal must overwrite it or
 * [data-style] CSS keeps the old look.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import type { StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyParagraphStyle, setParaAttrs } from '../src/renderer/components/ribbon-tabs'
import { styleIdForOutlineLevel } from '../src/renderer/editor/headings'

const styled = { type: 'docTextStyle', attrs: { sizeHalfPoints: 48, color: 'FF0000' } }

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'chapter title', marks: [styled] }],
        },
        { type: 'docParagraph', content: [{ type: 'text', text: 'body text' }] },
      ],
    },
  })
}

function savedHeadingEditor(styleId = 'Heading1', level = 1): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docHeading',
          attrs: { level, styleId },
          content: [{ type: 'text', text: 'chapter title' }],
        },
      ],
    },
  })
}

describe('applyParagraphStyle', () => {
  it('switches the block to a heading and sheds direct size/color', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h2')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docHeading')
    expect(block.attrs.level).toBe(2)
    expect(block.attrs.styleId).toBe('Heading2')
    const marks = block.firstChild!.marks.filter((m) => m.type.name === 'docTextStyle')
    expect(marks.length).toBe(0)
    editor.destroy()
  })

  it('returns a heading to a normal paragraph', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h1')
    applyParagraphStyle(editor, 'p')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docParagraph')
    expect(block.attrs.styleId).toBe('Normal')
    editor.destroy()
  })

  it('after save, switching Heading 1 → Heading 2 overwrites the parsed styleId', () => {
    const editor = savedHeadingEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h2')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docHeading')
    expect(block.attrs.level).toBe(2)
    expect(block.attrs.styleId).toBe('Heading2')
    expect(editor.view.dom.querySelector('h2')?.getAttribute('data-style')).toBe('Heading2')
    editor.destroy()
  })

  it('after save, Heading 1 → Normal drops the heading pStyle', () => {
    const editor = savedHeadingEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'p')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docParagraph')
    expect(block.attrs.styleId).toBe('Normal')
    editor.destroy()
  })

  it('uses the document heading style ids when they are not HeadingN', () => {
    const editor = savedHeadingEditor('标题1', 1)
    editor.storage.listNumbering.styles = new Map<string, StyleInfo>([
      ['正文', { styleId: '正文', name: '正文', type: 'paragraph', isDefault: true }],
      ['标题1', { styleId: '标题1', name: '标题 1', type: 'paragraph', headingLevel: 1 }],
      ['标题2', { styleId: '标题2', name: '标题 2', type: 'paragraph', headingLevel: 2 }],
    ])
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h2')
    expect(editor.state.doc.child(0).attrs.styleId).toBe('标题2')
    applyParagraphStyle(editor, 'p')
    expect(editor.state.doc.child(0).attrs.styleId).toBe('正文')
    editor.destroy()
  })
})

describe('styleIdForOutlineLevel', () => {
  it('falls back to HeadingN / Normal without a styles map', () => {
    expect(styleIdForOutlineLevel(undefined, 0)).toBe('Normal')
    expect(styleIdForOutlineLevel(undefined, 1)).toBe('Heading1')
    expect(styleIdForOutlineLevel(undefined, 3)).toBe('Heading3')
  })
})

describe('line spacing via setParaAttrs (⌘1/⌘2/⌘5 path)', () => {
  it('sets the multiple on every paragraph in the selection and clears exact rules', () => {
    const editor = makeEditor()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3, 20)),
    )
    setParaAttrs(editor, { lineSpacing: 1.5, lineRule: null, lineRawTwips: null })
    expect(editor.state.doc.child(0).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(1).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(0).attrs.lineRule).toBeNull()
    editor.destroy()
  })
})
