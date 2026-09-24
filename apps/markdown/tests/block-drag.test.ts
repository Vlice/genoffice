import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { moveTopLevelBlock } from '../src/renderer/editor/blockDragHandle'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(md: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

function posOf(editor: Editor, type: string): number {
  let found = -1
  editor.state.doc.forEach((node, pos) => {
    if (node.type.name === type) found = pos
  })
  return found
}

function lastContentType(editor: Editor): string | undefined {
  const doc = editor.state.doc
  for (let i = doc.childCount - 1; i >= 0; i--) {
    const child = doc.child(i)
    if (child.type.name === 'paragraph' && child.content.size === 0) continue
    return child.type.name
  }
}

function tables(editor: Editor): { count: number; texts: string[] } {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'table') texts.push(node.textContent)
  })
  return { count: texts.length, texts }
}

const DOC = [
  '# heading',
  '',
  '| Name | Value |',
  '| --- | --- |',
  '| alpha | 1 |',
  '| beta | 2 |',
  '',
  'tail paragraph',
].join('\n')

describe('block drag move (tables)', () => {
  it('keeps the table NodeSelection when allowTableNodeSelection is on', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'table')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, from)))
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    expect((editor.state.selection as NodeSelection).node.type.name).toBe('table')
  })

  it('native deleteSelection on a table NodeSelection removes the whole table', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'table')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, from)))
    editor.view.dispatch(editor.state.tr.deleteSelection())
    expect(tables(editor).count).toBe(0)
  })

  it('moves the table as one block and does not leave an empty copy', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'table')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, from)))
    const tr = moveTopLevelBlock(editor.state, from, 0)
    expect(tr).not.toBeNull()
    editor.view.dispatch(tr!)
    const after = tables(editor)
    expect(after.count).toBe(1)
    expect(after.texts[0]).toContain('alpha')
    expect(after.texts[0]).toContain('beta')
    expect(editor.state.doc.firstChild?.type.name).toBe('table')
    expect(editor.getMarkdown()).toContain('| alpha')
  })

  it('can drop the table after the last block', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'table')
    const tr = moveTopLevelBlock(editor.state, from, editor.state.doc.content.size)
    editor.view.dispatch(tr!)
    expect(lastContentType(editor)).toBe('table')
    expect(tables(editor).count).toBe(1)
    expect(tables(editor).texts[0]).toContain('alpha')
  })

  it('is a no-op when dropping onto the same block', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'table')
    const node = editor.state.doc.nodeAt(from)!
    expect(moveTopLevelBlock(editor.state, from, from)).toBeNull()
    expect(moveTopLevelBlock(editor.state, from, from + node.nodeSize)).toBeNull()
  })

  it('still moves a non-table block', () => {
    const editor = createEditor(DOC)
    const from = posOf(editor, 'heading')
    const tr = moveTopLevelBlock(editor.state, from, editor.state.doc.content.size)
    editor.view.dispatch(tr!)
    expect(lastContentType(editor)).toBe('heading')
  })
})
