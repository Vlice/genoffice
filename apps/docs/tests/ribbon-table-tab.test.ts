/**
 * Opening a document that contains a table must keep the Home tab.
 * Word only auto-activates Table Layout when the user enters a table this session.
 */
import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { tableModelToPmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

function makeEditorWithTable(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'before table' }] }],
    },
  })
  editor.commands.insertContentAt(
    editor.state.doc.content.size,
    tableModelToPmNode({
      rows: [
        [{ paras: ['a'] }, { paras: ['b'] }],
        [{ paras: ['c'] }, { paras: ['d'] }],
      ],
      colWidthsPct: [50, 50],
    }),
  )
  return editor
}

function firstCellPos(editor: Editor): number {
  let cellPos = -1
  editor.state.doc.descendants((node, pos) => {
    if (cellPos === -1 && ['docTableCell', 'docTableHeader'].includes(node.type.name)) {
      cellPos = pos
    }
  })
  return cellPos
}

function mountRibbon(editor: Editor, docEpoch = 0) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const render = (epoch = docEpoch) =>
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          docEpoch: epoch,
        }),
      ),
    )
  render()
  return { container, root, render }
}

function activeTab(container: HTMLElement): string | null {
  return container.querySelector('.ribbon-tab.active')?.textContent ?? null
}

describe('ribbon tab when a document contains a table', () => {
  it('stays on Home when opening a document lands the caret in a table', () => {
    const editor = makeEditorWithTable()
    editor.commands.setTextSelection(1)
    expect(computeFormatState(editor).inTable).toBe(false)

    const { container, root, render } = mountRibbon(editor, 0)
    expect(activeTab(container)).toBe(t('ribbonTabHome'))

    editor.commands.setTextSelection(firstCellPos(editor) + 2)
    render(1)
    expect(computeFormatState(editor).inTable).toBe(true)
    expect(activeTab(container)).toBe(t('ribbonTabHome'))
    expect(container.textContent).toContain(t('ribbonTabTableLayout'))

    act(() => root.unmount())
    editor.destroy()
  })

  it('switches to Table Layout when the user enters a table', () => {
    const editor = makeEditorWithTable()
    editor.commands.setTextSelection(1)
    expect(computeFormatState(editor).inTable).toBe(false)

    const { container, root, render } = mountRibbon(editor)
    expect(activeTab(container)).toBe(t('ribbonTabHome'))

    editor.commands.setTextSelection(firstCellPos(editor) + 2)
    render()
    expect(computeFormatState(editor).inTable).toBe(true)
    expect(activeTab(container)).toBe(t('ribbonTabTableLayout'))

    act(() => root.unmount())
    editor.destroy()
  })

  it('returns to Home when a new document is opened while the caret is in a table', () => {
    const editor = makeEditorWithTable()
    editor.commands.setTextSelection(1)
    const { container, root, render } = mountRibbon(editor, 0)

    editor.commands.setTextSelection(firstCellPos(editor) + 2)
    render(0)
    expect(activeTab(container)).toBe(t('ribbonTabTableLayout'))

    render(1)
    expect(activeTab(container)).toBe(t('ribbonTabHome'))
    expect(container.textContent).toContain(t('ribbonTabTableLayout'))

    act(() => root.unmount())
    editor.destroy()
  })
})
