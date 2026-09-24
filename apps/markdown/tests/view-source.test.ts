import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { markdownFromEditor } from '../src/renderer/App'
import { buildExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
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

describe('markdownFromEditor', () => {
  it('serializes a GFM table without throwing', () => {
    const editor = createEditor('| 字段 | 类型 |\n| --- | --- |\n| name | string |\n')
    const md = markdownFromEditor(editor)
    expect(md).toContain('|')
    expect(md).toContain('name')
  })

  it('returns the fallback when serialize throws', () => {
    const editor = {
      getMarkdown: () => {
        throw new Error('boom')
      },
    } as unknown as Editor
    expect(markdownFromEditor(editor, '# kept')).toBe('# kept')
  })
})
