import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeTool, parseAttachmentIndex } from '../src/renderer/ai/tools'
import type { AttachmentMeta } from '../src/shared/ipc'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'announcement' }],
        },
      ],
    },
  })
  editors.add(editor)
  return editor
}

const NUM_IDS = { bullet: null, ordered: null }
const PNG_B64 = 'iVBORw0KGgoAAAAA'
const PHOTO: AttachmentMeta = {
  path: '/tmp/photo.png',
  name: 'photo.png',
  ext: 'png',
  sizeBytes: 2048,
}

/** jsdom never decodes images; fake one that reports a fixed natural size */
class FakeImage {
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  naturalWidth = 100
  naturalHeight = 80
  set src(_v: string) {
    queueMicrotask(() => this.onload?.())
  }
}

type DesktopStub = { desktop?: unknown }

async function runInsert(
  desktop: Record<string, unknown>,
  input: Record<string, unknown>,
  getAttachments?: () => AttachmentMeta[],
) {
  const editor = createEditor()
  const w = window as unknown as DesktopStub
  const saved = w.desktop
  const savedImage = globalThis.Image
  w.desktop = desktop
  globalThis.Image = FakeImage as unknown as typeof Image
  try {
    return {
      exec: await executeTool(
        editor,
        { id: 't', name: 'insert_image', input },
        NUM_IDS,
        undefined,
        undefined,
        null,
        undefined,
        undefined,
        getAttachments,
      ),
      editor,
    }
  } finally {
    w.desktop = saved
    globalThis.Image = savedImage
  }
}

describe('parseAttachmentIndex', () => {
  it('parses attachment:N refs and rejects everything else', () => {
    expect(parseAttachmentIndex('attachment:0')).toBe(0)
    expect(parseAttachmentIndex('attachment:12')).toBe(12)
    expect(parseAttachmentIndex('https://example.com/a.png')).toBeNull()
    expect(parseAttachmentIndex('file:///tmp/photo.png')).toBeNull()
    expect(parseAttachmentIndex('attachment:-1')).toBeNull()
  })
})

describe('insert_image attachment:N', () => {
  it('embeds the local attachment bytes and does not generate or fetch a new image', async () => {
    let generated = false
    let fetched: string | null = null
    const { exec, editor } = await runInsert(
      {
        aiGenerateImage: () => {
          generated = true
          return Promise.resolve({ url: 'https://example.com/gen.png' })
        },
        fetchImage: (url: string) => {
          fetched = url
          return Promise.resolve({ base64: PNG_B64, mime: 'image/png' })
        },
        readAttachmentImage: (path: string) => {
          expect(path).toBe(PHOTO.path)
          return Promise.resolve({ ok: true, base64: PNG_B64, mime: 'image/png' })
        },
      },
      { url: 'attachment:0' },
      () => [PHOTO],
    )
    expect(exec.isError).toBeFalsy()
    expect(generated).toBe(false)
    expect(fetched).toBeNull()
    const img = [...Array(editor.state.doc.childCount).keys()]
      .map((i) => editor.state.doc.child(i))
      .find((n) => n.type.name === 'docProtected' && n.attrs.blockType === 'image')
    expect(img).toBeTruthy()
    expect(img?.attrs.label).toContain('photo.png')
    expect(img?.attrs.genImage?.base64).toBe(PNG_B64)
    expect(img?.attrs.genImage?.mime).toBe('image/png')
  })

  it('rejects a missing attachment index without calling generate_image', async () => {
    let generated = false
    const { exec, editor } = await runInsert(
      {
        aiGenerateImage: () => {
          generated = true
          return Promise.resolve({ url: 'https://example.com/gen.png' })
        },
      },
      { url: 'attachment:0' },
      () => [],
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('invalid attachment index')
    expect(generated).toBe(false)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('still rejects file:// and other non-http urls', async () => {
    const { exec, editor } = await runInsert({}, { url: 'file:///tmp/photo.png' })
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('invalid url')
    expect(editor.state.doc.childCount).toBe(1)
  })
})
