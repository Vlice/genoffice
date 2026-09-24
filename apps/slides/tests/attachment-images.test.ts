import { describe, expect, it } from 'vitest'
import type { AttachmentMeta } from '../src/shared/ipc'
import {
  attachmentImageRefs,
  filterSlideImageRefs,
  mergeAttachmentRefs,
  seedCoverWithUnusedAttachments,
} from '../src/renderer/ai/attachment-images'

const photo: AttachmentMeta = { path: '/tmp/cat.png', name: 'cat.png', ext: 'png', sizeBytes: 12 }
const notes: AttachmentMeta = { path: '/tmp/n.md', name: 'n.md', ext: 'md', sizeBytes: 4 }

describe('attachment image helpers', () => {
  it('lists only image attachments as attachment:N refs', () => {
    expect(attachmentImageRefs([notes, photo])).toEqual(['attachment:1'])
  })

  it('keeps http and attachment refs, drops file urls', () => {
    expect(
      filterSlideImageRefs(['https://x/a.png', 'attachment:0', 'file:///etc/passwd', '  ']),
    ).toEqual(['https://x/a.png', 'attachment:0'])
  })

  it('appends unused user photos after model-supplied refs', () => {
    expect(mergeAttachmentRefs(['https://x/a.png'], [photo])).toEqual([
      'https://x/a.png',
      'attachment:0',
    ])
    expect(mergeAttachmentRefs(['attachment:0'], [photo])).toEqual(['attachment:0'])
  })

  it('seeds unused photos onto the first page only', () => {
    const pages = [
      { image_queries: ['great wall'] },
      { image_queries: [] },
    ]
    seedCoverWithUnusedAttachments(pages, [photo])
    expect(pages[0]!.image_queries).toEqual(['attachment:0', 'great wall'])
    expect(pages[1]!.image_queries).toEqual([])
  })

  it('does not duplicate a photo the model already placed', () => {
    const pages = [{ image_queries: ['attachment:0'] }, { image_queries: [] }]
    seedCoverWithUnusedAttachments(pages, [photo])
    expect(pages[0]!.image_queries).toEqual(['attachment:0'])
  })
})
