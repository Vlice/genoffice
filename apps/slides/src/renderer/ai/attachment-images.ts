import type { AttachmentMeta } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import { isSlideImageRef, parseAttachmentIndex } from '../../shared/image-refs'

export function attachmentImageRefs(atts: AttachmentMeta[]): string[] {
  return atts.flatMap((a, i) => (ATTACHMENT_IMAGE_EXTS.has(a.ext) ? [`attachment:${i}`] : []))
}

export function filterSlideImageRefs(raw: unknown): string[] {
  if (!Array.isArray(raw)) return []
  return raw
    .map((x) => String(x).trim())
    .filter((s) => s.length > 0 && isSlideImageRef(s))
}

/** Keep model-supplied refs, then append any unused user-uploaded photos. */
export function mergeAttachmentRefs(explicit: string[], atts: AttachmentMeta[]): string[] {
  const extras = attachmentImageRefs(atts).filter((r) => !explicit.includes(r))
  return [...explicit, ...extras]
}

/**
 * If the model forgot to place user photos, put unused attachment refs on page 1
 * so a "put this picture on the cover" request still embeds the file.
 */
export function seedCoverWithUnusedAttachments(
  pages: Array<{ image_queries?: unknown }>,
  atts: AttachmentMeta[],
): void {
  const refs = attachmentImageRefs(atts)
  if (!refs.length || pages.length === 0) return
  const used = new Set<string>()
  for (const p of pages) {
    if (!Array.isArray(p.image_queries)) continue
    for (const q of p.image_queries) {
      const s = String(q).trim()
      if (refs.includes(s)) used.add(s)
    }
  }
  const missing = refs.filter((r) => !used.has(r))
  if (!missing.length) return
  const first = pages[0]!
  const existing = Array.isArray(first.image_queries) ? first.image_queries.map(String) : []
  first.image_queries = [...missing, ...existing]
}

export function isImageAttachment(att: AttachmentMeta | undefined): att is AttachmentMeta {
  return !!att && ATTACHMENT_IMAGE_EXTS.has(att.ext)
}

export async function collectAttachmentImageBytes(
  images: string[],
  atts: AttachmentMeta[],
): Promise<Record<string, { base64: string; ext: string }>> {
  const out: Record<string, { base64: string; ext: string }> = {}
  for (const ref of images) {
    const idx = parseAttachmentIndex(ref)
    if (idx == null) continue
    const att = atts[idx]
    if (!isImageAttachment(att)) continue
    const r = await window.desktop.readAttachmentImage(att.path)
    if (!r.ok || !r.base64) continue
    out[ref] = { base64: r.base64, ext: att.ext === 'jpeg' ? 'jpg' : att.ext }
  }
  return out
}

export { parseAttachmentIndex }
