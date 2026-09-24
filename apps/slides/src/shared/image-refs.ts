/** Image sources the slide generator and insert tools may embed. */

export const HTTP_IMAGE_RE = /^https?:\/\//i
export const ATTACHMENT_REF_RE = /^attachment:(\d+)$/i
export const DATA_IMAGE_RE =
  /^data:image\/(png|jpe?g|gif|webp);base64,([A-Za-z0-9+/=\s]+)$/i

export function isHttpImageUrl(s: string): boolean {
  return HTTP_IMAGE_RE.test(s.trim())
}

/** `attachment:0` → 0; otherwise null. */
export function parseAttachmentIndex(s: string): number | null {
  const m = s.trim().match(ATTACHMENT_REF_RE)
  if (!m) return null
  const n = Number(m[1])
  return Number.isInteger(n) && n >= 0 ? n : null
}

/** http(s) URL, `attachment:N`, or a data:image base64 URL. */
export function isSlideImageRef(s: string): boolean {
  const t = s.trim()
  return isHttpImageUrl(t) || parseAttachmentIndex(t) != null || DATA_IMAGE_RE.test(t)
}

export function decodeDataImageUrl(url: string): { base64: string; ext: string } | null {
  const m = DATA_IMAGE_RE.exec(url.trim())
  if (!m) return null
  const kind = m[1]!.toLowerCase()
  const ext = kind === 'jpeg' || kind === 'jpg' ? 'jpg' : kind
  return { base64: m[2]!.replace(/\s/g, ''), ext }
}

export function dataImageBytes(url: string): { bytes: Uint8Array; ext: string } | null {
  const decoded = decodeDataImageUrl(url)
  if (!decoded) return null
  return { bytes: base64ToBytes(decoded.base64), ext: decoded.ext }
}

function base64ToBytes(b64: string): Uint8Array {
  if (typeof Buffer !== 'undefined') return Uint8Array.from(Buffer.from(b64, 'base64'))
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
