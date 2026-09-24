/** Rasterize PDF pages to PNG for 导出图片. Keep this free of scenario heuristics. */

export const EXPORT_PNG_DPI = 150
/** Browser tabs encode PNG much slower than Electron; 96dpi is still print-preview sharp. */
export const EXPORT_PNG_DPI_WEB = 96
export const EXPORT_PNG_MAX_EDGE = 4096

export function exportPngDpi(): number {
  if (typeof window === 'undefined') return EXPORT_PNG_DPI_WEB
  const w = window as Window & { electronAPI?: unknown; parent?: Window & { electronAPI?: unknown } }
  try {
    if (w.electronAPI || w.parent?.electronAPI) return EXPORT_PNG_DPI
  } catch {
    /* cross-origin parent */
  }
  return EXPORT_PNG_DPI_WEB
}

/**
 * 150dpi is sharp enough for page screenshots. Cap the long edge so poster-size
 * pages do not freeze the tab on PNG encode.
 */
export function exportPngScale(
  unscaledWidth: number,
  unscaledHeight: number,
  dpi = EXPORT_PNG_DPI,
  maxEdge = EXPORT_PNG_MAX_EDGE,
): number {
  const w = Number(unscaledWidth)
  const h = Number(unscaledHeight)
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return dpi / 72
  const scale = dpi / 72
  const longest = Math.max(w, h) * scale
  if (longest <= maxEdge) return scale
  return scale * (maxEdge / longest)
}

export function pngBytesToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const buf = bytes.buffer
  if (bytes.byteOffset === 0 && bytes.byteLength === buf.byteLength && buf instanceof ArrayBuffer) {
    return buf
  }
  return buf.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
}

function dataUrlToBytes(dataUrl: string): Uint8Array {
  const comma = dataUrl.indexOf(',')
  const b64 = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl
  const binary = atob(b64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

/** Async PNG encode — avoids the main-thread stall of `toDataURL`. */
export async function canvasToPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  const convertible = canvas as HTMLCanvasElement & {
    convertToBlob?: (opts?: { type?: string; quality?: number }) => Promise<Blob>
  }
  if (typeof convertible.convertToBlob === 'function') {
    const blob = await convertible.convertToBlob({ type: 'image/png' })
    if (blob) return new Uint8Array(await blob.arrayBuffer())
  }
  if (typeof canvas.toBlob === 'function') {
    const blob = await new Promise<Blob | null>((resolve) => {
      canvas.toBlob(resolve, 'image/png')
    })
    if (blob) return new Uint8Array(await blob.arrayBuffer())
  }
  return dataUrlToBytes(canvas.toDataURL('image/png'))
}

/** Paint unsaved ink/shapes from the on-screen DrawLayer onto the export bitmap. */
export async function paintDrawLayerOnCanvas(
  canvas: HTMLCanvasElement,
  pageEl: Element | null,
): Promise<void> {
  if (!pageEl) return
  const svg = pageEl.querySelector('svg.pdf-draw-layer')
  if (!(svg instanceof SVGSVGElement)) return
  const clone = svg.cloneNode(true) as SVGSVGElement
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
  clone.querySelectorAll('.pdf-draw-handle, title').forEach((node) => node.remove())
  const xml = new XMLSerializer().serializeToString(clone)
  const url = URL.createObjectURL(new Blob([xml], { type: 'image/svg+xml;charset=utf-8' }))
  try {
    const img = new Image()
    img.decoding = 'async'
    img.src = url
    await img.decode()
    const ctx = canvas.getContext('2d')
    if (ctx) ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
  } catch {
    /* overlay is optional */
  } finally {
    URL.revokeObjectURL(url)
  }
}
