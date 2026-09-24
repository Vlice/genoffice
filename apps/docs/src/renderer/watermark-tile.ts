/** Repeating diagonal text tile used by canvas overlays and pagination preview. */

export const WATERMARK_TILE_W = 260
export const WATERMARK_TILE_H = 176

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function fontSizeFor(text: string): number {
  const n = [...text].length
  if (n <= 4) return 28
  if (n <= 8) return 22
  if (n <= 14) return 16
  return 13
}

/** CSS `background-image` value: one diagonal copy, tiled by the browser. */
export function watermarkTileBackground(text: string): string {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${WATERMARK_TILE_W}" height="${WATERMARK_TILE_H}"` +
    ` viewBox="0 0 ${WATERMARK_TILE_W} ${WATERMARK_TILE_H}">` +
    `<text x="50%" y="50%" fill="rgb(128,128,128)" fill-opacity="0.22" font-size="${fontSizeFor(text)}"` +
    ` font-weight="600" font-family="DengXian, Microsoft YaHei, PingFang SC, sans-serif"` +
    ` text-anchor="middle" dominant-baseline="middle"` +
    ` transform="rotate(-24 ${WATERMARK_TILE_W / 2} ${WATERMARK_TILE_H / 2})">${escapeXml(text)}</text>` +
    `</svg>`
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}
