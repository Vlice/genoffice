import { WATERMARK_TILE_H, WATERMARK_TILE_W, watermarkTileBackground } from '../watermark-tile'

/** Per-page (or full-sheet) tiled diagonal watermark. */
export function PageWatermark({ text }: { text: string }) {
  return (
    <div
      className="page-watermark"
      aria-hidden="true"
      style={{
        backgroundImage: watermarkTileBackground(text),
        backgroundSize: `${WATERMARK_TILE_W}px ${WATERMARK_TILE_H}px`,
      }}
    />
  )
}
