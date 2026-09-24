import { describe, expect, it } from 'vitest'
import { WATERMARK_TILE_H, WATERMARK_TILE_W, watermarkTileBackground } from '../src/renderer/watermark-tile'

describe('watermark tile', () => {
  it('encodes the text into a repeating SVG background', () => {
    const bg = watermarkTileBackground('test')
    expect(bg.startsWith('url("data:image/svg+xml,')).toBe(true)
    expect(bg).toContain(encodeURIComponent('test'))
    expect(WATERMARK_TILE_W).toBeGreaterThan(100)
    expect(WATERMARK_TILE_H).toBeGreaterThan(100)
  })
})
