import { describe, expect, it } from 'vitest'
import { EXPORT_PNG_DPI, EXPORT_PNG_MAX_EDGE, exportPngScale } from '../src/renderer/export-images'

describe('exportPngScale', () => {
  it('uses 150dpi for letter-size pages', () => {
    const scale = exportPngScale(612, 792)
    expect(scale).toBeCloseTo(EXPORT_PNG_DPI / 72, 5)
    expect(Math.max(612, 792) * scale).toBeLessThanOrEqual(EXPORT_PNG_MAX_EDGE)
  })

  it('caps poster-size pages so PNG encode stays bounded', () => {
    const scale = exportPngScale(4000, 8000)
    expect(Math.max(4000, 8000) * scale).toBeCloseTo(EXPORT_PNG_MAX_EDGE, 5)
    expect(scale).toBeLessThan(EXPORT_PNG_DPI / 72)
  })

  it('falls back for invalid geometry', () => {
    expect(exportPngScale(0, 100)).toBeCloseTo(EXPORT_PNG_DPI / 72, 5)
    expect(exportPngScale(Number.NaN, 100)).toBeCloseTo(EXPORT_PNG_DPI / 72, 5)
  })
})
