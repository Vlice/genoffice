import { describe, expect, it } from 'vitest'
import { markupQuadRect, unifyLineQuads } from '../src/shared/markup-quads'

const q = (x1: number, y1: number, x2: number, y2: number) => [x1, y2, x2, y2, x1, y1, x2, y1]

describe('unifyLineQuads', () => {
  it('leaves a single quad unchanged', () => {
    const one = [q(10, 20, 40, 32)]
    expect(unifyLineQuads(one)).toEqual(one)
  })

  it('merges same-line boxes of mixed height onto one baseline', () => {
    // Latin vs CJK on one line: different bottoms would zigzag an underline
    const merged = unifyLineQuads([q(10, 22, 40, 34), q(42, 18, 80, 36)])
    expect(merged).toHaveLength(1)
    expect(markupQuadRect(merged[0]!)).toEqual([10, 18, 80, 36])
  })

  it('closes small gaps so the underline is not dashed', () => {
    const merged = unifyLineQuads([q(10, 20, 30, 32), q(34, 20, 60, 32)])
    expect(merged).toHaveLength(1)
    expect(markupQuadRect(merged[0]!)).toEqual([10, 20, 60, 32])
  })

  it('keeps a large gap as separate segments', () => {
    const merged = unifyLineQuads([q(10, 20, 30, 32), q(80, 20, 100, 32)])
    expect(merged).toHaveLength(2)
  })

  it('does not join two distinct lines', () => {
    const merged = unifyLineQuads([q(10, 50, 80, 64), q(10, 20, 80, 34)])
    expect(merged).toHaveLength(2)
  })

  it('merges along PDF y when the page is rotated 90°', () => {
    const merged = unifyLineQuads([q(20, 10, 36, 40), q(18, 42, 34, 80)], 90)
    expect(merged).toHaveLength(1)
    expect(markupQuadRect(merged[0]!)).toEqual([18, 10, 36, 80])
  })
})
