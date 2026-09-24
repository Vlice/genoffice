import { describe, expect, it } from 'vitest'
import {
  clipMarkupQuadsForTextEdit,
  remainingTextRect,
  savedMarkupShouldDeleteForTextEdit,
} from '../src/renderer/markup-over-text-edit'
import { markupRectQuad } from '../src/shared/markup-quads'

const box: [number, number, number, number] = [10, 20, 110, 32]
const hl = [markupRectQuad([40, 20, 80, 32])]

describe('remainingTextRect', () => {
  it('returns null when the replacement is empty', () => {
    expect(remainingTextRect(box, 'hello', '')).toBeNull()
    expect(remainingTextRect(box, 'hello', '   ')).toBeNull()
  })

  it('maps a prefix onto the left of the old box', () => {
    const r = remainingTextRect(box, 'abcdef', 'abc')
    expect(r?.[0]).toBeCloseTo(10)
    expect(r?.[2]).toBeCloseTo(10 + 100 * 0.5)
  })
})

describe('clipMarkupQuadsForTextEdit', () => {
  it('drops a highlight when the whole run is cleared', () => {
    expect(clipMarkupQuadsForTextEdit(hl, box, null)).toEqual([])
  })

  it('leaves a highlight that does not meet the run', () => {
    const other = [markupRectQuad([200, 20, 240, 32])]
    expect(clipMarkupQuadsForTextEdit(other, box, null)).toEqual(other)
  })

  it('keeps only the overlap with remaining glyphs', () => {
    const remaining: [number, number, number, number] = [10, 20, 50, 32]
    const next = clipMarkupQuadsForTextEdit(hl, box, remaining)
    expect(next).toHaveLength(1)
    expect(next[0]![0]).toBeCloseTo(40)
    expect(next[0]![2]).toBeCloseTo(50)
  })
})

describe('savedMarkupShouldDeleteForTextEdit', () => {
  it('deletes a saved highlight sitting on cleared paper', () => {
    expect(savedMarkupShouldDeleteForTextEdit(hl, box, box, null)).toBe(true)
  })

  it('keeps a saved highlight that still sits on remaining text', () => {
    const remaining: [number, number, number, number] = [10, 20, 90, 32]
    expect(savedMarkupShouldDeleteForTextEdit(hl, box, box, remaining)).toBe(false)
  })
})

describe('overlay insert highlight clip', () => {
  it('drops both wrapped highlight bars when the insert is cleared', () => {
    const insertBox: [number, number, number, number] = [72, 700, 420, 728]
    const bars = [
      markupRectQuad([72, 714, 400, 728]),
      markupRectQuad([72, 700, 140, 713]),
    ]
    expect(clipMarkupQuadsForTextEdit(bars, insertBox, null)).toEqual([])
  })
})
