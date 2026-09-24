import { describe, expect, it } from 'vitest'
import type { PageGeom } from '../src/renderer/annotations'
import {
  INSERT_ALIGN_MARGIN,
  alignLineOffset,
  insertBoxWidth,
  insertBoxWidthFromView,
  textInsertPreviewStyle,
} from '../src/renderer/text-edit-preview'
import type { LocalTextInsert } from '../src/renderer/text-edit-preview'

const geom = (rot = 0, pw = 200, ph = 300): PageGeom => ({ pw, ph, rot })

const insert = (
  origin: [number, number],
  align: 'left' | 'center' | 'right',
): LocalTextInsert => ({
  id: 't',
  input: {
    pageIndex: 0,
    origin,
    text: '测试',
    fontSize: 14,
    color: [0, 112, 192],
    align,
  },
})

describe('alignLineOffset', () => {
  it('keeps left-aligned lines at the box origin', () => {
    expect(alignLineOffset('left', 40, 100)).toBe(0)
  })

  it('pushes right/center toward the right of the box, never left of the origin', () => {
    expect(alignLineOffset('center', 40, 100)).toBe(30)
    expect(alignLineOffset('right', 40, 100)).toBe(60)
    expect(alignLineOffset('right', 40, 100)).toBeGreaterThan(0)
  })

  it('does not go negative when the line is wider than the box', () => {
    expect(alignLineOffset('right', 120, 100)).toBe(0)
    expect(alignLineOffset('center', 120, 100)).toBe(0)
  })
})

describe('insertBoxWidth', () => {
  it('is the remaining display width from the origin to the right margin', () => {
    expect(insertBoxWidthFromView(geom(), 40)).toBe(200 - 40 - INSERT_ALIGN_MARGIN)
    expect(insertBoxWidth(geom(), [40, 250])).toBe(200 - 40 - INSERT_ALIGN_MARGIN)
  })

  it('never shrinks below the minimum box', () => {
    expect(insertBoxWidthFromView(geom(), 190)).toBe(24)
  })
})

describe('textInsertPreviewStyle', () => {
  it('anchors the box at the click and aligns inside it (no translateX flip)', () => {
    const style = textInsertPreviewStyle(insert([40, 250], 'right'), geom(), 1)
    expect(style.left).toBe(40)
    expect(style.textAlign).toBe('right')
    expect(style.width).toBe(200 - 40 - INSERT_ALIGN_MARGIN)
    expect(style.transform).toBeUndefined()
  })

  it('keeps left alignment at the origin without shifting the box', () => {
    const style = textInsertPreviewStyle(insert([40, 250], 'left'), geom(), 2)
    expect(style.left).toBe(80)
    expect(style.textAlign).toBe('left')
    expect(style.transform).toBeUndefined()
  })
})
