import { describe, expect, it } from 'vitest'
import { MARKUP_COLORS, type LocalMarkup } from '../src/renderer/annotations'
import { planMarkupApply, rgbClose } from '../src/renderer/markup-apply'
import type { SavedMarkupAnnot } from '../src/renderer/edit-state'

const QUADS = [[10, 80, 90, 80, 10, 60, 90, 60]]
const RED: [number, number, number] = [1, 0, 0]
const YELLOW = MARKUP_COLORS.highlight

const pending = (over: Partial<LocalMarkup> = {}): LocalMarkup => ({
  id: 'm1',
  pageIndex: 0,
  type: 'highlight',
  color: YELLOW,
  quads: QUADS,
  ...over,
})

const saved = (over: Partial<SavedMarkupAnnot> = {}): SavedMarkupAnnot => ({
  pageIndex: 0,
  objNum: 12,
  type: 'highlight',
  quads: QUADS,
  rect: [10, 60, 90, 80],
  ...over,
})

describe('rgbClose', () => {
  it('treats 8-bit-adjacent floats as equal', () => {
    expect(rgbClose(YELLOW, YELLOW)).toBe(true)
    expect(rgbClose(YELLOW, [1, 0.87, 0.35 + 0.5 / 255])).toBe(true)
    expect(rgbClose(YELLOW, RED)).toBe(false)
  })
})

describe('planMarkupApply toggle', () => {
  it('adds highlight to an unmarked selection', () => {
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'toggle',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [] }],
    })
    expect(plan?.add).toEqual([{ pageIndex: 0, type: 'highlight', color: RED, quads: QUADS }])
    expect(plan?.removePendingIds).toEqual([])
  })

  it('removes when every page is already marked, even if the color differs', () => {
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'toggle',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: pending() }] }],
    })
    expect(plan).toEqual({
      color: RED,
      removePendingIds: ['m1'],
      recolorPendingIds: [],
      deleteSaved: [],
      add: [],
    })
  })

  it('removes every stacked layer instead of adding another', () => {
    const a = pending({ id: 's1', type: 'strikeout' })
    const b = pending({ id: 's2', type: 'strikeout' })
    const plan = planMarkupApply({
      type: 'strikeout',
      color: MARKUP_COLORS.strikeout,
      intent: 'toggle',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: a }, { pending: b }] }],
    })
    expect(plan?.removePendingIds).toEqual(['s1', 's2'])
    expect(plan?.add).toEqual([])
  })

  it('replaces a partial mark with one covering the whole selection (no stack)', () => {
    const small = pending({
      id: 'part',
      type: 'strikeout',
      quads: [[10, 80, 40, 80, 10, 60, 40, 60]],
    })
    const plan = planMarkupApply({
      type: 'strikeout',
      color: MARKUP_COLORS.strikeout,
      intent: 'toggle',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: small }] }],
    })
    expect(plan?.removePendingIds).toEqual(['part'])
    expect(plan?.add).toEqual([
      { pageIndex: 0, type: 'strikeout', color: MARKUP_COLORS.strikeout, quads: QUADS },
    ])
  })

  it('marks unmarked pages of a partial selection without stacking on marked pages', () => {
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'toggle',
      matches: [
        { origIdx: 0, quads: QUADS, overlapping: [{ pending: pending() }] },
        { origIdx: 1, quads: QUADS, overlapping: [] },
      ],
    })
    expect(plan?.removePendingIds).toEqual(['m1'])
    expect(plan?.add).toEqual([
      { pageIndex: 0, type: 'highlight', color: RED, quads: QUADS },
      { pageIndex: 1, type: 'highlight', color: RED, quads: QUADS },
    ])
  })
})

describe('planMarkupApply apply (palette pick)', () => {
  it('adds highlight when the selection is unmarked', () => {
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'apply',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [] }],
    })
    expect(plan?.add).toHaveLength(1)
    expect(plan?.add[0]?.color).toEqual(RED)
    expect(plan?.removePendingIds).toEqual([])
  })

  it('recolors an existing pending highlight instead of toggling it off', () => {
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'apply',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: pending() }] }],
    })
    expect(plan).toEqual({
      color: RED,
      removePendingIds: [],
      recolorPendingIds: ['m1'],
      deleteSaved: [],
      add: [],
    })
  })

  it('no-ops when the pending highlight already has the picked color', () => {
    expect(
      planMarkupApply({
        type: 'highlight',
        color: YELLOW,
        intent: 'apply',
        matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: pending() }] }],
      }),
    ).toBeNull()
  })

  it('replaces a saved highlight so the new color can be written on save', () => {
    const annot = saved()
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'apply',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ saved: annot }] }],
    })
    expect(plan?.deleteSaved).toEqual([annot])
    expect(plan?.add).toEqual([{ pageIndex: 0, type: 'highlight', color: RED, quads: QUADS }])
    expect(plan?.removePendingIds).toEqual([])
  })

  it('collapses stacked pending highlights to one color', () => {
    const a = pending({ id: 'h1' })
    const b = pending({ id: 'h2' })
    const plan = planMarkupApply({
      type: 'highlight',
      color: RED,
      intent: 'apply',
      matches: [{ origIdx: 0, quads: QUADS, overlapping: [{ pending: a }, { pending: b }] }],
    })
    expect(plan?.removePendingIds).toEqual(['h1', 'h2'])
    expect(plan?.add).toEqual([{ pageIndex: 0, type: 'highlight', color: RED, quads: QUADS }])
  })
})
