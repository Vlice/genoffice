import type { MarkupType } from '../shared/ipc'
import { quadSetsCoveredBy } from './annotations'
import type { LocalMarkup } from './annotations'
import type { SavedMarkupAnnot } from './edit-state'

/** Channel-wise equality for 0–1 RGB (one 8-bit step). */
export function rgbClose(
  a: readonly [number, number, number],
  b: readonly [number, number, number],
  eps = 1 / 255,
): boolean {
  return a.every((v, i) => Math.abs(v - (b[i] ?? 0)) <= eps)
}

export type MarkupMatch = { pending: LocalMarkup } | { saved: SavedMarkupAnnot }

export interface MarkupApplyPage {
  origIdx: number
  quads: number[][]
  /** Every markup of this type whose region overlaps the selection (may be stacked). */
  overlapping: MarkupMatch[]
}

/** `toggle` = Word highlighter button (remove if already marked).
    `apply` = palette pick (set this color; never remove). */
export type MarkupApplyIntent = 'toggle' | 'apply'

export interface MarkupApplyPlan {
  color: [number, number, number]
  removePendingIds: string[]
  recolorPendingIds: string[]
  deleteSaved: SavedMarkupAnnot[]
  add: Array<Omit<LocalMarkup, 'id'>>
}

function emptyPlan(color: [number, number, number]): MarkupApplyPlan {
  return { color, removePendingIds: [], recolorPendingIds: [], deleteSaved: [], add: [] }
}

function planIsEmpty(plan: MarkupApplyPlan): boolean {
  return (
    plan.removePendingIds.length === 0 &&
    plan.recolorPendingIds.length === 0 &&
    plan.deleteSaved.length === 0 &&
    plan.add.length === 0
  )
}

function addMarkup(
  plan: MarkupApplyPlan,
  type: MarkupType,
  origIdx: number,
  quads: number[][],
): void {
  plan.add.push({ pageIndex: origIdx, type, color: plan.color, quads })
}

function matchQuads(m: MarkupMatch): number[][] {
  return 'pending' in m ? m.pending.quads : m.saved.quads
}

function dropOverlapping(plan: MarkupApplyPlan, overlapping: MarkupMatch[]): void {
  for (const m of overlapping) {
    if ('pending' in m) plan.removePendingIds.push(m.pending.id)
    else plan.deleteSaved.push(m.saved)
  }
}

function pageCovered(page: MarkupApplyPage): boolean {
  if (page.overlapping.length === 0) return false
  return quadSetsCoveredBy(
    page.quads,
    page.overlapping.flatMap((m) => matchQuads(m)),
  )
}

/**
 * Decide how to change markups for a selection. `matches` is one entry per
 * page the selection covers. Returns null when the UI should no-op.
 *
 * Toggle is Word-style: one click marks the whole selection; a second click
 * removes every overlapping markup of that type (never stacks another layer).
 */
export function planMarkupApply(opts: {
  type: MarkupType
  color: [number, number, number]
  intent: MarkupApplyIntent
  matches: MarkupApplyPage[]
}): MarkupApplyPlan | null {
  const { type, color, intent, matches } = opts
  if (matches.length === 0) return null
  const plan = emptyPlan(color)

  if (intent === 'toggle') {
    if (matches.every(pageCovered)) {
      for (const { overlapping } of matches) dropOverlapping(plan, overlapping)
      return planIsEmpty(plan) ? null : plan
    }
    for (const { origIdx, quads, overlapping } of matches) {
      dropOverlapping(plan, overlapping)
      addMarkup(plan, type, origIdx, quads)
    }
    return planIsEmpty(plan) ? null : plan
  }

  // apply: paint this color onto the selection (recolor existing, add missing).
  // Multiple overlapping layers collapse to one so a palette pick cannot stack.
  for (const { origIdx, quads, overlapping } of matches) {
    if (overlapping.length === 0) {
      addMarkup(plan, type, origIdx, quads)
      continue
    }
    const pending = overlapping.filter((m): m is { pending: LocalMarkup } => 'pending' in m)
    const saved = overlapping.filter((m): m is { saved: SavedMarkupAnnot } => 'saved' in m)
    if (pending.length === 1 && saved.length === 0) {
      if (!rgbClose(pending[0]!.pending.color, color)) {
        plan.recolorPendingIds.push(pending[0]!.pending.id)
      }
      continue
    }
    dropOverlapping(plan, overlapping)
    addMarkup(plan, type, origIdx, quads)
  }
  return planIsEmpty(plan) ? null : plan
}
