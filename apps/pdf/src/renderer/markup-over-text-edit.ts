/**
 * When a text run is cleared or shortened, drop/clip markups that sat on the
 * vanished glyphs so a highlighter does not remain on blank paper.
 */
import { markupRectQuad } from '../shared/markup-quads'
import { unionCover, type LocalTextEdit } from './text-edit-preview'

export type PdfRect = [number, number, number, number]

export function inflateRect(r: PdfRect, pad: number): PdfRect {
  return [r[0] - pad, r[1] - pad, r[2] + pad, r[3] + pad]
}

function rectArea(r: PdfRect): number {
  return Math.max(0, r[2] - r[0]) * Math.max(0, r[3] - r[1])
}

export function intersectRects(a: PdfRect, b: PdfRect): PdfRect | null {
  const x1 = Math.max(a[0], b[0])
  const y1 = Math.max(a[1], b[1])
  const x2 = Math.min(a[2], b[2])
  const y2 = Math.min(a[3], b[3])
  if (x2 - x1 < 0.4 || y2 - y1 < 0.4) return null
  return [x1, y1, x2, y2]
}

export function rectsOverlap(a: PdfRect, b: PdfRect, minArea = 0.25): boolean {
  const hit = intersectRects(a, b)
  return hit ? rectArea(hit) >= minArea : false
}

function quadRect(q: number[]): PdfRect {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
}

/** Original ink box of a pending text edit (layout rect ∪ validated glyph bounds). */
export function textEditClearedRect(edit: LocalTextEdit): PdfRect | null {
  const r = edit.input.rect
  if (!r || r.length < 4) return null
  return inflateRect(unionCover(r, edit.cover), 6)
}

/**
 * Approximate the remaining glyph box after a replacement.
 * Empty / whitespace-only → nothing remains. If `newText` is a substring of
 * `oldText`, map that span onto the old box; otherwise shrink from the start
 * by character-count ratio (LTR runs).
 */
export function remainingTextRect(
  oldRect: PdfRect,
  oldText: string,
  newText: string,
): PdfRect | null {
  const neu = newText
  if (!neu.trim()) return null
  const [x1, y1, x2, y2] = oldRect
  const w = x2 - x1
  if (w <= 0) return null
  const oldLen = Math.max(1, oldText.length)
  let start = oldText.indexOf(neu)
  let span = neu.length
  if (start < 0) {
    start = 0
    span = Math.min(neu.length, oldLen)
  }
  if (span <= 0) return null
  const t0 = start / oldLen
  const t1 = Math.min(1, (start + span) / oldLen)
  if (t1 - t0 < 0.02) return null
  return [x1 + t0 * w, y1, x1 + t1 * w, y2]
}

/** Clip markup quads to the remaining ink; [] means the markup should be removed. */
export function clipMarkupQuadsForTextEdit(
  quads: number[][],
  cleared: PdfRect,
  remaining: PdfRect | null,
): number[][] {
  if (!quads.length) return quads
  const hitsCleared = quads.some((q) => rectsOverlap(quadRect(q), cleared))
  if (!hitsCleared) return quads
  if (!remaining) return []
  const next: number[][] = []
  for (const q of quads) {
    const box = quadRect(q)
    if (!rectsOverlap(box, cleared)) {
      next.push(q)
      continue
    }
    const keep = intersectRects(box, remaining)
    if (keep) next.push(markupRectQuad(keep))
  }
  return next
}

export function savedMarkupShouldDeleteForTextEdit(
  quads: number[][],
  rect: PdfRect,
  cleared: PdfRect,
  remaining: PdfRect | null,
): boolean {
  const box = quads.length ? quads.map(quadRect).reduce((acc, r) => [
    Math.min(acc[0], r[0]),
    Math.min(acc[1], r[1]),
    Math.max(acc[2], r[2]),
    Math.max(acc[3], r[3]),
  ] as PdfRect) : rect
  if (!rectsOverlap(box, cleared)) return false
  if (!remaining) return true
  const stay = intersectRects(box, remaining)
  if (!stay) return true
  return rectArea(stay) < 0.35 * Math.max(rectArea(box), 1)
}
