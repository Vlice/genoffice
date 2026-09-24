/**
 * Markup QuadPoints are [x1,yMax,x2,yMax,x1,yMin,x2,yMin] in PDF user space.
 * Text-layer getClientRects() yields one box per span; mixed faces (Latin vs CJK)
 * have different heights, so per-box underline/strikeout reads as a broken polyline.
 */

export function markupQuadRect(q: number[]): [number, number, number, number] {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
}

export function markupRectQuad(r: readonly [number, number, number, number]): number[] {
  const [x1, y1, x2, y2] = r
  return [x1, y2, x2, y2, x1, y1, x2, y1]
}

function yOverlap(
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): number {
  return Math.max(0, Math.min(a[3], b[3]) - Math.max(a[1], b[1]))
}

function xOverlap(
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): number {
  return Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0]))
}

function onSameLine(
  line: { x1: number; y1: number; x2: number; y2: number },
  r: readonly [number, number, number, number],
  alongX: boolean,
): boolean {
  const lineRect: [number, number, number, number] = [line.x1, line.y1, line.x2, line.y2]
  if (alongX) {
    const h = Math.min(line.y2 - line.y1, r[3] - r[1])
    return h <= 0 || yOverlap(lineRect, r) > 0.4 * h
  }
  const w = Math.min(line.x2 - line.x1, r[2] - r[0])
  return w <= 0 || xOverlap(lineRect, r) > 0.4 * w
}

/**
 * Collapse same-line quads onto a shared baseline and join small gaps.
 * `pageRot` is the displayed page rotation: at 90/270 the line runs along PDF y.
 */
export function unifyLineQuads(quads: number[][], pageRot = 0): number[][] {
  if (quads.length <= 1) return quads
  const rot = ((pageRot % 360) + 360) % 360
  const alongX = rot % 180 === 0
  const rects = quads.map(markupQuadRect)
  const order = rects
    .map((_, i) => i)
    .sort((a, b) =>
      alongX
        ? rects[b]![3] - rects[a]![3] || rects[a]![0] - rects[b]![0]
        : rects[a]![0] - rects[b]![0] || rects[b]![3] - rects[a]![3],
    )

  const lines: {
    x1: number
    y1: number
    x2: number
    y2: number
    items: [number, number, number, number][]
  }[] = []
  for (const i of order) {
    const r = rects[i]!
    const line = lines.find((ln) => onSameLine(ln, r, alongX))
    if (line) {
      line.items.push(r)
      line.x1 = Math.min(line.x1, r[0])
      line.y1 = Math.min(line.y1, r[1])
      line.x2 = Math.max(line.x2, r[2])
      line.y2 = Math.max(line.y2, r[3])
    } else {
      lines.push({ x1: r[0], y1: r[1], x2: r[2], y2: r[3], items: [r] })
    }
  }

  const out: number[][] = []
  for (const line of lines) {
    const { x1: lx1, y1, x2: lx2, y2, items } = line
    if (alongX) {
      const gapTol = Math.max(2, (y2 - y1) * 0.35)
      const sorted = [...items].sort((a, b) => a[0] - b[0])
      let x1 = sorted[0]![0]
      let x2 = sorted[0]![2]
      for (const r of sorted.slice(1)) {
        if (r[0] <= x2 + gapTol) {
          x2 = Math.max(x2, r[2])
        } else {
          out.push(markupRectQuad([x1, y1, x2, y2]))
          x1 = r[0]
          x2 = r[2]
        }
      }
      out.push(markupRectQuad([x1, y1, x2, y2]))
    } else {
      const gapTol = Math.max(2, (lx2 - lx1) * 0.35)
      const sorted = [...items].sort((a, b) => a[1] - b[1])
      let a = sorted[0]![1]
      let b = sorted[0]![3]
      for (const r of sorted.slice(1)) {
        if (r[1] <= b + gapTol) {
          b = Math.max(b, r[3])
        } else {
          out.push(markupRectQuad([lx1, a, lx2, b]))
          a = r[1]
          b = r[3]
        }
      }
      out.push(markupRectQuad([lx1, a, lx2, b]))
    }
  }
  return out
}
