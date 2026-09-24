/**
 * When the layout engine measured with HeuristicMetrics / a missing-font
 * substitute, Konva paints with the browser's real face (PingFang SC, YaHei…).
 * Those faces are wider → glyphs overlap, and the edit overlay's fixed
 * fragment widths wrap at hyphens (Multi- / Agent) so double-click looks
 * shifted. Re-measure each fragment with the same font stack the canvas uses
 * and re-pack / re-wrap to the box width.
 */
import type { GlyphRun, RenderTextLayout, ShapeRenderNode, TextLine } from '@genoffice/pptx-render'
import { displayFontFamily } from './display-font'

const reflowed = new WeakSet<RenderTextLayout>()

function fontCss(run: GlyphRun): string {
  const family = run.fontFamily || 'sans-serif'
  const bold = !!run.bold && !/lucida (sans|grande)/i.test(family)
  // Same stack Konva / the edit overlay use. A generic PingFang fallback here used
  // to measure KaiTi/Calibri as a different face, so click-in wrapping disagreed
  // with the canvas.
  return `${run.italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${run.fontSizePx}px ${displayFontFamily(family)}`
}

function measureCtx(): CanvasRenderingContext2D | null {
  if (typeof document === 'undefined') return null
  return document.createElement('canvas').getContext('2d')
}

export function canvasCanMeasure(): boolean {
  const ctx = measureCtx()
  if (!ctx) return false
  ctx.font = '20px sans-serif'
  return ctx.measureText('MM').width > 1
}

export function measureGlyphRun(run: GlyphRun, ctx: CanvasRenderingContext2D): number {
  if (!run.text) return run.widthPx
  ctx.font = fontCss(run)
  let w = ctx.measureText(run.text).width
  if (run.letterSpacingPx) w += run.letterSpacingPx * run.text.length
  return w
}

function groupParagraphs(lines: TextLine[]): TextLine[][] {
  const paras: TextLine[][] = []
  for (const line of lines) {
    if (line.paraStart === false && paras.length) paras[paras.length - 1]!.push(line)
    else paras.push([line])
  }
  return paras
}

function alignOrigin(align: TextLine['align'], innerW: number, contentW: number, marL: number): number {
  if (align === 'center') return marL + Math.max(0, (innerW - marL - contentW) / 2)
  if (align === 'right') return marL + Math.max(0, innerW - marL - contentW)
  return marL
}

function packLine(
  template: TextLine,
  body: GlyphRun[],
  bullets: GlyphRun[],
  top: number,
  paraStart: boolean,
  innerW: number,
): TextLine {
  const marL = template.marLPx ?? 0
  const contentW = body.reduce((s, r) => s + r.widthPx, 0)
  let x = alignOrigin(template.align, innerW, contentW, marL)
  const runs = [...bullets, ...body].map((r) => {
    if (r.isBullet) return { ...r, baselineY: top + r.baselineY }
    const next = { ...r, x, baselineY: top + r.baselineY }
    x += r.widthPx
    return next
  })
  return {
    ...template,
    top,
    runs,
    paraStart,
    trailingSpace: undefined,
  }
}

/** Mutates `text` once so canvas paint and the edit overlay share advances. */
export function reflowTextLayoutToCanvas(
  text: RenderTextLayout,
  innerW: number,
  measure?: (run: GlyphRun) => number,
): boolean {
  if (reflowed.has(text) || text.vert || innerW < 8) return false
  const ctx = measure ? null : measureCtx()
  if (!measure && (!ctx || !canvasCanMeasure())) return false
  const meas = measure ?? ((run: GlyphRun) => measureGlyphRun(run, ctx!))

  const wrap = text.wrap !== false
  const newLines: TextLine[] = []
  let maxDrift = 0
  for (const para of groupParagraphs(text.lines)) {
    const template = para[0]!
    const bullets: GlyphRun[] = []
    const body: GlyphRun[] = []
    for (const ln of para) {
      for (const r of ln.runs) {
        const w = meas(r)
        if (r.widthPx > 1) maxDrift = Math.max(maxDrift, Math.abs(w - r.widthPx) / r.widthPx)
        // baselineY stored relative to the source line top so new lines can rebase it
        const run = { ...r, widthPx: w, baselineY: r.baselineY - ln.top }
        if (r.isBullet) bullets.push(run)
        else body.push(run)
      }
    }
    const avail = Math.max(innerW - (template.marLPx ?? 0), 1)
    const packed: GlyphRun[][] = []
    let cur: GlyphRun[] = []
    let curW = 0
    for (const run of body) {
      const space = !run.text.trim()
      if (wrap && cur.length && !space && curW + run.widthPx > avail) {
        packed.push(cur)
        cur = []
        curW = 0
      }
      cur.push(run)
      curW += run.widthPx
    }
    if (cur.length || !packed.length) packed.push(cur)

    let top = template.top
    packed.forEach((runs, i) => {
      newLines.push(
        packLine(template, runs, i === 0 ? bullets : [], top, i === 0 ? template.paraStart !== false : false, innerW),
      )
      top += template.height
    })
  }

  reflowed.add(text)
  // Real faces (PingFang/YaHei) drift well above 3%; skip when layout already matched paint.
  if (!newLines.length || maxDrift < 0.03) return false
  text.lines = newLines
  const last = newLines[newLines.length - 1]!
  text.contentHeight = last.top + last.height
  const lastDescent = Math.max(0, ...last.runs.map((r) => (r.ascentPx ?? r.fontSizePx * 0.2) * 0.25))
  text.inkBottom = last.top + last.height + lastDescent
  reflowed.add(text)
  return true
}

export function reflowShapeTextToCanvas(node: ShapeRenderNode): void {
  const text = node.text
  if (!text) return
  const inner = Math.max(node.box.w - text.insets.l - text.insets.r, 1)
  reflowTextLayoutToCanvas(text, inner)
}
