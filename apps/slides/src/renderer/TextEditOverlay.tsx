/**
 * 3.2 DOM overlay text editing (run-level rich text) — a contentEditable stacked over the text
 * box takes over input (caret/selection/IME for free). Bold/italic/underline are triggered by the
 * ribbon font group (execCommand on the selection). On exit, walk the DOM to extract the
 * paragraph/run structure (each run's format preserved independently) and go through IPC editText.
 */
import React, { useEffect, useRef } from 'react'
import { DEFAULT_INSETS_EMU, emuToPx } from '@genoffice/pptx-render'
import type { GlyphRun, ShapeRenderNode, TextLine } from '@genoffice/pptx-render'
import type { EditParagraph, EditRun, LinkTargetOp } from '../shared/ipc'
import { decodeLinkTarget, encodeLinkTarget } from '../shared/run-link'
import { displayFontFamily, konvaBaselineDrop } from './konva-adapter'
import {
  claimParaBreak,
  claimSoftBreak,
  isEnterKey,
  isSoftLineBreakInput,
  isSoftLineBreakShortcut,
  notePhysicalShift,
  releaseParaBreak,
  releaseSoftBreak,
  resetPhysicalShift,
} from './text-edit-keys'
import { ZOOM_PREVIEW_EVENT } from './zoom-preview'
import { reflowShapeTextToCanvas } from './canvas-text-reflow'
import { FONT_SIZES } from './components/ribbon-shared'

/** Shift+Enter in-paragraph break: insert a "\n" text node (pre-wrap shows it as a line
 * break; extractParagraphs maps it to <a:br/>). Default contentEditable <br> / insertText
 * both become paragraph splits. */
function insertTextBoxSoftBreak(): void {
  if (!claimSoftBreak()) return
  const sel = window.getSelection()
  if (!sel || !sel.rangeCount) return
  const range = sel.getRangeAt(0)
  range.deleteContents()
  const tn = document.createTextNode('\n')
  range.insertNode(tn)
  range.setStartAfter(tn)
  range.collapse(true)
  sel.removeAllRanges()
  sel.addRange(range)
}

/**
 * Plain Enter must split the paragraph DIV (data-src-para). Chromium otherwise
 * splits the inner layout-line SPAN (display:block; nowrap), which looks like a
 * newline while editing but extractParagraphs only splits on DIV/P/BR — commit
 * then joins the lines back into one wrapping paragraph.
 */
function insertTextBoxParagraphBreak(): void {
  if (!claimParaBreak()) return
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const sel = window.getSelection()
  if (!sel?.rangeCount) return
  const range = sel.getRangeAt(0)
  range.deleteContents()

  let blk: HTMLElement | null =
    range.startContainer === root
      ? null
      : range.startContainer instanceof HTMLElement
        ? range.startContainer
        : range.startContainer.parentElement
  while (blk && blk.parentElement !== root) blk = blk.parentElement
  if (!(blk instanceof HTMLElement) || blk.parentElement !== root) {
    const wrap = document.createElement('div')
    while (root.firstChild) wrap.appendChild(root.firstChild)
    if (!wrap.childNodes.length) wrap.appendChild(document.createElement('br'))
    root.appendChild(wrap)
    blk = wrap
    range.selectNodeContents(blk)
    range.collapse(false)
  }

  const rest = document.createRange()
  try {
    rest.setStart(range.startContainer, range.startOffset)
    rest.setEnd(blk, blk.childNodes.length)
  } catch {
    rest.selectNodeContents(blk)
    rest.collapse(false)
  }
  const tail = rest.extractContents()
  const next = document.createElement('div')
  for (const key of Object.keys(blk.dataset)) {
    if (key === 'layoutLine' || key === 'layoutSpaced') continue
    next.dataset[key] = blk.dataset[key]!
  }
  next.appendChild(tail)
  if (!next.textContent) next.appendChild(document.createElement('br'))
  if (!blk.textContent) blk.appendChild(document.createElement('br'))
  blk.after(next)
  const caret = document.createRange()
  caret.setStart(next, 0)
  caret.collapse(true)
  sel.removeAllRanges()
  sel.addRange(caret)
  root.dispatchEvent(new Event('input', { bubbles: true }))
}

interface Props {
  node: ShapeRenderNode
  /** Viewport scale (RenderSlide.scale) — editing renders in viewport px; divided back out when committing to pt */
  scale: number
  onCommit: (paragraphs: EditParagraph[]) => void
  onCancel: () => void
  /** First real keystroke / IME insert — not merely opening the overlay. */
  onDirty?: () => void
  /** Tab/Shift+Tab (for table cell editing): commit current content and jump to the next/previous cell.
   * paragraphs=null means content unchanged (the host may skip committing and only jump). */
  onTabNav?: (paragraphs: EditParagraph[] | null, dir: 1 | -1) => void
  /** Viewport coordinates of the double-click: select the word there when entering editing; defaults to caret at end */
  caretPoint?: { x: number; y: number }
  /** Entered by typing directly on a selected shape: select all, then replace the whole content with that character */
  replaceWith?: string
  /** ⌘/Ctrl+click on a linked run follows the link (slide jump / external url) */
  onFollowLink?: (target: LinkTargetOp) => void
  /** Edit-frame color (matches the canvas selection chrome: white on dark slide backgrounds) */
  frameColor?: string
  /** Canvas CSS zoom: the outline divides by it to keep a constant on-screen weight */
  zoom?: number
}

/** PowerPoint "single" spacing is 1.2em (not CSS 1.0). Matches text-layout.ts PPT_SINGLE. */
const PPT_SINGLE_LINE = 1.2

/** First-strong-character inference over a paragraph's logical text (mirrors what dir="auto" does). */
function inferParaRtl(paraLines: TextLine[]): boolean {
  const runs = paraLines
    .flatMap((l) => l.runs)
    .filter((r) => !r.isBullet)
    .sort((a, b) => (a.logicalOrder ?? 0) - (b.logicalOrder ?? 0))
  for (const r of runs) {
    for (const ch of r.text) {
      if (/[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/.test(ch)) return true
      // eslint-disable-next-line no-misleading-character-class -- broad strong-LTR ranges; combining marks inside are irrelevant for a per-char strong-direction probe
      if (/[A-Za-z\u00c0-\u058f\u0900-\ud7ff\uf900-\ufdcf]/.test(ch)) return false
    }
  }
  return false
}

/** Layout lines → paragraph grouping (paraStart marks wrap boundaries; missing means an independent paragraph, backward compatible). */
function groupLinesToParagraphs(lines: TextLine[]): TextLine[][] {
  const paras: TextLine[][] = []
  for (const line of lines) {
    if (line.paraStart === false && paras.length) paras[paras.length - 1]!.push(line)
    else paras.push([line])
  }
  return paras
}

/** Browser inline-box metrics of a font (Chromium: an inline text box is exactly
 * ascent+descent tall, and a zero-size inline-block probe's offsetTop is the baseline).
 * Cached per font. Returns zero heights in layout-less environments (jsdom) — callers
 * skip compensation there. */
const fontBoxCache = new Map<string, { height: number; ascent: number }>()
function browserFontBox(
  family: string,
  sizePx: number,
  bold?: boolean,
  italic?: boolean,
): { height: number; ascent: number } {
  const key = `${family}|${Math.round(sizePx * 10)}|${bold ? 'b' : ''}${italic ? 'i' : ''}`
  const hit = fontBoxCache.get(key)
  if (hit) return hit
  const host = document.createElement('div')
  host.style.cssText =
    'position:absolute;left:-9999px;top:0;visibility:hidden;line-height:normal;white-space:pre'
  host.style.fontFamily = family
  host.style.fontSize = `${sizePx}px`
  host.style.fontWeight = bold ? 'bold' : 'normal'
  host.style.fontStyle = italic ? 'italic' : 'normal'
  const text = document.createElement('span')
  text.textContent = 'Hg'
  const probe = document.createElement('span')
  probe.style.cssText = 'display:inline-block;width:0;height:0'
  host.append(text, probe)
  document.body.appendChild(host)
  const hostTop = host.getBoundingClientRect().top
  const m = {
    height: text.getBoundingClientRect().height,
    // zero-size inline-block: its box top sits exactly on the baseline (fractional,
    // unlike offsetTop which rounds to whole px)
    ascent: probe.getBoundingClientRect().top - hostTop,
  }
  host.remove()
  fontBoxCache.set(key, m)
  return m
}

/**
 * Layout lines → the editor's initial DOM: one <div> per model paragraph (data-src-para records
 * the source paragraph index, with paragraph alignment); wrap/word-split fragments merged back
 * into their runs by srcRunIdx (data-src-run), one span per model run; bullet glyphs injected by
 * layout are skipped (not body text; committing them would turn them into text).
 * Font sizes use layout viewport px directly (including scale/autofit fontScale); on commit
 * they're divided by norm back to pt.
 * Line/paragraph spacing is derived back from layout lines (lnSpc/lineExact/lnSpcReduction baked
 * into height, spcBef/spcAft baked into gaps in line tops), so editing's vertical metrics match the canvas.
 * anchorDy = the whole offset that middle/bottom anchoring bakes into line tops (editing implements
 * anchoring with flex, so it must be removed from the first paragraph's top or the offset doubles).
 * vertical = bodyPr vert editing (the contentEditable is in writing-mode: vertical-rl): layout
 * "lines" are columns whose top/height/advance are column metrics, so every horizontal-flow
 * baking (line-height, paragraph gaps, baseline/alignment compensation, fragment advances) is
 * skipped and the browser lays the text out vertically itself.
 * Exported so tests can do "layout → DOM → extractParagraphs" round-trip assertions.
 */
export function populateEditorDom(
  div: HTMLElement,
  lines: TextLine[],
  anchorDy = 0,
  innerW?: number,
  vertical = false,
): void {
  div.innerHTML = ''
  delete div.dataset.layoutReleased
  // Root/strut font: the paragraph divs inherit it and it participates in every line box
  const rootCs = div.isConnected ? window.getComputedStyle(div) : null
  const strutFont = rootCs ? { family: rootCs.fontFamily, size: parseFloat(rootCs.fontSize) } : null
  // Widest laid-out line: in a nowrap box the block grows to max-content, so centered/right
  // paragraphs align within this width instead of the box (the canvas splits overflow to
  // both sides) — compensated per paragraph below. Bullet glyphs are skipped: they never
  // enter the editor DOM and the engine's alignment width excludes them too.
  const maxLineW = Math.max(
    0,
    ...lines.map((ln) => ln.runs.reduce((acc, r) => acc + (r.isBullet ? 0 : r.widthPx), 0)),
  )
  let prevEnd = anchorDy
  let prevDyFix = 0
  groupLinesToParagraphs(lines).forEach((paraLines, pi) => {
    const p = document.createElement('div')
    p.dataset.srcPara = String(pi)
    const first = paraLines[0]!
    const last = paraLines[paraLines.length - 1]!
    // ── Glyph-position fidelity vs the canvas renderer ──
    // Vertical: the canvas draws the dominant run's baseline at
    // lineTop + engineAscent + konvaBaselineDrop (0 for resolved faces, the fallback
    // font's offset from the legacy 0.8em rule otherwise — see the adapter); CSS puts
    // the DOM baseline at half-leading + browser ascent. The difference is several px
    // on fallback-drawn CJK/serif or lnSpc ≠ 100% text and reads as the text jumping
    // when editing starts. Measure both sides and cancel the difference with a
    // relative offset (flow is unaffected).
    let engineAscent = 0
    let domBaseline = 0
    let dominant: GlyphRun | null = null
    for (const r of first.runs) {
      const a = r.ascentPx ?? r.fontSizePx * 0.8
      if (a > engineAscent) {
        engineAscent = a
        dominant = r
      }
    }
    const drop = dominant
      ? konvaBaselineDrop(
          displayFontFamily(dominant.fontFamily ?? ''),
          dominant.fontSizePx,
          dominant.bold,
          dominant.italic,
        )
      : 0
    // lnSpc>100%: the canvas pins glyphs to the slot bottom (leadAbove below the line top)
    const canvasBaseline = (first.leadAbove ?? 0) + engineAscent + drop
    const participants: Array<{ family: string; size: number; bold?: boolean; italic?: boolean }> =
      first.runs.map((r) => ({
        family: displayFontFamily(r.fontFamily ?? ''),
        size: r.fontSizePx,
        bold: r.bold,
        italic: r.italic,
      }))
    if (strutFont) participants.push({ family: strutFont.family, size: strutFont.size })
    // Half-leading distributes over the CSS line-height, which is the advance
    // (box + external leading) when the font has an hhea lineGap
    const cssLineH = first.advance ?? first.height
    for (const f of participants) {
      if (!f.family || !f.size) continue
      const m = browserFontBox(f.family, f.size, f.bold, f.italic)
      if (!m.height) continue
      domBaseline = Math.max(domBaseline, (cssLineH - m.height) / 2 + m.ascent)
    }
    const dyFix =
      !vertical && canvasBaseline > 0 && domBaseline > 0 ? canvasBaseline - domBaseline : 0
    if (!vertical) {
      p.style.lineHeight = `${cssLineH}px`
      // Kill UA <p>/<div> margins; only layout-derived spacing should remain.
      p.style.marginBottom = '0'
      p.style.padding = '0'
      // So Enter-clones recompute THIS block's PPT line box from its own font
      // (a heading's px line-height must not stick on a body paragraph).
      if (dominant?.fontSizePx) {
        const pct = (cssLineH / (PPT_SINGLE_LINE * dominant.fontSizePx)) * 100
        // Only stamp non-default spacing so extract doesn't invent a format patch
        if (Number.isFinite(pct) && Math.abs(pct - 100) > 0.5)
          p.dataset.lineSpacingPct = String(Math.round(pct * 10) / 10)
      }
      const gap = first.top - prevEnd
      // Each paragraph's baseline correction uses position:relative;top. Different
      // fonts (e.g. bold heading vs body) get different dyFix values, which would
      // silently shrink/expand the *visual* gap between paragraphs even when
      // marginTop matches the layout spacing. Compensate so consecutive baselines
      // keep the canvas gap: visualGap = marginTop + dyFix - prevDyFix.
      const marginTop = gap - dyFix + prevDyFix
      if (Math.abs(marginTop) > 0.01) p.style.marginTop = `${marginTop}px`
      else p.style.marginTop = '0'
      p.dataset.layoutSpaced = '1'
      // Model spacing (not the first-line applied gap): Enter clones these so a
      // split of paragraph 0 still previews spaceBefore+spaceAfter like the canvas.
      if (first.spcBefPx) p.dataset.spcBefPx = String(first.spcBefPx)
      if (first.spcAftPx) p.dataset.spcAftPx = String(first.spcAftPx)
    }
    // The DOM block ends one advance below the last line top (external leading renders
    // inside the block, unlike the canvas) — margins of following paragraphs compensate
    prevEnd = last.top + (last.advance ?? last.height)
    prevDyFix = dyFix
    // RTL paragraphs (Arabic/Hebrew) align in editing as on canvas: the browser sets direction
    // by the first strong character. An explicit a:pPr rtl can disagree with that inference
    // (TextLine.rtl carries the effective base) — then the browser needs an explicit dir
    const effRtl = first.rtl === true
    p.dir = effRtl === inferParaRtl(paraLines) ? 'auto' : effRtl ? 'rtl' : 'ltr'
    const align = paraLines[0]?.align
    if (align) p.style.textAlign = align
    // Body text starts at marL, exactly like the canvas (lists/indent used to snap to the
    // inset edge on entering edit); first-line indent applies only without a bullet
    const marL = (!vertical && first.marLPx) || 0
    if (marL) p.style.marginLeft = `${marL}px`
    const indentPx = (!vertical && first.indentPx) || 0
    if (indentPx && !first.runs.some((r) => r.isBullet)) p.style.textIndent = `${indentPx}px`
    // Horizontal: nowrap overflow — the canvas centers/right-aligns within the box and
    // spills both ways; the DOM block is max-content wide and anchored at the box's left.
    // The difference is a constant per box (zero when the content fits or the box wraps).
    let dxFix = 0
    if (innerW != null && innerW > 0 && (align === 'center' || align === 'right')) {
      const over = Math.max(0, maxLineW - innerW)
      dxFix = align === 'center' ? -over / 2 : -over
    }
    if (Math.abs(dyFix) > 0.1 || Math.abs(dxFix) > 0.1) {
      p.style.position = 'relative'
      if (Math.abs(dyFix) > 0.1) p.style.top = `${dyFix}px`
      if (Math.abs(dxFix) > 0.1) p.style.left = `${dxFix}px`
    }
    const level = paraLines[0]?.level ?? 0
    if (level) {
      p.dataset.level = String(level)
      // Visual indent hint only when the real marL isn't known (the canvas lays out by marL)
      if (!marL && !vertical) p.style.marginLeft = `${level * 24}px`
    }
    // Original bullet kind, for the ribbon's toggle-off semantics while editing
    const bulletRun = first.runs.find((r) => r.isBullet)
    if (bulletRun) p.dataset.hadBullet = /^\d/.test(bulletRun.text) ? 'number' : 'char'
    // Horizontal: one nowrap row per canvas line so click-in cannot re-wrap
    // (caret / subpixel / slightly tighter overlay width used to push the last
    // Latin grapheme onto the next line — "dddd" → "ddd"+"d"). Vertical editing
    // keeps a single flow: the engine's advances are horizontal-only.
    const lineHosts: HTMLElement[] = []
    if (vertical) {
      lineHosts.push(p)
    } else {
      paraLines.forEach(() => {
        const lineEl = document.createElement('span')
        lineEl.dataset.layoutLine = 'true'
        lineEl.style.display = 'block'
        lineEl.style.whiteSpace = 'nowrap'
        p.appendChild(lineEl)
        lineHosts.push(lineEl)
      })
      if (!lineHosts.length) lineHosts.push(p)
    }
    paraLines.forEach((line, li) => {
      const host = lineHosts[Math.min(li, lineHosts.length - 1)]!
      if (li > 0 && paraLines[li - 1]!.trailingSpace) {
        const prevHost = lineHosts[Math.min(li - 1, lineHosts.length - 1)]!
        const frags = prevHost.querySelectorAll('[data-layout-fragment]')
        const lastFrag = frags[frags.length - 1] as HTMLElement | undefined
        if (lastFrag) lastFrag.textContent = (lastFrag.textContent ?? '') + ' '
      }
      const logicalRuns = [...line.runs].sort(
        (a, b) =>
          (a.logicalOrder ?? Number.MAX_SAFE_INTEGER) - (b.logicalOrder ?? Number.MAX_SAFE_INTEGER),
      )
      for (const run of logicalRuns) {
        if (run.isBullet || run.text === '') continue
        appendEditorRun(host, run, run.srcRunIdx, run.text, vertical)
      }
      if (line.softBreakAfter != null) {
        const span = document.createElement('span')
        span.textContent = '\n'
        span.dataset.srcRun = String(line.softBreakAfter)
        host.appendChild(span)
      }
    })
    if (!p.textContent) p.appendChild(document.createElement('br')) // Empty paragraph placeholder
    div.appendChild(p)
  })
}

/** Append one canvas glyph run into the editor (run container + layout fragment). */
function appendEditorRun(
  host: HTMLElement,
  run: GlyphRun,
  srcRun: number | undefined,
  text: string,
  vertical: boolean,
): void {
  const prev = host.lastElementChild as HTMLElement | null
  const src = srcRun != null ? String(srcRun) : undefined
  const reuse = prev?.dataset.runContainer === 'true' && prev.dataset.srcRun === src
  // Linked runs become <a href> so execCommand('unlink') and extraction see them natively
  const span = reuse ? prev! : document.createElement(run.link ? 'a' : 'span')
  if (!reuse) {
    span.dataset.runContainer = 'true'
    if (src) span.dataset.srcRun = src
    if (run.link) span.setAttribute('href', run.link)
    if (run.bold) span.style.fontWeight = 'bold'
    if (run.italic) span.style.fontStyle = 'italic'
    const deco = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(
      Boolean,
    )
    if (deco.length) span.style.textDecoration = deco.join(' ')
    else if (run.link) span.style.textDecoration = 'none' // suppress the UA <a> underline
    // Super/subscript: the initial DOM must restore it (otherwise extraction sends explicit 0/false and wipes the original format)
    if (run.baselinePct) span.style.verticalAlign = run.baselinePct > 0 ? 'super' : 'sub'
    span.style.fontSize = `${run.fontSizePx}px`
    if (run.fontFamily) {
      const display = displayFontFamily(run.fontFamily)
      span.style.fontFamily = display
      // Record what was baked in for display: on extraction, if the first item of the stack is
      // still this, the user didn't change the font — commit the model's original name (data-font),
      // or nothing at all when the model run has no explicit font (run.fontFamily is then a layout
      // default / missing-font substitution like Arial, which must never be written into the file).
      span.dataset.displayFont = firstFontFamily(display)
      if (run.srcFontFamily) span.dataset.font = run.srcFontFamily
    }
    span.style.color = normalizeCss(run.color)
    // Text highlight: display-only (extraction never reads it back; the patch path keeps <a:highlight>)
    if (run.highlight) span.style.backgroundColor = normalizeCss(run.highlight)
    host.appendChild(span)
  }
  const fragment = document.createElement('span')
  fragment.dataset.layoutFragment = 'true'
  fragment.textContent = text
  // Each fragment occupies the exact advance measured by the layout engine. CJK is normally
  // one grapheme per fragment; Latin/SEA keep their script-aware token boundaries. RTL stays
  // in normal inline flow so Chromium can preserve joining and bidirectional shaping.
  // Vertical editing skips fixed advances entirely: the engine's widthPx is a horizontal
  // measure, and inline-block cells would break writing-mode glyph orientation
  if (!run.rtl && !vertical) {
    fragment.style.display = 'inline-block'
    fragment.style.width = `${run.widthPx}px`
    // Keep the fragment as one unit: CSS otherwise wraps at '-' inside a
    // too-narrow box (English Shift+Enter is unrelated; this is Multi-Agent).
    fragment.style.whiteSpace = 'nowrap'
    fragment.style.verticalAlign = 'top'
  }
  // Keep the browser editor visually aligned with the canvas renderer. This is display-only:
  // extraction intentionally preserves the source run's PPT letter spacing through srcRun.
  if (run.letterSpacingPx) fragment.style.letterSpacing = `${run.letterSpacingPx}px`
  span.appendChild(fragment)
}

/** Fixed fragment advances align the untouched editor with canvas, but become stale after typing.
 * Release them so inserted/deleted text can reflow naturally. */
export function releaseEditorLayoutConstraints(root: HTMLElement): void {
  if (root.dataset.layoutReleased === 'true') return
  root.dataset.layoutReleased = 'true'
  root.querySelectorAll<HTMLElement>('[data-layout-fragment]').forEach((fragment) => {
    fragment.style.display = ''
    fragment.style.width = ''
    fragment.style.whiteSpace = ''
    fragment.style.verticalAlign = ''
  })
  root.querySelectorAll<HTMLElement>('[data-layout-line]').forEach(releaseLayoutLine)
}

function releaseLayoutLine(el: HTMLElement): void {
  if (el.dataset.layoutLine !== 'true') return
  el.style.display = ''
  el.style.whiteSpace = ''
}

/** Font size the canvas will use for this block's line box (first explicit run, else inherited). */
function primaryFontSizePx(el: HTMLElement, fallback: number): number {
  const styled = el.querySelector<HTMLElement>('[style*="font-size"]')
  const n = parseFloat(styled?.style.fontSize || el.style.fontSize || '')
  if (Number.isFinite(n) && n > 0) return n
  const cs = parseFloat(window.getComputedStyle(el).fontSize)
  return Number.isFinite(cs) && cs > 0 ? cs : fallback
}

function datasetNum(el: HTMLElement, key: string): number {
  const n = parseFloat(el.dataset[key] ?? '')
  return Number.isFinite(n) ? n : 0
}

/** Model space-before/after in overlay px (layout-stamped px wins; ribbon pt × viewport scale). */
function paragraphSpacingPx(el: HTMLElement, side: 'before' | 'after'): number {
  const fromLayout = datasetNum(el, side === 'before' ? 'spcBefPx' : 'spcAftPx')
  if (fromLayout) return fromLayout
  const pt = datasetNum(el, side === 'before' ? 'spaceBeforePt' : 'spaceAfterPt')
  if (!pt) return 0
  const norm = parseFloat(el.parentElement?.dataset.norm ?? '1') || 1
  return ((pt * 96) / 72) * norm
}

/**
 * After Enter, Chromium clones the previous block — including the heading's
 * layout-seeded line-height and the first paragraph's marginTop 0 (PowerPoint
 * ignores space-before at the frame top). Clones must use THIS block's font size
 * for PPT single spacing, and the same spaceBefore+spaceAfter gap the canvas
 * will paint on commit (applyEditParagraphs copies those fields onto both halves).
 */
export function normalizeTypedParagraphBoxes(root: HTMLElement, baseFontSizePx: number): void {
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement =>
      el instanceof HTMLElement && (el.tagName === 'DIV' || el.tagName === 'P'),
  )
  blocks.forEach((el, i) => {
    el.style.marginBottom = '0'
    el.style.paddingTop = '0'
    el.style.paddingBottom = '0'
    const sizePx = primaryFontSizePx(el, baseFontSizePx)
    const pct = datasetNum(el, 'lineSpacingPct')
    const spacingPct = pct > 0 ? pct : 100
    el.style.lineHeight = `${(PPT_SINGLE_LINE * sizePx * spacingPct) / 100}px`
    // Baseline compensation was computed for the original layout line box; after
    // typing/cloning it fights the font-derived line-height and reads as extra gap.
    if (el.style.position === 'relative') el.style.top = ''
    if (i === 0) {
      el.style.marginTop = '0'
      return
    }
    const prev = blocks[i - 1]!
    const gap = paragraphSpacingPx(prev, 'after') + paragraphSpacingPx(el, 'before')
    el.style.marginTop = gap ? `${gap}px` : '0'
  })
}

/** Release one fragment's fixed advance (stale once its text or format changes). */
function releaseFragment(el: Element | null | undefined): void {
  if (!(el instanceof HTMLElement) || el.dataset.layoutFragment !== 'true') return
  el.style.display = ''
  el.style.width = ''
  el.style.whiteSpace = ''
  el.style.verticalAlign = ''
  const line = el.closest('[data-layout-line]')
  if (line instanceof HTMLElement) releaseLayoutLine(line)
}

/** Nearest enclosing layout fragment of a DOM point (bounded by the editor root). */
function fragmentAround(node: Node | null | undefined, root: HTMLElement): HTMLElement | null {
  let el = node instanceof HTMLElement ? node : (node?.parentElement ?? null)
  while (el && el !== root) {
    if (el.dataset.layoutFragment === 'true') return el
    el = el.parentElement
  }
  return null
}

/** Release only the fragments an edit touches (the given target ranges plus the current
 * selection; a collapsed caret also frees its neighbor fragments — Backspace/Delete at a
 * fragment edge mutates them). Untouched fragments keep their canvas-measured advances, so
 * the rest of the text stays put: releasing everything on the first keystroke re-measured
 * and re-wrapped the whole box with browser rules (natural advances + CJK line-break
 * prohibitions the canvas engine doesn't apply) and made all the text visibly jump.
 * Exported for tests. */
export function releaseFragmentsAtEdit(
  root: HTMLElement,
  targetRanges: readonly AbstractRange[] = [],
): void {
  const ranges: AbstractRange[] = [...targetRanges]
  const sel = window.getSelection()
  if (sel?.rangeCount && root.contains(sel.anchorNode)) ranges.push(sel.getRangeAt(0))
  if (!ranges.length) return
  const frags = [...root.querySelectorAll<HTMLElement>('[data-layout-fragment]')]
  const releaseLineAround = (node: Node | null) => {
    let el = node instanceof HTMLElement ? node : node?.parentElement
    while (el && el !== root) {
      if (el.dataset.layoutLine === 'true') {
        releaseLayoutLine(el)
        return
      }
      el = el.parentElement
    }
  }
  for (const r of ranges) {
    if (r.collapsed) {
      const f = fragmentAround(r.startContainer, root)
      if (f) {
        const i = frags.indexOf(f)
        releaseFragment(f)
        releaseFragment(frags[i - 1])
        releaseFragment(frags[i + 1])
      } else {
        releaseLineAround(r.startContainer)
      }
      continue
    }
    // Ranged edit (selection replace/delete, execCommand format): free every intersecting fragment
    const live =
      r instanceof Range
        ? r
        : (() => {
            const x = document.createRange()
            x.setStart(r.startContainer, r.startOffset)
            x.setEnd(r.endContainer, r.endOffset)
            return x
          })()
    for (const f of frags) if (live.intersectsNode(f)) releaseFragment(f)
    releaseLineAround(r.startContainer)
  }
}

export function TextEditOverlay({
  node,
  scale,
  onCommit,
  onCancel,
  onDirty,
  onTabNav,
  caretPoint,
  replaceWith,
  onFollowLink,
  frameColor = '#232425',
  zoom = 1,
}: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const frameRef = useRef<HTMLDivElement>(null)

  // The edit frame lives inside the CSS-scaled stage: during a zoom gesture only the
  // transform advances, so the zoom-compensated outline would thicken and then snap at
  // commit. Counter-scale it per previewed frame, same as the selection chrome.
  useEffect(() => {
    const onPreview = (e: Event) => {
      const z = (e as CustomEvent<number>).detail
      const el = frameRef.current
      if (typeof z !== 'number' || !el) return
      el.style.outlineWidth = `${2 / (globalThis.devicePixelRatio || 1) / Math.max(z, 0.1)}px`
    }
    window.addEventListener(ZOOM_PREVIEW_EVENT, onPreview)
    return () => window.removeEventListener(ZOOM_PREVIEW_EVENT, onPreview)
  }, [])

  const box = node.box
  // A shape with no text body yet: preview the body setText is about to create so the
  // content does not jump between typing and commit. Every fresh body gets the standard
  // bodyPr insets (createTextBody writes them unconditionally); only an autoshape also
  // gets the centered anchor/alignment — keep both halves in step with setText.
  const freshBody = !node.text && node.type === 'shape'
  const fresh = freshBody && !node.placeholder && !node.txBox
  const insets =
    node.text?.insets ??
    (freshBody
      ? {
          l: emuToPx(DEFAULT_INSETS_EMU.l, scale),
          t: emuToPx(DEFAULT_INSETS_EMU.t, scale),
          r: emuToPx(DEFAULT_INSETS_EMU.r, scale),
          b: emuToPx(DEFAULT_INSETS_EMU.b, scale),
        }
      : { l: 0, t: 0, r: 0, b: 0 })
  const anchor = node.text?.anchor ?? (fresh ? 'middle' : 'top')
  // Editing uses layout viewport px directly (including viewport scale and autofit fontScale): layout height =
  // visual height, so the edit box isn't inflated; on commit extractParagraphs divides by norm back to model pt
  const norm = scale * (node.text?.fontScale ?? 1) || 1
  const wrap = node.text?.wrap !== false
  // bodyPr vert: edit in a CSS vertical writing mode matching the canvas engine —
  // eaVert: vertical-rl/mixed (CJK upright, Latin rotated, columns right→left);
  // wordArtVert: vertical-lr/upright (every glyph upright, stacked, columns left→right);
  // vert: vertical-rl/sideways (whole block rotated 90° cw, CJK included);
  // vert270: sideways-lr (whole block rotated 90° ccw, lines flow left→right)
  const vertMode = node.text?.vert
  const vertText = !!vertMode
  // Modes whose line stacking runs left→right (the frame flexes as a plain row there)
  const vertLtr = vertMode === 'vert270' || vertMode === 'wordArtVert'
  // WPS/PPT Insert>Text Box: width stays put (wrap), height follows content.
  const growVert =
    !vertText &&
    !node.placeholder &&
    node.text?.autofit !== 'shrink' &&
    (node.text?.autofit === 'resize' || !!node.txBox)
  const firstRun =
    node.text?.lines[0]?.runs.find((r) => !r.isBullet) ?? node.text?.lines[0]?.runs[0]
  // Fallback = the layout engine's 18pt default in px, so typing into an empty body
  // matches the canvas line height (18*norm would be px-as-pt: 13.5pt)
  const baseFontSize = firstRun?.fontSizePx ?? ((18 * 96) / 72) * norm
  const baseColor = normalizeCss(firstRun?.color ?? '#000')
  const baseFont = displayFontFamily(firstRun?.fontFamily ?? 'Calibri')

  // Initial content: see populateEditorDom; snapshot the initial extraction, unchanged commits go through onCancel
  // (no empty undo step / no dirty flag — opening to look and clicking away isn't an edit)
  const initialRef = useRef<string>('')
  useEffect(() => {
    const div = ref.current
    if (!div) return
    // Prefer <div> over <p> so UA paragraph margins never invent canvas-invisible gaps
    try {
      document.execCommand('defaultParagraphSeparator', false, 'div')
    } catch {
      /* ignore */
    }
    div.dataset.norm = String(norm) // The ribbon helpers for font size increase/decrease/set take the conversion factor from here
    // Same anchor offset as the engine (the dy text-layout bakes into line tops), removed back
    // during populate — the engine anchors against the glyph extent (inkBottom), so mirror it
    const extraH =
      Math.max(box.h - insets.t - insets.b, 1) -
      (node.text?.inkBottom ?? node.text?.contentHeight ?? 0)
    const anchorDy = anchor === 'middle' ? extraH / 2 : anchor === 'bottom' ? extraH : 0
    // Vertical: anchoring/overflow compensation are horizontal-flow corrections — the
    // row-reverse frame flex implements the (right-edge-anchored) flow instead
    // Same font stack as Konva paint: heuristic layout widths overlap CJK/Latin and the
    // overlay would wrap at hyphens (Multi- / Agent) unless we reflow first.
    reflowShapeTextToCanvas(node)
    populateEditorDom(
      div,
      node.text?.lines ?? [],
      vertText ? 0 : anchorDy,
      vertText ? undefined : box.w - insets.l - insets.r,
      vertText,
    )
    initialRef.current = JSON.stringify(extractParagraphs(div, norm))
    div.focus()
    const sel = window.getSelection()
    if (sel && replaceWith) {
      // Type-to-replace: select all, then replace the whole content with the first typed character
      const range = document.createRange()
      range.selectNodeContents(div)
      sel.removeAllRanges()
      sel.addRange(range)
      document.execCommand('insertText', false, replaceWith)
      return
    }
    if (sel) {
      // Entering by double-click: select the word at the click; without coordinates/no hit, caret to end
      const hit = caretPoint ? document.caretRangeFromPoint(caretPoint.x, caretPoint.y) : null
      if (hit && div.contains(hit.startContainer)) {
        sel.removeAllRanges()
        sel.addRange(hit)
        const s = sel as Selection & {
          modify?: (alter: string, dir: string, granularity: string) => void
        }
        s.modify?.('move', 'backward', 'word')
        s.modify?.('extend', 'forward', 'word')
      } else {
        const range = document.createRange()
        range.selectNodeContents(div)
        range.collapse(false)
        sel.removeAllRanges()
        sel.addRange(range)
      }
    }
  }, [node, norm, caretPoint, replaceWith])

  const commit = () => {
    savedSel = null
    const div = ref.current
    if (!div) return onCancel()
    const paras = extractParagraphs(div, norm)
    if (JSON.stringify(paras) === initialRef.current) return onCancel()
    onCommit(paras)
  }

  // Targeted layout release (native listeners: React's onBeforeInput synthetic event does
  // not map to the real `beforeinput`). beforeinput sees the pre-mutation target ranges
  // (Backspace at a fragment edge deletes into the neighbor); the input listener covers
  // execCommand formatting (bold/font — width-changing, fires input without beforeinput).
  useEffect(() => {
    const div = ref.current
    if (!div) return
    const onBeforeInput = (ev: InputEvent) => {
      releaseFragmentsAtEdit(div, ev.getTargetRanges?.() ?? [])
      // IME-safe backup for Shift+Enter: Chinese IMEs may skip our keydown match
      // (key="Process") or mis-label the insert as a paragraph while Shift is down.
      if (!ev.isComposing && isSoftLineBreakInput(ev.inputType)) {
        ev.preventDefault()
        insertTextBoxSoftBreak()
      } else if (!ev.isComposing && ev.inputType === 'insertParagraph') {
        ev.preventDefault()
        insertTextBoxParagraphBreak()
      }
    }
    const onInput = () => {
      releaseFragmentsAtEdit(div)
      // Enter clones the previous block's marginTop/lineHeight. Clear UA margins and
      // give new blocks the same PPT single-spacing the canvas will use on commit.
      normalizeTypedParagraphBoxes(div, baseFontSize)
      onDirty?.()
      // Unlock after the same-turn keydown + beforeinput pair has been seen.
      queueMicrotask(() => {
        releaseSoftBreak()
        releaseParaBreak()
      })
    }
    div.addEventListener('beforeinput', onBeforeInput)
    div.addEventListener('input', onInput)
    return () => {
      div.removeEventListener('beforeinput', onBeforeInput)
      div.removeEventListener('input', onInput)
    }
  }, [baseFontSize, onDirty])

  // Physical Shift latch: IME 中/英 toggle often clears event.shiftKey on the
  // following Enter. Track ShiftLeft/ShiftRight at capture so both layouts work.
  useEffect(() => {
    const onKey = (ev: KeyboardEvent) => {
      notePhysicalShift(ev)
      if (ev.type === 'keyup' && isEnterKey(ev)) {
        releaseSoftBreak()
        releaseParaBreak()
      }
    }
    const onBlur = () => {
      resetPhysicalShift()
      releaseSoftBreak()
      releaseParaBreak()
    }
    window.addEventListener('keydown', onKey, true)
    window.addEventListener('keyup', onKey, true)
    window.addEventListener('blur', onBlur)
    return () => {
      window.removeEventListener('keydown', onKey, true)
      window.removeEventListener('keyup', onKey, true)
      window.removeEventListener('blur', onBlur)
      resetPhysicalShift()
      releaseSoftBreak()
      releaseParaBreak()
    }
  }, [])

  // While focus is parked on a keep-edit control the editor is already blurred, so its
  // onBlur can't fire again — a press anywhere else must still commit instead of silently dropping the edit
  const commitRef = useRef(commit)
  commitRef.current = commit
  useEffect(() => {
    const onDown = (ev: MouseEvent) => {
      const div = ref.current
      if (!div || div.contains(document.activeElement)) return
      const t = ev.target instanceof HTMLElement ? ev.target : null
      if (t && (div.parentElement?.contains(t) || t.closest('[data-keep-edit]'))) return
      commitRef.current()
    }
    document.addEventListener('mousedown', onDown, true)
    return () => {
      document.removeEventListener('mousedown', onDown, true)
      savedSel = null
    }
  }, [])

  return (
    // Outer layer = the whole text box (the edit-frame border is drawn here);
    // inner contentEditable edits in place with a transparent background (the canvas already hides this node's text), flex implements the vertical anchor.
    // Insert>Text Box / spAutoFit: height follows content (WPS). Width stays the shape
    // width so wrap happens instead of stretching the box sideways.
    // Other shapes keep a fixed height; overflowing content shows past the outline.
    // border uses outline (takes no layout space) so the inner usable size matches canvas layout exactly
    <div
      ref={frameRef}
      style={{
        position: 'absolute',
        left: box.x,
        top: box.y,
        width: box.w,
        height: growVert ? 'auto' : box.h,
        minHeight: growVert ? box.h : undefined,
        transform: `rotate(${box.rotationDeg ?? 0}deg) scale(${box.flipH ? -1 : 1}, ${box.flipV ? -1 : 1})`,
        transformOrigin: 'center center',
        zIndex: 20,
        display: 'flex',
        overflow: 'visible',
        // Vertical text: the row direction keeps the anchor mapping (top = the
        // flow-start edge, like the engine's anchoring)
        flexDirection: vertText ? (vertLtr ? 'row' : 'row-reverse') : 'column',
        justifyContent:
          growVert || anchor === 'top'
            ? 'flex-start'
            : anchor === 'middle'
              ? 'center'
              : anchor === 'bottom'
                ? 'flex-end'
                : 'flex-start',
        // 2 device px, zoom-compensated: the canvas is CSS-scaled, and 2 CSS px reads
        // twice as heavy on retina displays
        outline: `${2 / (globalThis.devicePixelRatio || 1) / Math.max(zoom, 0.1)}px solid ${frameColor}`,
      }}
    >
      <div
        ref={ref}
        className="ppt-txedit"
        contentEditable
        spellCheck
        suppressContentEditableWarning
        onBlur={(e) => {
          // Keep-edit controls (font size input, native color picker) take focus without committing;
          // they save/restore the selection and apply to it instead of element-level
          const to = e.relatedTarget instanceof HTMLElement ? e.relatedTarget : null
          if (to?.closest('[data-keep-edit]')) return
          commit()
        }}
        onClick={(e) => {
          // MoreAI embed: click follows (docs-like). Native: ⌘/Ctrl+click, matching PowerPoint.
          const moreai = new URLSearchParams(window.location.search).get('moreai') === '1'
          if ((!moreai && !(e.metaKey || e.ctrlKey)) || !onFollowLink) return
          const a = e.target instanceof Node ? linkAround(e.target) : null
          const target = a && decodeLinkTarget(a.getAttribute('href'))
          if (target) {
            e.preventDefault()
            onFollowLink(target)
          }
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            // Esc = commit the text and return to shape-selected state (input not lost);
            // when unchanged, commit internally goes through onCancel and produces no history step
            e.preventDefault()
            commit()
          } else if (isEnterKey(e.nativeEvent) && (e.metaKey || e.ctrlKey)) {
            e.preventDefault()
            commit()
          } else if (e.key === 'Tab' && onTabNav) {
            // Table cells: Tab commits and jumps to the next cell / Shift+Tab previous;
            // block the default focus move (otherwise blur falsely triggers commit-and-exit)
            e.preventDefault()
            const div = ref.current
            if (div) {
              const paras = extractParagraphs(div, norm)
              const changed = JSON.stringify(paras) !== initialRef.current
              onTabNav(changed ? paras : null, e.shiftKey ? -1 : 1)
            }
          } else if (e.key === 'Tab') {
            // Multi-level lists: Tab/⇧Tab adjust the caret paragraph's indent level (lvl written on commit;
            // editing only shows a marginLeft visual hint, real indentation is laid out by the canvas per master styles)
            e.preventDefault()
            const selNow = window.getSelection()
            let blk: HTMLElement | null =
              selNow?.anchorNode instanceof HTMLElement
                ? selNow.anchorNode
                : (selNow?.anchorNode?.parentElement ?? null)
            while (blk && blk !== ref.current && blk.tagName !== 'DIV') blk = blk.parentElement
            if (blk && blk !== ref.current) {
              const cur = parseInt(blk.dataset.level ?? '0', 10) || 0
              const next = Math.max(0, Math.min(8, cur + (e.shiftKey ? -1 : 1)))
              if (next) {
                blk.dataset.level = String(next)
                blk.style.marginLeft = `${next * 24}px`
              } else {
                delete blk.dataset.level
                blk.style.marginLeft = ''
              }
            }
          } else if (isSoftLineBreakShortcut(e.nativeEvent)) {
            // Shift+Enter = in-paragraph soft break (<a:br/>). Match physical Enter/Shift so
            // Chinese IME (key="Process", cleared shiftKey) works the same as English layout.
            e.preventDefault()
            insertTextBoxSoftBreak()
          } else if (
            isEnterKey(e.nativeEvent) &&
            !e.nativeEvent.isComposing &&
            !e.metaKey &&
            !e.ctrlKey &&
            !e.altKey
          ) {
            e.preventDefault()
            insertTextBoxParagraphBreak()
          }
        }}
        style={{
          // wrap=none: no wrapping (width follows content, same as the canvas overflow behavior); soft-break \n still breaks via pre.
          // insets use padding not margin: paragraph divs' marginTop (paragraph spacing) must not collapse with the root node
          // +2px: contenteditable caret/IME underline can otherwise wrap the last glyph
          // of a line that the canvas still paints as one ("dddd" → "ddd" / "d").
          width: wrap ? Math.max(box.w, 40 + insets.l + insets.r) + 2 : 'max-content',
          // nowrap keeps max-content growth for overflow, but never below the box width:
          // a narrower block would defeat per-paragraph text-align (centered titles would
          // visually snap left on entering edit) — the canvas centers within the box
          minWidth: wrap ? undefined : Math.max(box.w, 40 + insets.l + insets.r),
          // Only an empty body needs a synthetic height (so typing matches the canvas line
          // height); inflating a laid-out body distorts the flex vertical anchor — a
          // middle-anchored single line with tight spacing sat a few px too high in edit
          minHeight: node.text?.lines.length ? undefined : baseFontSize * 1.2 + insets.t + insets.b,
          // extractParagraphs reads the root alignment back, so the fresh-shape preview
          // is also what gets committed
          ...(fresh ? { textAlign: 'center' as const } : {}),
          padding: `${insets.t}px ${insets.r}px ${insets.b}px ${insets.l}px`,
          fontSize: baseFontSize,
          fontFamily: baseFont,
          color: baseColor,
          // Paragraph divs carry PPT line heights in px from populateEditorDom.
          // Freshly typed blocks inherit this 1.2 (PowerPoint "single") until
          // normalizeTypedParagraphBoxes writes an explicit px value.
          lineHeight: PPT_SINGLE_LINE,
          outline: 'none',
          background: 'transparent',
          caretColor: baseColor,
          boxSizing: 'border-box',
          whiteSpace: wrap ? 'pre-wrap' : 'pre',
          overflow: 'visible',
          overflowWrap: 'normal',
          wordBreak: 'normal',
          flexShrink: 0,
          // Vertical editing: the browser lays out real vertical text (upright CJK, rotated
          // Latin). The block axis is horizontal, so width follows content (columns) and the
          // row flex stretch pins the height to the box so wrap breaks columns at box height.
          ...(vertText
            ? {
                writingMode:
                  vertMode === 'vert270'
                    ? ('sideways-lr' as const)
                    : vertMode === 'wordArtVert'
                      ? ('vertical-lr' as const)
                      : ('vertical-rl' as const),
                textOrientation:
                  vertMode === 'vert'
                    ? ('sideways' as const)
                    : vertMode === 'wordArtVert'
                      ? ('upright' as const)
                      : ('mixed' as const),
                width: 'max-content',
                minWidth: undefined,
                minHeight: undefined,
                // Keep natural column width: overflow spills left (the flow direction), it
                // must not compress into extra column breaks
                flexShrink: 0,
              }
            : {}),
        }}
      />
    </div>
  )
}

/** css text-align → paragraph align (start/empty treated as unspecified, main process falls back to the original value). */
function cssAlign(v: string): EditParagraph['align'] | undefined {
  if (v === 'left' || v === 'center' || v === 'right' || v === 'justify') return v
  return undefined
}

/** Whether adjacent runs share source and format (mergeable losslessly). When both srcRun are undefined, merge newly typed text by format. */
function sameRunFormat(a: EditRun, b: EditRun): boolean {
  return (
    a.srcRun === b.srcRun &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.baseline === b.baseline &&
    a.fontSize === b.fontSize &&
    a.fontFamily === b.fontFamily &&
    a.color === b.color &&
    (a.link ? encodeLinkTarget(a.link) : '') === (b.link ? encodeLinkTarget(b.link) : '')
  )
}

/** Merge adjacent same-source same-format runs: stitch fragments split by CJK per-char/latin per-word/execCommand
 * back into whole runs, keeping model run boundaries stable (required for the lossless in-place patch path on save). */
function mergeAdjacentRuns(runs: EditRun[]): EditRun[] {
  const out: EditRun[] = []
  for (const r of runs) {
    const last = out[out.length - 1]
    if (last && sameRunFormat(last, r)) last.text += r.text
    else out.push({ ...r })
  }
  return out
}

/** Split "\n" (soft breaks) inside run text into standalone sentinel runs: at the model/save layer a soft
 * break = a standalone "\n" run (maps to <a:br/>). Spans that are entirely "\n" (initially rendered sentinels) keep their srcRun as is. */
function splitSoftBreaks(runs: EditRun[]): EditRun[] {
  const out: EditRun[] = []
  for (const r of runs) {
    if (r.text === '\n' || !r.text.includes('\n')) {
      out.push(r)
      continue
    }
    r.text.split('\n').forEach((part, i) => {
      if (i > 0)
        out.push({
          text: '\n',
          bold: false,
          italic: false,
          underline: false,
          strike: false,
          baseline: 0,
        })
      if (part) out.push({ ...r, text: part })
    })
  }
  return out
}

/** Walk the contentEditable DOM → paragraph/run structure. <br>/<div> split paragraphs, span styles → run format,
 * div/root text-align → paragraph alignment (product of execCommand justify*).
 * data-src-para/data-src-run carry back source indexes (browser Enter splits divs copying data attributes,
 * so both halves inherit the same source).
 * bold/italic/underline are committed as explicit booleans (the DOM is the authoritative state), avoiding
 * the main process's ?? fallback inheriting wrongly.
 * norm = viewport scale × autofit fontScale: DOM font sizes are viewport px, divided by norm to model pt. */
export function extractParagraphs(root: HTMLElement, norm: number): EditParagraph[] {
  const paragraphs: EditParagraph[] = []
  let cur: EditRun[] = []
  const rootAlign = cssAlign(root.style.textAlign)
  let curAlign = rootAlign
  let curSrcPara: number | undefined
  let curLevel = 0
  let curFmt: Partial<EditParagraph> = {}
  const pushPara = () => {
    paragraphs.push({
      runs: cur.length
        ? splitSoftBreaks(mergeAdjacentRuns(cur))
        : [{ text: '', bold: false, italic: false, underline: false, strike: false, baseline: 0 }],
      ...(curAlign ? { align: curAlign } : {}),
      level: curLevel,
      ...(curSrcPara != null ? { srcPara: curSrcPara } : {}),
      ...curFmt,
    })
    cur = []
  }

  const walk = (node: Node, inherited: Partial<EditRun>) => {
    if (node.nodeType === Node.TEXT_NODE) {
      // Caret-only superscript inserts a zwsp so the next key lands in the span.
      // It is an editing marker, not document text.
      let text = node.textContent ?? ''
      if ((node.parentElement as HTMLElement | null)?.dataset.baseline != null)
        text = text.replace(/\u200b/g, '')
      if (text) {
        cur.push({
          text,
          bold: !!inherited.bold,
          italic: !!inherited.italic,
          underline: !!inherited.underline,
          strike: !!inherited.strike,
          baseline: inherited.baseline ?? 0,
          ...(inherited.fontSize != null ? { fontSize: inherited.fontSize } : {}),
          ...(inherited.fontFamily ? { fontFamily: inherited.fontFamily } : {}),
          ...(inherited.color ? { color: inherited.color } : {}),
          ...(inherited.srcRun != null ? { srcRun: inherited.srcRun } : {}),
          link: inherited.link ?? null, // explicit null = no link (the DOM is authoritative here)
        })
      }
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const el = node as HTMLElement
    if (el.tagName === 'BR') {
      pushPara()
      return
    }
    const isBlock = el.tagName === 'DIV' || el.tagName === 'P'
    if (isBlock && cur.length) {
      // Some browsers use <div> for line breaks
      pushPara()
    }
    if (isBlock) {
      curAlign = cssAlign(el.style.textAlign) ?? rootAlign
      const sp = el.dataset ? parseInt(el.dataset.srcPara ?? '', 10) : NaN
      curSrcPara = Number.isNaN(sp) ? undefined : sp
      const lv = el.dataset ? parseInt(el.dataset.level ?? '0', 10) : 0
      curLevel = Number.isNaN(lv) ? 0 : lv
      curFmt = {}
      const ds = el.dataset
      if (ds?.bullet === 'char' || ds?.bullet === 'number' || ds?.bullet === 'none') {
        curFmt.bullet = ds.bullet
        if (ds.bullet === 'char' && ds.bulletChar) curFmt.bulletChar = ds.bulletChar
      }
      const num = (v: string | undefined) => {
        const n = parseFloat(v ?? '')
        return Number.isNaN(n) ? undefined : n
      }
      const ls = num(ds?.lineSpacingPct)
      if (ls != null) curFmt.lineSpacingPct = ls
      const sb = num(ds?.spaceBeforePt)
      if (sb != null) curFmt.spaceBeforePt = sb
      const sa = num(ds?.spaceAfterPt)
      if (sa != null) curFmt.spaceAfterPt = sa
      if (ds?.rtl === '1' || ds?.rtl === '0') curFmt.rtl = ds.rtl === '1'
    }
    const style = el.style
    const cs = window.getComputedStyle(el)
    const next: Partial<EditRun> = { ...inherited }
    const sr = el.dataset ? parseInt(el.dataset.srcRun ?? '', 10) : NaN
    if (!Number.isNaN(sr)) next.srcRun = sr
    if (el.tagName === 'A') {
      const link = decodeLinkTarget(el.getAttribute('href'))
      if (link) next.link = link
    }
    if (
      style.fontWeight === 'bold' ||
      cs.fontWeight === '700' ||
      el.tagName === 'B' ||
      el.tagName === 'STRONG'
    )
      next.bold = true
    if (style.fontStyle === 'italic' || el.tagName === 'I' || el.tagName === 'EM')
      next.italic = true
    if ((style.textDecoration || cs.textDecorationLine).includes('underline') || el.tagName === 'U')
      next.underline = true
    if (
      (style.textDecoration || cs.textDecorationLine).includes('line-through') ||
      el.tagName === 'S' ||
      el.tagName === 'STRIKE' ||
      el.tagName === 'DEL'
    )
      next.strike = true
    // Super/subscript. data-baseline is what the ribbon writes while focused
    // (vertical-align does not repaint in a focused contentEditable). <sub>/<sup>
    // and vertical-align still round-trip text that was already superscript.
    // Layout fragments use vertical-align:top — that must not wipe an inherited super/sub.
    const marked = el.dataset?.baseline
    if (marked != null && marked !== '') {
      const n = Number(marked)
      if (Number.isFinite(n)) next.baseline = n
    } else if (el.tagName === 'SUP' || style.verticalAlign === 'super') next.baseline = 30
    else if (el.tagName === 'SUB' || style.verticalAlign === 'sub') next.baseline = -25
    else if (style.verticalAlign === 'baseline') next.baseline = 0
    if (style.color) next.color = rgbToHex(style.color)
    if (style.fontSize) {
      const px = parseFloat(style.fontSize)
      if (!Number.isNaN(px)) next.fontSize = pxToPt(px, norm)
    }
    if (style.fontFamily) {
      const fam = firstFontFamily(style.fontFamily)
      if (fam) {
        const orig = el.dataset?.font
        const baked = el.dataset?.displayFont
        if (baked && baked.toLowerCase() === fam.toLowerCase()) {
          // Display font unchanged since populate: commit the model's original name; a run without
          // an explicit model font commits none (the layout's Arial default / missing-font
          // substitution baked into the DOM is not a user change)
          if (orig) next.fontFamily = orig
          else delete next.fontFamily
        } else {
          next.fontFamily =
            orig && firstFontFamily(displayFontFamily(orig)).toLowerCase() === fam.toLowerCase()
              ? orig
              : fam
        }
      }
    }
    for (const child of Array.from(el.childNodes)) walk(child, next)
    if (isBlock) {
      // Block end = paragraph end (alignment/source paragraph doesn't leak into following siblings)
      if (cur.length) pushPara()
      curAlign = rootAlign
      curLevel = 0
      curSrcPara = undefined
      curFmt = {}
    }
  }

  for (const child of Array.from(root.childNodes)) walk(child, {})
  pushPara()
  // Drop trailing empty paragraphs (unless everything is empty)
  while (paragraphs.length > 1 && paragraphs[paragraphs.length - 1]!.runs.every((r) => !r.text)) {
    paragraphs.pop()
  }
  return paragraphs
}

/**
 * Selection handoff for ribbon controls that must take real focus (font size input, native color
 * picker): save the editor's Range before focus leaves, restore it (re-focusing the
 * editor) right before applying, so the command hits the original selection instead of element-level.
 */
let savedSel: { root: HTMLElement; range: Range } | null = null

export function saveEditSelection(): void {
  const root = document.activeElement
  const sel = window.getSelection()
  if (!(root instanceof HTMLElement) || !root.isContentEditable || !sel?.rangeCount) return
  savedSel = { root, range: sel.getRangeAt(0).cloneRange() }
}

export function restoreEditSelection(): boolean {
  if (!savedSel?.root.isConnected) return false
  savedSel.root.focus()
  const sel = window.getSelection()
  if (!sel) return false
  sel.removeAllRanges()
  sel.addRange(savedSel.range)
  return true
}

/** Nearest <a href> wrapping a node (bounded by the contentEditable root). */
function linkAround(node: Node | null): HTMLAnchorElement | null {
  const el = node instanceof HTMLElement ? node : node?.parentElement
  const a = el?.closest('a[href]')
  return a instanceof HTMLAnchorElement && el?.closest('[contenteditable="true"]') ? a : null
}

function activeEditRange(): Range | null {
  const sel = window.getSelection()
  if (sel?.rangeCount) return sel.getRangeAt(0)
  return savedSel?.range ?? null
}

/**
 * The editor selection's current hyperlink (nearest <a> around the selection start; the saved
 * ribbon-handoff selection is consulted when focus already left the editor). For dialog echo-back.
 */
export function selectionLink(): LinkTargetOp | null {
  const range = activeEditRange()
  const a = linkAround(range?.startContainer ?? null)
  return a ? decodeLinkTarget(a.getAttribute('href')) : null
}

/** Visible text of the current (or saved) editor selection — prefill for “text to display”. */
export function selectionDisplayText(): string {
  const range = activeEditRange()
  if (!range || range.collapsed) return ''
  return range.toString()
}

/** Expand a caret to the word / text node it sits in, so Insert Link does not dump the URL as new copy. */
function expandCollapsedRange(range: Range): void {
  const a = linkAround(range.startContainer)
  if (a) {
    range.selectNodeContents(a)
    return
  }
  const node = range.startContainer
  if (node.nodeType !== Node.TEXT_NODE) return
  const text = node.textContent ?? ''
  if (!text) return
  let start = range.startOffset
  let end = range.startOffset
  const isBreak = (ch: string) => /\s/.test(ch)
  while (start > 0 && !isBreak(text[start - 1]!)) start--
  while (end < text.length && !isBreak(text[end]!)) end++
  if (end > start) {
    range.setStart(node, start)
    range.setEnd(node, end)
    return
  }
  range.setStart(node, 0)
  range.setEnd(node, text.length)
}

function wrapRangeWithLink(range: Range, href: string, label?: string): void {
  const selected = range.toString()
  const text = (label ?? '').trim()
  if (text && text !== selected) {
    range.deleteContents()
    const a = document.createElement('a')
    a.setAttribute('href', href)
    a.textContent = text
    range.insertNode(a)
    return
  }
  const sel = window.getSelection()
  if (sel) {
    sel.removeAllRanges()
    sel.addRange(range)
  }
  document.execCommand('createLink', false, href)
}

/**
 * Set/clear a hyperlink on the editor selection (restoring the saved selection first — the link
 * dialog took focus). A collapsed caret expands to the current word/run so the URL is applied to
 * existing text instead of replacing it. `displayText` only replaces glyphs when the user edited it.
 */
export function applySelectionLink(target: LinkTargetOp | null, displayText?: string): boolean {
  if (!restoreEditSelection()) {
    const root = document.querySelector('.ppt-txedit')
    if (!(root instanceof HTMLElement) || !root.isContentEditable) return false
    root.focus()
    const selAll = window.getSelection()
    if (!selAll) return false
    const all = document.createRange()
    all.selectNodeContents(root)
    selAll.removeAllRanges()
    selAll.addRange(all)
  }
  const sel = window.getSelection()
  if (!sel?.rangeCount) return false
  const range = sel.getRangeAt(0)
  if (!target) {
    if (range.collapsed) {
      const a = linkAround(range.startContainer)
      if (!a) return false
      range.selectNodeContents(a)
      sel.removeAllRanges()
      sel.addRange(range)
    }
    document.execCommand('unlink')
    return true
  }
  if (range.collapsed) expandCollapsedRange(range)
  if (range.collapsed) {
    const label = (displayText ?? '').trim()
    if (!label) return false
    const a = document.createElement('a')
    a.setAttribute('href', encodeLinkTarget(target))
    a.textContent = label
    range.insertNode(a)
    return true
  }
  wrapRangeWithLink(range, encodeLinkTarget(target), displayText)
  return true
}

/**
 * Per-paragraph format while editing: mark the paragraph divs covered by the
 * caret/selection; extractParagraphs carries the marks to the main process on commit.
 * Toggle-off (clicking the active bullet kind again) resolves against the first covered
 * paragraph's current state. Returns false when no editor selection is available
 * (the caller falls back to the element-level op).
 */
export function applySelectionParagraphFormat(patch: {
  bullet?: 'char' | 'number' | 'none'
  bulletChar?: string
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  rtl?: boolean
}): boolean {
  const sel = window.getSelection()
  if (!sel?.rangeCount) return false
  const range = sel.getRangeAt(0)
  // Editor root resolved from the selection itself (works when a keep-edit control holds focus)
  const startEl =
    range.startContainer instanceof HTMLElement
      ? range.startContainer
      : range.startContainer.parentElement
  const root =
    startEl?.closest('[data-src-para]')?.parentElement ??
    (startEl?.querySelector('[data-src-para]') ? startEl : null) ??
    (savedSel?.root.isConnected ? savedSel.root : null)
  if (!root) return false
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement =>
      el instanceof HTMLElement && el.tagName === 'DIV' && range.intersectsNode(el),
  )
  if (!blocks.length) return false
  let bullet = patch.bullet
  if (bullet && bullet !== 'none' && !patch.bulletChar) {
    const cur = blocks[0]!.dataset.bullet ?? blocks[0]!.dataset.hadBullet
    if (cur === bullet) bullet = 'none'
  }
  for (const b of blocks) {
    if (bullet) {
      b.dataset.bullet = bullet
      if (bullet === 'char' && patch.bulletChar) b.dataset.bulletChar = patch.bulletChar
      else if (bullet !== 'char') delete b.dataset.bulletChar
    }
    if (patch.lineSpacingPct != null) {
      b.dataset.lineSpacingPct = String(patch.lineSpacingPct)
      // Match text-layout.ts PPT_SINGLE = 1.2: 100% → 1.2em, not CSS unitless 1.0
      const sizePx =
        parseFloat(
          (
            b.querySelector('[style*="font-size"]') as HTMLElement | null
          )?.style.fontSize.replace(/px$/i, '') ?? '',
        ) ||
        parseFloat(window.getComputedStyle(b).fontSize) ||
        18
      b.style.lineHeight = `${(PPT_SINGLE_LINE * sizePx * patch.lineSpacingPct) / 100}px`
    }
    if (patch.spaceBeforePt != null) {
      b.dataset.spaceBeforePt = String(patch.spaceBeforePt)
      const norm = parseFloat(root.dataset.norm ?? '1') || 1
      // Live preview: pt → overlay viewport px. Neighbor after-spacing still lives on
      // the next block's marginTop (same stacking populateEditorDom uses).
      const beforePx = ((patch.spaceBeforePt * 96) / 72) * norm
      b.dataset.spcBefPx = String(beforePx)
      const prev = b.previousElementSibling as HTMLElement | null
      const prevAfter = prev ? paragraphSpacingPx(prev, 'after') : 0
      b.style.marginTop = `${prevAfter + beforePx}px`
      b.dataset.layoutSpaced = '1'
    }
    if (patch.spaceAfterPt != null) {
      b.dataset.spaceAfterPt = String(patch.spaceAfterPt)
      const norm = parseFloat(root.dataset.norm ?? '1') || 1
      const afterPx = ((patch.spaceAfterPt * 96) / 72) * norm
      b.dataset.spcAftPx = String(afterPx)
      const next = b.nextElementSibling as HTMLElement | null
      if (next) {
        const nextBefore = paragraphSpacingPx(next, 'before')
        next.style.marginTop = `${afterPx + nextBefore}px`
        next.dataset.layoutSpaced = '1'
      }
    }
    if (patch.rtl != null) {
      b.dataset.rtl = patch.rtl ? '1' : '0'
      b.dir = patch.rtl ? 'rtl' : 'ltr' // live preview; the canvas re-lays out on commit
    }
  }
  return true
}

/** Effective base direction at the editing selection, read from the overlay DOM (computed
 * direction covers dir="auto" inference and explicit toggles alike). undefined = no overlay
 * mounted; null = mixed. */
export function liveRtl(): boolean | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!(root instanceof HTMLElement)) return undefined
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length) return false
  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  const found = new Set<boolean>()
  for (const b of blocks) {
    if (range && !range.intersectsNode(b)) continue
    found.add(window.getComputedStyle(b).direction === 'rtl')
  }
  if (!found.size) for (const b of blocks) found.add(window.getComputedStyle(b).direction === 'rtl')
  return found.size === 1 ? [...found][0]! : null
}

/** Bullet-gallery highlight while editing: union of the live paragraph marks across the edit
 * root ('' = none, '#num' = numbered, glyph = char bullet). `undefined` = no uncommitted
 * paragraph-format change, the render tree is still accurate; `null` = unknowable (mixed, or a
 * re-toggled char bullet whose glyph lives only in the engine's uncommitted state). */
export function liveBulletChar(): string | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!root) return undefined
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length || !blocks.some((b) => b.dataset.bullet)) return undefined
  const found = new Set<string>()
  for (const b of blocks) {
    const kind = b.dataset.bullet ?? b.dataset.hadBullet
    if (!kind || kind === 'none') {
      found.add('')
      continue
    }
    if (kind === 'number') {
      found.add('#num')
      continue
    }
    // char: explicit glyph from the gallery, engine default ('•') for a fresh bullet; a
    // paragraph whose original glyph never reached the DOM stays unknowable
    const glyph =
      b.dataset.bulletChar ??
      (b.dataset.bullet === 'char' && b.dataset.hadBullet == null ? '•' : null)
    if (glyph == null) return null
    found.add(glyph)
  }
  return found.size === 1 ? [...found][0]! : null
}

/** Paragraph alignment at the editing selection, read from the overlay DOM (execCommand
 * justify* products live only there until commit). Blocks intersecting the selection count —
 * the caret's block when collapsed; a block with no inline text-align falls back to the
 * root's, then 'left' (the engine default, so some alignment is always current).
 * undefined = no overlay mounted; null = mixed. */
export function liveAlign(): 'left' | 'center' | 'right' | 'justify' | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!(root instanceof HTMLElement)) return undefined
  const rootAlign = cssAlign(root.style.textAlign)
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length) return rootAlign ?? 'left'
  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  const found = new Set<NonNullable<ReturnType<typeof cssAlign>>>()
  for (const b of blocks) {
    if (range && !range.intersectsNode(b)) continue
    found.add(cssAlign(b.style.textAlign) ?? rootAlign ?? 'left')
  }
  // selection outside the overlay (e.g. focus stolen by ribbon chrome): read all blocks
  if (!found.size)
    for (const b of blocks) found.add(cssAlign(b.style.textAlign) ?? rootAlign ?? 'left')
  return found.size === 1 ? [...found][0]! : null
}

/** True when `node` lies entirely inside `range` (overlap is not enough). */
function rangeFullyContainsNode(range: Range, node: Node): boolean {
  const nr = document.createRange()
  try {
    nr.selectNode(node)
  } catch {
    nr.selectNodeContents(node)
  }
  return (
    range.compareBoundaryPoints(Range.START_TO_START, nr) <= 0 &&
    range.compareBoundaryPoints(Range.END_TO_END, nr) >= 0
  )
}

/**
 * Font size increase/decrease while editing. Chromium often emits <font size="7">
 * only on the first execCommand; later clicks apply `font-size: xxx-large` on a
 * span — querying font[size=7] then no-ops (grow once, shrink never). Only step
 * sized spans the selection fully covers; a partial hit inside a run-container
 * must use execCommand so only the selected characters grow.
 */
export function resizeSelectionFont(dir: 1 | -1): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const norm = parseFloat(root.dataset.norm ?? '') || 1
  const pxOf = (pt: number) => `${(pt * 96 * norm) / 72}px`
  const stepPx = (px: number) => stepFontSizePt(pxToPt(px, norm), dir)

  const paint = (el: HTMLElement, pt: number) => {
    el.style.fontSize = pxOf(pt)
    releaseFragment(el.closest('[data-layout-fragment]'))
    el.querySelectorAll<HTMLElement>('[data-layout-fragment]').forEach(releaseFragment)
  }

  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  const existing: HTMLElement[] = []
  if (range && !range.collapsed) {
    root.querySelectorAll<HTMLElement>('span').forEach((span) => {
      const css = span.style.fontSize
      if (!css || /xxx-large/i.test(css)) return
      if (!rangeFullyContainsNode(range, span)) return
      existing.push(span)
    })
  }
  const deepest = existing.filter((s) => !existing.some((o) => o !== s && s.contains(o)))
  if (deepest.length) {
    for (const span of deepest) {
      const px = parseFloat(span.style.fontSize) || parseFloat(window.getComputedStyle(span).fontSize) || 18
      paint(span, stepPx(px))
    }
    reselectSpans(deepest)
    return
  }

  document.execCommand('styleWithCSS', false, 'false')
  document.execCommand('fontSize', false, '7')
  const spans: HTMLElement[] = []
  const fromPlaceholder = (el: HTMLElement, basePx: number) => {
    const pt = stepPx(basePx)
    const span = document.createElement('span')
    while (el.firstChild) span.appendChild(el.firstChild)
    el.replaceWith(span)
    paint(span, pt)
    spans.push(span)
  }
  root.querySelectorAll('font[size="7"]').forEach((f) => {
    const font = f as HTMLElement
    const basePx = parseFloat(window.getComputedStyle(font.parentElement ?? root).fontSize) || 18
    fromPlaceholder(font, basePx)
  })
  root.querySelectorAll<HTMLElement>('span').forEach((s) => {
    if (!/xxx-large/i.test(s.style.fontSize)) return
    const basePx = parseFloat(window.getComputedStyle(s.parentElement ?? root).fontSize) || 18
    s.style.fontSize = pxOf(stepPx(basePx))
    paint(s, stepPx(basePx))
    spans.push(s)
  })
  reselectSpans(spans)
}

/** Next/previous ladder size; beyond the ladder ±10pt, clamped to 8~400 */
export function stepFontSizePt(cur: number, dir: 1 | -1): number {
  const max = FONT_SIZES[FONT_SIZES.length - 1]!
  if (dir > 0) return cur >= max ? Math.min(400, cur + 10) : FONT_SIZES.find((s) => s > cur)!
  if (cur > max) return Math.max(max, cur - 10)
  for (let i = FONT_SIZES.length - 1; i >= 0; i--) if (FONT_SIZES[i]! < cur) return FONT_SIZES[i]!
  return FONT_SIZES[0]!
}

/** replaceWith kills the live selection — re-select the new spans so the highlight and repeated grow/shrink clicks survive */
function reselectSpans(spans: HTMLElement[]): void {
  if (!spans.length) return
  const sel = window.getSelection()
  if (!sel) return
  const range = document.createRange()
  range.setStartBefore(spans[0]!)
  range.setEndAfter(spans[spans.length - 1]!)
  sel.removeAllRanges()
  sel.addRange(range)
  saveEditSelection()
}

/** super / sub / explicit off, walking from `start` up to (not including) the editor root.
 * undefined = this node carries no baseline mark. */
function baselineKindFrom(start: Node | null, root: HTMLElement): 'super' | 'sub' | null | undefined {
  let n: Node | null = start
  while (n && n !== root) {
    if (n instanceof HTMLElement) {
      const marked = n.dataset.baseline
      if (marked != null && marked !== '') {
        const v = Number(marked)
        if (v > 0) return 'super'
        if (v < 0) return 'sub'
        return null
      }
      if (n.tagName === 'SUP' || n.style.verticalAlign === 'super') return 'super'
      if (n.tagName === 'SUB' || n.style.verticalAlign === 'sub') return 'sub'
      if (n.style.verticalAlign === 'baseline') return null
    }
    n = n.parentNode
  }
  return undefined
}

/** reselectSpans parks the range on the parent, around the styled span. */
function elementCoveredByRange(range: Range): HTMLElement | null {
  if (range.startContainer !== range.endContainer) return null
  if (range.startContainer.nodeType !== Node.ELEMENT_NODE) return null
  const el = range.startContainer.childNodes[range.startOffset]
  const end = range.startContainer.childNodes[range.endOffset - 1]
  return el instanceof HTMLElement && el === end ? el : null
}

/** Current caret/selection super/sub from overlay DOM (Chromium queryCommandState is unreliable). */
export function selectionBaselineKind(): 'super' | 'sub' | null {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return null
  const sel = window.getSelection()
  if (!sel?.anchorNode || !root.contains(sel.anchorNode)) return null
  const range = sel.rangeCount ? sel.getRangeAt(0) : null
  const covered = range && !range.collapsed ? elementCoveredByRange(range) : null
  const direct = baselineKindFrom(covered ?? sel.anchorNode, root)
  if (direct !== undefined) return direct
  if (!range || range.collapsed) return null
  // The range can cover several styled spans; the anchor then sits on their parent.
  let sawSuper = false
  let sawSub = false
  root.querySelectorAll<HTMLElement>('span,sup,sub,a').forEach((el) => {
    if (!range.intersectsNode(el)) return
    const k = baselineKindFrom(el, root)
    if (k === 'super') sawSuper = true
    else if (k === 'sub') sawSub = true
  })
  if (sawSuper && !sawSub) return 'super'
  if (sawSub && !sawSuper) return 'sub'
  return null
}

/** Canvas shift is 30% / −25% of the font size. position updates while contentEditable
 * is focused; vertical-align does not (it shows up only after blur). */
function paintBaseline(el: HTMLElement, va: 'super' | 'sub' | 'baseline'): HTMLElement {
  const pct = va === 'super' ? '30' : va === 'sub' ? '-25' : '0'
  const shift = va === 'super' ? '-0.3em' : va === 'sub' ? '0.25em' : ''
  const tagged = el.tagName === 'SUP' || el.tagName === 'SUB'
  const painted =
    el.style.verticalAlign === 'super' ||
    el.style.verticalAlign === 'sub' ||
    el.style.verticalAlign === 'baseline'
  let target = el
  // A vertical-align / <sup> that was painted before focus keeps its used offset
  // until blur, even after the property is cleared. Swap in a fresh node.
  if (tagged || painted) {
    const neu = document.createElement(tagged ? 'span' : el.tagName.toLowerCase())
    for (const attr of [...el.attributes]) {
      if (attr.name === 'style') continue
      neu.setAttribute(attr.name, attr.value)
    }
    neu.style.cssText = el.style.cssText
    neu.style.verticalAlign = ''
    while (el.firstChild) neu.appendChild(el.firstChild)
    el.replaceWith(neu)
    target = neu
  }
  target.dataset.baseline = pct
  target.style.verticalAlign = ''
  if (shift) {
    target.style.position = 'relative'
    target.style.top = shift
    target.dataset.baselineShift = '1'
  } else if (target.dataset.baselineShift === '1') {
    target.style.position = ''
    target.style.top = ''
    delete target.dataset.baselineShift
  }
  return target
}

function wrapEditRange(range: Range): HTMLElement {
  const span = document.createElement('span')
  try {
    range.surroundContents(span)
  } catch {
    const contents = range.extractContents()
    span.appendChild(contents)
    range.insertNode(span)
  }
  return span
}

/**
 * Toggle super/subscript on the editing selection. vertical-align and execCommand('superscript')
 * do not move glyphs while the editor is focused (the raise showed up only after blur, when the
 * canvas redrew). Shift with position, which paints immediately and matches the canvas
 * (super +30% of font size, sub −25%). Leave layout fragments pinned so the rest of the line
 * does not reflow.
 */
export function toggleSelectionBaseline(kind: 'super' | 'sub'): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  if (!range) return
  const want = kind === 'super' ? 'super' : 'sub'
  const va = selectionBaselineKind() === kind ? 'baseline' : want

  const existing: HTMLElement[] = []
  root.querySelectorAll<HTMLElement>('span,sup,sub,a').forEach((el) => {
    if (!range.intersectsNode(el)) return
    const tag = el.tagName
    const cur = el.style.verticalAlign
    const marked = el.dataset.baseline
    if (
      (marked != null && marked !== '') ||
      tag === 'SUP' ||
      tag === 'SUB' ||
      cur === 'super' ||
      cur === 'sub' ||
      cur === 'baseline'
    ) {
      existing.push(el)
    }
  })
  const deepest = existing.filter((s) => !existing.some((o) => o !== s && s.contains(o)))
  if (deepest.length) {
    reselectSpans(deepest.map((el) => paintBaseline(el, va)))
    return
  }

  if (range.collapsed) {
    const span = document.createElement('span')
    span.textContent = '\u200b'
    range.insertNode(span)
    const painted = paintBaseline(span, va)
    const text = painted.firstChild
    if (text && sel) {
      const caret = document.createRange()
      caret.setStart(text, text.textContent?.length ?? 0)
      caret.collapse(true)
      sel.removeAllRanges()
      sel.addRange(caret)
      saveEditSelection()
    }
    return
  }

  reselectSpans([paintBaseline(wrapEditRange(range), va)])
}

/** css font-family list → first family name (unquoted). Save/display only care about the preferred font. */
export function firstFontFamily(cssList: string): string {
  return (cssList.split(',')[0] ?? '').trim().replace(/^['"]|['"]$/g, '')
}

/**
 * Change the selection's font while editing: execCommand('fontName') on the selection, then unify
 * the product into span style.fontFamily = the display stack with CJK fallbacks (extractParagraphs
 * takes only the first item on commit).
 */
export function applySelectionFontFamily(family: string): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  document.execCommand('styleWithCSS', false, 'true')
  document.execCommand('fontName', false, family)
  const display = displayFontFamily(family)
  const displayFirst = firstFontFamily(display)
  // Handle both products: <font face="…"> (path where styleWithCSS doesn't apply) and span style
  root.querySelectorAll('font[face]').forEach((f) => {
    const font = f as HTMLElement
    const span = document.createElement('span')
    span.style.fontFamily = display
    span.dataset.font = family
    span.dataset.displayFont = displayFirst
    while (font.firstChild) span.appendChild(font.firstChild)
    font.replaceWith(span)
  })
  root.querySelectorAll('span').forEach((s) => {
    const cur = s.style.fontFamily
    if (!cur) return
    const first = firstFontFamily(cur)
    if (first.toLowerCase() !== family.toLowerCase() && first.toLowerCase() !== displayFirst.toLowerCase())
      return
    s.style.fontFamily = display
    s.dataset.font = family
    s.dataset.displayFont = displayFirst
  })
}

/** Set an absolute font size (pt) on the selection while editing: fontSize=7 placeholder then replaced by a px span (same as resizeSelectionFont). */
export function setSelectionFontSizePt(pt: number): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const norm = parseFloat(root.dataset.norm ?? '') || 1
  const px = Math.min(400, Math.max(8, (pt * 96) / 72)) * norm
  document.execCommand('styleWithCSS', false, 'false')
  document.execCommand('fontSize', false, '7')
  const spans: HTMLElement[] = []
  root.querySelectorAll('font[size="7"]').forEach((f) => {
    const font = f as HTMLElement
    const span = document.createElement('span')
    span.style.fontSize = `${px}px`
    while (font.firstChild) span.appendChild(font.firstChild)
    font.replaceWith(span)
    spans.push(span)
  })
  reselectSpans(spans)
}

// Viewport px font size → model pt (divide back by viewport scale × autofit fontScale).
// Half-pt resolution, matching the ribbon size box (commitSizeDraft).
function pxToPt(px: number, norm: number): number {
  return Math.round(((px * 72) / (96 * norm)) * 2) / 2
}

function normalizeCss(c: string): string {
  if (/^#?[0-9A-Fa-f]{8}$/.test(c)) return `#${c.replace(/^#/, '').slice(0, 6)}`
  return c.startsWith('#') ? c : `#${c}`
}

function rgbToHex(rgb: string): string | undefined {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(rgb)
  if (!m) return undefined
  const h = (n: string) => parseInt(n, 10).toString(16).padStart(2, '0')
  return `#${h(m[1]!)}${h(m[2]!)}${h(m[3]!)}`.toUpperCase()
}
