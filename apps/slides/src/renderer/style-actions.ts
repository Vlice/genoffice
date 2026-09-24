/**
 * Element/style actions extracted from App.tsx: fonts, paragraph
 * format, fill/stroke, background, theme, table style, and chart edits.
 * Functions read the latest App state through ActionCtx.
 */
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import type {
  EditBackgroundOp,
  EditChartOp,
  EditStrokeOp,
  EditTableStyleOp,
  GradientFillSpec,
} from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { FIT_WIDTH } from './app-constants'
import {
  applySelectionFontFamily,
  applySelectionParagraphFormat,
  resizeSelectionFont,
  restoreEditSelection,
  setSelectionFontSizePt,
  stepFontSizePt,
  toggleSelectionBaseline,
} from './TextEditOverlay'
import type { FormatCmd } from './components/Ribbon'
import type { SlideThemePreset } from './themes'
import { t } from './i18n/locale'

export function onFormat(ctx: ActionCtx, cmd: FormatCmd): void {
  if (cmd === 'fontSizeUp' || cmd === 'fontSizeDown') {
    const dir = cmd === 'fontSizeUp' ? 1 : -1
    if (ctx.editing || ctx.editingCell) {
      resizeSelectionFont(dir)
      return
    }
    void stepSelectedElementsFont(ctx, dir)
    return
  }
  if (cmd === 'superscript' || cmd === 'subscript') {
    const kind = cmd === 'superscript' ? 'super' : 'sub'
    if (ctx.editing || ctx.editingCell) {
      toggleSelectionBaseline(kind)
      return
    }
    void stepSelectedElementsBaseline(ctx, kind)
    return
  }
  document.execCommand(cmd)
}

type TextSizeProbe = {
  fontScale?: number
  lines?: Array<{ runs: Array<{ fontSizePx?: number; isBullet?: boolean }> }>
}

function runSizePts(
  node: { type?: string; text?: TextSizeProbe; cells?: Array<{ text?: TextSizeProbe }> },
  scale: number,
): number[] {
  const out: number[] = []
  const push = (text?: TextSizeProbe) => {
    const norm = scale * (text?.fontScale ?? 1) || 1
    for (const line of text?.lines ?? []) {
      for (const run of line.runs ?? []) {
        if (run.isBullet || !(typeof run.fontSizePx === 'number') || !(run.fontSizePx > 0)) continue
        out.push(Math.round((((run.fontSizePx / norm) * 72) / 96) * 2) / 2)
      }
    }
  }
  if (node.type === 'text' || node.type === 'shape') push(node.text)
  else if (node.type === 'table') for (const cell of node.cells ?? []) push(cell.text)
  return out
}

async function stepSelectedElementsFont(ctx: ActionCtx, dir: 1 | -1): Promise<void> {
  if (!ctx.selectedIds.length) return
  const scale = ctx.slide?.scale ?? 1
  const sizes: number[] = []
  for (const id of ctx.selectedIds) {
    const node = ctx.findNodeCtx(id)?.node
    if (node) sizes.push(...runSizePts(node, scale))
  }
  if (!sizes.length) return
  const cur = dir > 0 ? Math.max(...sizes) : Math.min(...sizes)
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  const r = await window.slidesApi.setElementFont({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    fontSizePt: stepFontSizePt(cur, dir),
    ...(groupId ? { groupId } : {}),
  })
  if (r) ctx.applySlide(ctx.current, r)
}

async function stepSelectedElementsBaseline(ctx: ActionCtx, kind: 'super' | 'sub'): Promise<void> {
  if (!ctx.selectedIds.length) return
  const pcts: number[] = []
  const push = (text?: { lines?: Array<{ runs: Array<{ text?: string; isBullet?: boolean; baselinePct?: number }> }> }) => {
    for (const line of text?.lines ?? []) {
      for (const run of line.runs ?? []) {
        if (run.isBullet || !(run.text || '').trim()) continue
        pcts.push(run.baselinePct ?? 0)
      }
    }
  }
  for (const id of ctx.selectedIds) {
    const node = ctx.findNodeCtx(id)?.node
    if (!node) continue
    if (node.type === 'text' || node.type === 'shape') push((node as ShapeRenderNode).text)
    else if (node.type === 'table') {
      for (const cell of (node as { cells?: Array<{ text?: Parameters<typeof push>[0] }> }).cells ?? []) {
        push(cell.text)
      }
    }
  }
  if (!pcts.length) return
  const allOn = kind === 'super' ? pcts.every((p) => p > 0) : pcts.every((p) => p < 0)
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  const r = await window.slidesApi.setElementFont({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    baseline: allOn ? 0 : kind === 'super' ? 30 : -25,
    ...(groupId ? { groupId } : {}),
  })
  if (r) ctx.applySlide(ctx.current, r)
}

// While editing, change the caret selection; with only an element selected, change it wholesale (applies to all the element's text runs)
export function onFontFamily(ctx: ActionCtx, family: string): void {
  if (ctx.editing || ctx.editingCell) {
    applySelectionFontFamily(family)
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      fontFamily: family,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export function onFontSize(ctx: ActionCtx, pt: number): void {
  if (ctx.editing || ctx.editingCell) {
    setSelectionFontSizePt(pt)
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      fontSizePt: pt,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// While editing, change the caret's paragraph; with only an element selected, use the element-level paragraph format op (all paragraphs)
export function onAlign(ctx: ActionCtx, align: 'left' | 'center' | 'right' | 'justify'): void {
  if (ctx.editing || ctx.editingCell) {
    document.execCommand(
      align === 'left'
        ? 'justifyLeft'
        : align === 'center'
          ? 'justifyCenter'
          : align === 'right'
            ? 'justifyRight'
            : 'justifyFull',
    )
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementParagraphFormat({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      align,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// Toggling B/I/U/strikethrough on a selected element without editing: if all runs have it on, turn off, else turn all on
export function onTextToggle(
  ctx: ActionCtx,
  kind: 'bold' | 'italic' | 'underline' | 'strike',
): void {
  if (!ctx.selectedIds.length) return
  ctx.markUnsaved()
  let allOn = true
  for (const id of ctx.selectedIds) {
    const node = ctx.findNodeCtx(id)?.node
    const text =
      node && (node.type === 'text' || node.type === 'shape')
        ? (node as ShapeRenderNode).text
        : undefined
    const runs =
      text?.lines.flatMap((l) => l.runs).filter((r) => !r.isBullet && r.text.trim()) ?? []
    if (!runs.length || runs.some((r) => !r[kind])) allOn = false
  }
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      [kind]: !allOn,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// Change font color directly on a selected element (editing mode goes through execCommand on the selection)
export function onElementTextColor(ctx: ActionCtx, hex: string): void {
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      color: hex,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export interface ParagraphFormatPatch {
  bullet?: 'char' | 'number' | 'none'
  bulletChar?: string
  bulletHangEmu?: number
  bulletSizePct?: number
  bulletColor?: string
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  rtl?: boolean
  indentDelta?: 1 | -1
}

/** Patch keys the editing-mode selection path can express; anything else stays element-level */
const SELECTION_PATCH_KEYS = new Set([
  'bullet',
  'bulletChar',
  'lineSpacingPct',
  'spaceBeforePt',
  'spaceAfterPt',
  'rtl',
])

// While editing, bullets/numbering/line spacing/paragraph spacing apply to the paragraphs covered
// by the caret/selection (PowerPoint semantics, committed with the edit); with an element selected
// they apply element-wide. Clicking the same bullet kind again = turn off (toggle semantics);
// editing mode judges by the paragraph div's marks, element mode by the render tree's bullet glyphs
export function onParagraphFormat(ctx: ActionCtx, patch: ParagraphFormatPatch): void {
  // Cell editing commits regenerate whole paragraphs from the overlay DOM, which round-trips
  // the rtl mark; the other selection keys keep their historical element-wide semantics there
  const selectable = ctx.editing
    ? Object.keys(patch).every((k) => SELECTION_PATCH_KEYS.has(k))
    : ctx.editingCell != null && Object.keys(patch).every((k) => k === 'rtl')
  if (selectable) {
    const active = document.activeElement
    if (!(active instanceof HTMLElement && active.isContentEditable)) restoreEditSelection()
    if (applySelectionParagraphFormat(patch)) return
  }
  if (!ctx.selectedIds.length) return
  // Picking an explicit char always applies (no toggle-off)
  if (
    patch.bullet &&
    patch.bullet !== 'none' &&
    !patch.bulletChar &&
    ctx.selectedIds.length === 1
  ) {
    const node = ctx.findNodeCtx(ctx.selectedIds[0]!)?.node
    const text =
      node && (node.type === 'text' || node.type === 'shape')
        ? (node as ShapeRenderNode).text
        : undefined
    const bulletRun = text?.lines.flatMap((l) => l.runs).find((r) => r.isBullet)
    const cur = bulletRun ? (/^\d/.test(bulletRun.text) ? 'number' : 'char') : null
    if (cur === patch.bullet) patch = { ...patch, bullet: 'none' }
  }
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementParagraphFormat({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      ...patch,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export async function onFill(
  ctx: ActionCtx,
  sourceId: string,
  fill: string | GradientFillSpec,
): Promise<void> {
  const groupId = ctx.groupIdOf(sourceId)
  const updated = await window.slidesApi.editFill({
    slideIndex: ctx.current,
    sourceId,
    fill,
    ...(groupId ? { groupId } : {}),
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

export async function onStroke(
  ctx: ActionCtx,
  sourceId: string,
  stroke: EditStrokeOp['stroke'],
): Promise<void> {
  const groupId = ctx.groupIdOf(sourceId)
  const updated = await window.slidesApi.editStroke({
    slideIndex: ctx.current,
    sourceId,
    stroke,
    ...(groupId ? { groupId } : {}),
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

/** Omit that distributes over union members (plain Omit collapses the EditBackgroundOp union). */
type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never

export async function onBackground(
  ctx: ActionCtx,
  op: DistributiveOmit<EditBackgroundOp, 'fitWidthPx'>,
): Promise<void> {
  if (!ctx.slide) return
  ctx.markUnsaved()
  const r = await window.slidesApi.editBackground({
    ...op,
    fitWidthPx: FIT_WIDTH,
  } as EditBackgroundOp)
  if (r) {
    ctx.setSlides(r)
    ctx.setDirty(true)
    ctx.setStatus(op.slideIndex === -1 ? t('appStatusBgAppliedAll') : '')
  }
}

// Apply theme: main process rewrites theme*.xml + per-page backgrounds then reparses, sending back the whole RenderSlide set.
// All element ids change (save→reopen); selection/edit state is cleared too.
export async function applyThemePreset(ctx: ActionCtx, preset: SlideThemePreset): Promise<void> {
  if (!ctx.slide) return
  ctx.markUnsaved()
  const r = await window.slidesApi.applyTheme({
    name: preset.name,
    colors: preset.colors,
    ...(preset.majorFont ? { majorFont: preset.majorFont } : {}),
    ...(preset.minorFont ? { minorFont: preset.minorFont } : {}),
    fitWidthPx: FIT_WIDTH,
  })
  if (r && !Array.isArray(r)) {
    ctx.setStatus(t('appStatusThemeApplyFailed', { error: r.error }))
    return
  }
  if (r) {
    ctx.setSlides(r)
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusThemeApplied', { name: preset.name }))
  }
}

/** Table style operations (delegated to IPC) */
export async function onEditTableStyle(
  ctx: ActionCtx,
  op: Omit<EditTableStyleOp, 'slideIndex' | 'sourceId'>,
): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'table') return
  const oldId = ctx.selectedNode.sourceId
  const result = await window.slidesApi.editTableStyle({
    ...op,
    slideIndex: ctx.current,
    sourceId: oldId,
  })
  if (result) {
    ctx.applySlide(ctx.current, result.slide)
    // Element ids change after reparse: re-select so the "Table Design" tab doesn't jump away
    if (result.sourceId) {
      ctx.setSelectedIds([result.sourceId])
      if (ctx.editingCell?.sourceId === oldId)
        ctx.setEditingCell({ ...ctx.editingCell, sourceId: result.sourceId })
    }
  }
}

/** Chart edit operations (delegated to IPC) */
export async function onEditChart(
  ctx: ActionCtx,
  op: Omit<EditChartOp, 'slideIndex' | 'sourceId'>,
): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'chart') return
  const result = await window.slidesApi.editChart({
    ...op,
    slideIndex: ctx.current,
    sourceId: ctx.selectedNode.sourceId,
  })
  if (result) {
    ctx.applySlide(ctx.current, result.slide)
    // Element ids change after reparse: re-select so the "Chart Design" tab doesn't jump away
    if (result.sourceId) ctx.setSelectedIds([result.sourceId])
  }
}

/** Open the chart data edit dialog */
export async function openChartDataDialog(ctx: ActionCtx): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'chart') return
  const data = await window.slidesApi.getChartData(ctx.current, ctx.selectedNode.sourceId)
  if (data) {
    ctx.setChartDataDialogInit(data)
    ctx.setChartDataDialogOpen(true)
  }
}
