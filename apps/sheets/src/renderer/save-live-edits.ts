/**
 * MoreAI autosave: dump the live Univer grid so disk matches what is on
 * screen after a soft-commit clears the journal.
 *
 * Snapshot `cell.s` is often a shared style-id (or a stale id after a
 * partial `{bl:1}` patch). Looking that id up and letting it overwrite the
 * journal is how header-bold landed on the data rows and disappeared from
 * the header. Prefer the composed cell style the ribbon reads, and never
 * let a live `bold:false` wipe a journaled `bold:true`.
 */
import type { WorkbookCellEdit } from '../shared/desktop-api'
import { concreteFontFamily, extractRichText, toNeutralStyle } from './edit-journal'
import type { UniverRuntime } from './univer-state'

type SnapshotCell = {
  v?: unknown
  f?: string
  s?: string | number | Record<string, unknown>
  p?: { body?: { dataStream?: string; textRuns?: unknown } }
}

type SnapshotSheet = {
  cellData?: Record<string, Record<string, SnapshotCell | null> | null>
}

export function isStyleFlagOn(value: unknown): boolean {
  return value === 1 || value === true
}

export function isStyleFlagOff(value: unknown): boolean {
  return value === 0 || value === false
}

export function univerStyleToEdit(
  style: Record<string, unknown> | undefined,
): WorkbookCellEdit['style'] | undefined {
  if (!style) return undefined
  const out: NonNullable<WorkbookCellEdit['style']> = {}
  if (isStyleFlagOn(style.bl)) out.bold = true
  if (isStyleFlagOff(style.bl)) out.bold = false
  if (isStyleFlagOn(style.it)) out.italic = true
  if (isStyleFlagOff(style.it)) out.italic = false
  if (style.ff && typeof style.ff === 'string') out.fontFamily = concreteFontFamily(style.ff)
  if (typeof style.fs === 'number' && style.fs > 0) out.fontSize = style.fs
  const ul = style.ul as { s?: number | boolean } | undefined
  if (ul && isStyleFlagOn(ul.s)) out.underline = true
  const st = style.st as { s?: number | boolean } | undefined
  if (st && isStyleFlagOn(st.s)) out.strikethrough = true
  const cl = style.cl as { rgb?: string } | undefined
  if (cl?.rgb && typeof cl.rgb === 'string' && cl.rgb.length > 0) {
    out.fontColor = cl.rgb.startsWith('#') ? cl.rgb : `#${cl.rgb}`
  }
  const bg = style.bg as { rgb?: string } | undefined
  if (bg?.rgb && typeof bg.rgb === 'string' && bg.rgb.length > 0) {
    out.fillColor = bg.rgb.startsWith('#') ? bg.rgb : `#${bg.rgb}`
  }
  if (style.tb === 3) out.wrapText = true
  if (Object.keys(out).length === 0) return undefined
  return out
}

export function resolveSnapshotStyle(
  raw: unknown,
  styleTable: Record<string, Record<string, unknown> | undefined>,
): Record<string, unknown> | undefined {
  if (typeof raw === 'string') return styleTable[raw]
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    return styleTable[String(raw)] ?? styleTable[raw]
  }
  if (typeof raw === 'object' && raw !== null) return raw as Record<string, unknown>
  return undefined
}

export function liveCellStyle(
  raw: unknown,
  styleTable: Record<string, Record<string, unknown> | undefined>,
): WorkbookCellEdit['style'] | undefined {
  const data = resolveSnapshotStyle(raw, styleTable)
  if (!data) return undefined
  const fromNeutral = toNeutralStyle(data)
  const fromUniver = univerStyleToEdit(data)
  if (!fromNeutral) return fromUniver
  if (!fromUniver) return fromNeutral
  return { ...fromNeutral, ...fromUniver }
}

/** Typed cells often store text only in `p`; `v` is empty until a later mutation. */
export function cellPlainValue(cell: SnapshotCell): string | number | boolean | undefined {
  const value = cell.v
  if (typeof value === 'string' && value.length > 0) return value
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'boolean') return value
  const stream = cell.p?.body?.dataStream
  if (typeof stream !== 'string' || stream.length === 0) return undefined
  const text = stream.replace(/\r\n$/, '').replace(/\r/g, '\n').replace(/\n$/, '')
  return text.length > 0 ? text : undefined
}

function firstRunFamily(rich: WorkbookCellEdit['rich'] | undefined): string | undefined {
  return rich?.find((run) => typeof run.family === 'string' && run.family.length > 0)?.family
}

type PaintedFlag = 'bold' | 'italic' | 'underline' | 'strikethrough'

/// Univer paints a `p` document from its runs, and only plain `v` cells from
/// the cell xf. A shared style-id can say the opposite of those runs (header
/// bold lands on the data xf; the header run stays regular). When every run
/// agrees, that flag is the format on screen.
function uniformRunFlag(
  rich: WorkbookCellEdit['rich'] | undefined,
  key: PaintedFlag,
): boolean | undefined {
  if (!rich?.length) return undefined
  const painted = rich[0]?.[key] === true
  return rich.every((run) => (run[key] === true) === painted) ? painted : undefined
}

/// Overlay the face the grid is actually painting onto the live style.
/// Journal keys still win later, so a real bold toggle is not undone by a
/// run that Univer has not patched yet.
export function applyPaintedRunStyle(edit: WorkbookCellEdit): WorkbookCellEdit {
  const rich = edit.rich
  if (!rich?.length) return edit
  const style: NonNullable<WorkbookCellEdit['style']> = { ...(edit.style ?? {}) }
  let changed = false
  for (const key of ['bold', 'italic', 'underline', 'strikethrough'] as const) {
    const painted = uniformRunFlag(rich, key)
    if (painted === undefined || style[key] === painted) continue
    style[key] = painted
    changed = true
  }
  if (!changed) return edit
  return { ...edit, style }
}

/// Pasted HTML paints from rich runs. Prefer that face over a journaled theme
/// xf (Aptos) that never described the glyphs on screen.
function preferRich(
  live: WorkbookCellEdit['rich'] | undefined,
  journal: WorkbookCellEdit['rich'] | undefined,
): WorkbookCellEdit['rich'] | undefined {
  if (firstRunFamily(live)) return live
  if (journal?.length) return journal
  return live?.length ? live : undefined
}

/**
 * Journal is the user's edit log and wins on overlapping style keys (a live
 * snapshot `bold:false` from a stale style-id must not un-bold the header).
 * Live fills in values the journal dropped and styles for cells the journal
 * never touched. An explicit live/run typeface wins over the journal: paste
 * often journals the theme face while the document runs hold the real one.
 */
export function mergeEditsPreferLive(
  journal: WorkbookCellEdit[],
  live: WorkbookCellEdit[],
): WorkbookCellEdit[] {
  const map = new Map<string, WorkbookCellEdit>()
  for (const edit of journal) {
    map.set(`${edit.sheetId}:${edit.row}:${edit.column}`, edit)
  }
  for (const edit of live) {
    const key = `${edit.sheetId}:${edit.row}:${edit.column}`
    const prev = map.get(key)
    const painted = applyPaintedRunStyle(edit)
    if (!prev) {
      map.set(key, painted)
      continue
    }
    const rich = preferRich(painted.rich, prev.rich)
    const pastedFont = firstRunFamily(rich) ?? painted.style?.fontFamily
    map.set(key, {
      ...prev,
      ...(painted.writeValue
        ? {
            writeValue: true,
            value: painted.value,
            ...(painted.formula !== undefined ? { formula: painted.formula } : { formula: prev.formula }),
          }
        : {}),
      ...(painted.style || prev.style || pastedFont
        ? {
            style: {
              ...painted.style,
              ...prev.style,
              ...(pastedFont ? { fontFamily: pastedFont } : {}),
            },
          }
        : {}),
      ...(rich ? { rich } : {}),
    })
  }
  return [...map.values()]
}

type StyleRange = {
  getCellStyleData(type?: 'row' | 'col' | 'cell'): Record<string, unknown> | null | undefined
}

type LiveSheet = {
  getRange(row: number, column: number): StyleRange
}

function readRangeStyle(
  worksheet: LiveSheet | null | undefined,
  row: number,
  column: number,
  type?: 'cell',
): Record<string, unknown> | undefined {
  if (!worksheet) return undefined
  try {
    const data = worksheet.getRange(row, column).getCellStyleData(type)
    return data && typeof data === 'object' ? data : undefined
  } catch {
    return undefined
  }
}

/// Cell-own style excludes the workbook default (often Arial) and row/column
/// styles. Composed style must not be written back: it replaces the file's
/// theme font and can stamp a shared xf's bold onto every cell.
function cellOwnStyle(
  worksheet: LiveSheet | null | undefined,
  row: number,
  column: number,
): Record<string, unknown> | undefined {
  return readRangeStyle(worksheet, row, column, 'cell')
}

export function collectLiveValueEdits(runtime: UniverRuntime | null): WorkbookCellEdit[] {
  try {
    const workbook = runtime?.univerAPI.getActiveWorkbook() as
      | {
          getId?(): string
          save?(): {
            styles?: Record<string, Record<string, unknown> | undefined>
            sheets?: Record<string, SnapshotSheet | null>
          }
          getSnapshot(): {
            styles?: Record<string, Record<string, unknown> | undefined>
            sheets?: Record<string, SnapshotSheet | null>
          }
          getSheetBySheetId?(id: string): LiveSheet | null
        }
      | null
      | undefined
    if (!workbook) return []
    if (workbook.getId?.() === 'new-workbook') return []
    const snapshot = workbook.save?.() ?? workbook.getSnapshot()
    const edits: WorkbookCellEdit[] = []
    for (const [sheetId, sheet] of Object.entries(snapshot.sheets ?? {})) {
      const cellData = sheet?.cellData ?? {}
      const worksheet = workbook.getSheetBySheetId?.(sheetId) ?? null
      for (const [rowKey, row] of Object.entries(cellData)) {
        const rowIndex = Number(rowKey)
        if (!Number.isInteger(rowIndex) || rowIndex < 0 || !row) continue
        for (const [colKey, cell] of Object.entries(row)) {
          const column = Number(colKey)
          if (!Number.isInteger(column) || column < 0 || !cell) continue
          const own = liveCellStyle(cellOwnStyle(worksheet, rowIndex, column), {})
          // A string style-id is shared. Resolving it stamps one cell's bold
          // onto every cell that still points at the id. Inline style objects
          // are already this cell's own delta.
          const inline =
            typeof cell.s === 'object' && cell.s !== null ? liveCellStyle(cell.s, {}) : undefined
          const rich = extractRichText(cell.p)?.runs
          const cellStyle = own ?? inline
          const explicitFont = firstRunFamily(rich) ?? cellStyle?.fontFamily
          let style = cellStyle
          if (explicitFont) style = { ...(style ?? {}), fontFamily: explicitFont }
          else if (style?.fontFamily) {
            const { fontFamily: _inherited, ...rest } = style
            style = Object.keys(rest).length > 0 ? rest : undefined
          }
          const painted = applyPaintedRunStyle({
            sheetId,
            row: rowIndex,
            column,
            writeValue: false,
            value: null,
            ...(style ? { style } : {}),
            ...(rich ? { rich } : {}),
          })
          style = painted.style
          const richPayload = rich ? { rich } : {}
          if (typeof cell.f === 'string' && cell.f.length > 0) {
            const formula = cell.f.startsWith('=') ? cell.f : `=${cell.f}`
            const cached = cellPlainValue(cell)
            edits.push({
              sheetId,
              row: rowIndex,
              column,
              writeValue: true,
              value: cached ?? formula,
              formula,
              ...(style ? { style } : {}),
              ...richPayload,
            })
            continue
          }
          const value = cellPlainValue(cell)
          if (value !== undefined) {
            edits.push({
              sheetId,
              row: rowIndex,
              column,
              writeValue: true,
              value,
              ...(style ? { style } : {}),
              ...richPayload,
            })
            continue
          }
          if (style || rich) {
            edits.push({
              sheetId,
              row: rowIndex,
              column,
              writeValue: false,
              value: null,
              ...(style ? { style } : {}),
              ...richPayload,
            })
          }
        }
      }
    }
    return edits
  } catch {
    return []
  }
}
