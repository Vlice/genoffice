/**
 * Autofill (fill-handle / copy-down) must copy values without restyling the
 * sheet. Univer's default hook CLEARS the target with `{v:null,s:null}` then
 * SETS a getStyleByCell snapshot of the source — streamed cells carry a
 * complete xf (wrap/overflow sentinels, empty-rgb fills). That rewrite is
 * journaled as styleReset and paints as a format change across the table.
 *
 * While an auto-fill command is running, drop `s` from those mutations so
 * destination formatting stays put. "Fill formatting only" is left alone.
 */
import { AutoFillCommand, SetRangeValuesMutation } from '@univerjs/sheets'

type CellValueMatrix = Record<string, Record<string, Record<string, unknown> | null | undefined>>

let installed = false
let autofillDepth = 0

export function stripAutofillStylePayload(cellValue: unknown): unknown {
  if (typeof cellValue !== 'object' || cellValue === null) return cellValue
  const matrix = cellValue as CellValueMatrix
  const out: CellValueMatrix = {}
  for (const [rowKey, rowValue] of Object.entries(matrix)) {
    if (typeof rowValue !== 'object' || rowValue === null) continue
    const row: Record<string, Record<string, unknown> | null | undefined> = {}
    for (const [columnKey, cell] of Object.entries(rowValue)) {
      if (cell === null || cell === undefined) {
        row[columnKey] = cell
        continue
      }
      if (typeof cell !== 'object') continue
      const next = { ...cell }
      delete next.s
      row[columnKey] = next
    }
    out[rowKey] = row
  }
  return out
}

function copiesFormattingOnly(params: unknown): boolean {
  if (typeof params !== 'object' || params === null) return false
  return (params as { applyType?: unknown }).applyType === 'ONLY_FORMAT'
}

export function installAutofillFormatFix(): void {
  if (installed) return
  installed = true
  const originalFill = AutoFillCommand.handler
  AutoFillCommand.handler = async (accessor, params) => {
    const preserveDestFormat = !copiesFormattingOnly(params)
    if (preserveDestFormat) autofillDepth += 1
    try {
      return await originalFill(accessor, params)
    } finally {
      if (preserveDestFormat) autofillDepth = Math.max(0, autofillDepth - 1)
    }
  }
  const originalSet = SetRangeValuesMutation.handler
  SetRangeValuesMutation.handler = (accessor, params) => {
    if (autofillDepth > 0 && params && typeof params === 'object' && 'cellValue' in params) {
      const next = params as { cellValue?: unknown }
      if (next.cellValue !== undefined) {
        next.cellValue = stripAutofillStylePayload(next.cellValue)
      }
    }
    return originalSet(accessor, params)
  }
}
