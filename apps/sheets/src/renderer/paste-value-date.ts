/**
 * Paste Special → Values drops number formats. A date is stored as an Excel
 * serial, so 2026-09-06 lands as 46271. WPS keeps the date on screen: carry
 * the date/time pattern with the value.
 *
 * Two sources:
 * - the copied cell already has a date/time format (real date serial)
 * - an external paste parses date text ("2026-09-06") into that serial and
 *   throws the pattern away
 *
 * Plain numbers stay unformatted. In-app text dates stay text.
 */
import { IUniverInstanceService, ObjectMatrix, getNumfmtParseValueFilter, numfmt } from '@univerjs/core'
import {
  SetRangeValuesMutation,
  SetRangeValuesUndoMutationFactory,
  type ISetRangeValuesMutationParams,
} from '@univerjs/sheets'
import {
  ISheetClipboardService,
  PREDEFINED_HOOK_NAME_PASTE,
  virtualizeDiscreteRanges,
} from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

export interface PasteValueCell {
  readonly v?: unknown
  readonly s?: string | { n?: { pattern?: string } | null } | null
}

const temporalCache = new Map<string, boolean>()

export function isTemporalNumberPattern(pattern: string): boolean {
  let known = temporalCache.get(pattern)
  if (known === undefined) {
    try {
      known = (numfmt.getFormatInfo(pattern) as { isDate?: boolean }).isDate === true
    } catch {
      known = false
    }
    temporalCache.set(pattern, known)
  }
  return known
}

function patternOnStyle(
  style: PasteValueCell['s'],
  stylePatternById: ((id: string) => string | undefined) | undefined,
): string | undefined {
  if (!style) return undefined
  if (typeof style === 'string') return stylePatternById?.(style)
  const pattern = style.n?.pattern
  return typeof pattern === 'string' && pattern.length > 0 ? pattern : undefined
}

/// The number format paste-values must keep so a date serial still reads as a
/// date. Null when the cell is not a date or time.
export function temporalPatternToCarry(
  cell: PasteValueCell | null | undefined,
  options: {
    external: boolean
    stylePatternById?: (id: string) => string | undefined
  },
): string | null {
  if (!cell || cell.v == null || cell.v === '') return null
  const fromStyle = patternOnStyle(cell.s, options.stylePatternById)
  if (fromStyle && isTemporalNumberPattern(fromStyle)) return fromStyle
  if (!options.external || typeof cell.v !== 'string') return null
  const parsed = getNumfmtParseValueFilter(cell.v.trim())
  if (!parsed || typeof parsed.v !== 'number' || typeof parsed.z !== 'string') return null
  return isTemporalNumberPattern(parsed.z) ? parsed.z : null
}

export function installPasteValueDateFormat(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const clipboardService = injector.get(ISheetClipboardService)
  const instanceService = injector.get(IUniverInstanceService)
  return clipboardService.addClipboardHook({
    id: 'genoffice-paste-value-date-format',
    onPasteCells(pasteFrom, pasteTo, data, payload) {
      const empty = { undos: [], redos: [] }
      // Only values-only. A normal paste already copies the cell style.
      if (payload?.pasteType !== PREDEFINED_HOOK_NAME_PASTE.SPECIAL_PASTE_VALUE) return empty
      const { unitId, subUnitId, range } = pasteTo
      const { mapFunc } = virtualizeDiscreteRanges([range])
      const sourceStyles = pasteFrom
        ? instanceService.getUniverSheetInstance(pasteFrom.unitId)?.getStyles()
        : undefined
      const matrix = new ObjectMatrix()
      data.forValue((row, col, value) => {
        const pattern = temporalPatternToCarry(value, {
          external: pasteFrom == null,
          stylePatternById: (id) => sourceStyles?.get(id)?.n?.pattern,
        })
        if (!pattern) return
        const mapped = mapFunc(row, col)
        if (!Number.isInteger(mapped.row) || !Number.isInteger(mapped.col)) return
        matrix.setValue(mapped.row, mapped.col, { s: { n: { pattern } } })
      })
      if (matrix.getLength() === 0) return empty
      const params: ISetRangeValuesMutationParams = {
        unitId,
        subUnitId,
        cellValue: matrix.getMatrix(),
      }
      return {
        redos: [{ id: SetRangeValuesMutation.id, params }],
        undos: [
          {
            id: SetRangeValuesMutation.id,
            params: injector.invoke(SetRangeValuesUndoMutationFactory, params),
          },
        ],
      }
    },
  })
}
