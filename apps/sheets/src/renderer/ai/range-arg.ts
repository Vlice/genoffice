/**
 * Models often omit `range`, nest it under `parameters`, or send `{start,end}`
 * instead of an A1 string. Reads should still land on the user's selection.
 */

export interface ResolvedToolRange {
  readonly range: string
  readonly sheetIdFromQualifier?: string
}

export interface RangeSheetRef {
  readonly id: string
  readonly name: string
}

const RANGE_KEYS = ['range', 'Range', 'address', 'a1', 'ref', 'cells'] as const

export function coerceA1(value: unknown): string {
  if (typeof value === 'string') return value.trim()
  if (Array.isArray(value)) {
    const parts = value.map((part) => (typeof part === 'string' ? part.trim() : '')).filter(Boolean)
    if (parts.length === 1) return parts[0]!
    if (parts.length >= 2) return `${parts[0]}:${parts[parts.length - 1]}`
    return ''
  }
  if (value && typeof value === 'object') {
    const obj = value as Record<string, unknown>
    for (const key of RANGE_KEYS) {
      const nested = coerceA1(obj[key])
      if (nested) return nested
    }
    const start = obj.start ?? obj.from ?? obj.startCell ?? obj.startAddress
    const end = obj.end ?? obj.to ?? obj.endCell ?? obj.endAddress
    if (typeof start === 'string' && start.trim()) {
      return typeof end === 'string' && end.trim() ? `${start.trim()}:${end.trim()}` : start.trim()
    }
  }
  return ''
}

function unwrapInput(input: Record<string, unknown>): Record<string, unknown> {
  for (const key of ['parameters', 'arguments', 'args', 'input'] as const) {
    const nested = input[key]
    if (!nested || typeof nested !== 'object' || Array.isArray(nested)) continue
    const hasOwn =
      RANGE_KEYS.some((name) => input[name] != null) ||
      input.sheetId != null ||
      input.addresses != null ||
      input.operations != null
    if (!hasOwn) return { ...input, ...(nested as Record<string, unknown>) }
  }
  return input
}

/** A1 range for a read tool. Empty `range` falls back to the send-time selection. */
export function resolveToolRange(
  input: Record<string, unknown>,
  selection: string | undefined,
  sheets: readonly RangeSheetRef[],
): ResolvedToolRange {
  const src = unwrapInput(input)
  let chosen = ''
  for (const key of RANGE_KEYS) {
    chosen = coerceA1(src[key])
    if (chosen) break
  }
  if (!chosen) chosen = (selection ?? '').trim()
  chosen = chosen.replace(/\$/g, '').trim()
  if (!chosen) return { range: '' }
  const bang = chosen.lastIndexOf('!')
  if (bang <= 0) return { range: chosen }
  const sheetName = chosen.slice(0, bang).trim().replace(/^'|'$/g, '')
  const range = chosen.slice(bang + 1).trim()
  if (!range) return { range: chosen }
  const match = sheets.find((sheet) => sheet.name === sheetName || sheet.id === sheetName)
  return match ? { range, sheetIdFromQualifier: match.id } : { range }
}
