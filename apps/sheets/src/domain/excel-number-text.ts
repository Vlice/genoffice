/**
 * Coerce plain numeric text to numbers so SUM/AVERAGE and number formats work.
 * Leaves long digit IDs and leading-zero codes as text.
 */

export function textToNumber(text: string): number | undefined {
  const trimmed = text
    .replace(/\u00a0/g, ' ')
    .trim()
  if (!trimmed) return undefined
  // Long digit runs / leading-zero codes stay text (phones, order ids, ZIP+).
  if (/^\d{7,}$/.test(trimmed)) return undefined
  if (/^0\d+$/.test(trimmed)) return undefined

  let working = trimmed
  let negative = false
  const paren = /^\((.+)\)$/.exec(working)
  if (paren) {
    working = paren[1]!.trim()
    negative = true
  }
  const hadPercent = /%$/.test(working)
  const cleaned = working
    .replace(/[\uFF0C,]/g, '') // fullwidth + ASCII thousands separators
    .replace(/^[¥￥$€£]\s?/, '')
    .replace(/\s/g, '')
    .replace(/%$/, '')
  if (!/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(cleaned)) return undefined
  let n = Number(cleaned)
  if (!Number.isFinite(n)) return undefined
  if (negative) n = -n
  return hadPercent ? n / 100 : n
}
