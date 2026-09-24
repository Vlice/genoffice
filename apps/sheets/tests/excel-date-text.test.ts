import { describe, expect, it } from 'vitest'

import { textToExcelDateSerial, textToExcelTimeFraction } from '../src/domain/excel-date-text'

describe('textToExcelDateSerial', () => {
  it('parses ISO and slash dates to Excel serials', () => {
    expect(textToExcelDateSerial('2025-02-01')).toBe(45689)
    expect(textToExcelDateSerial('2025/2/1')).toBe(45689)
    expect(textToExcelDateSerial('2025-02-01 12:00:00')).toBeCloseTo(45689.5, 5)
  })

  it('parses Chinese 年月日 dates', () => {
    expect(textToExcelDateSerial('2025年2月1日')).toBe(45689)
  })

  it('parses unambiguous day-first or month-first slash dates', () => {
    // 13 Feb 2025 — both D/M/Y (day>12) and M/D/Y (day>12) resolve here
    expect(textToExcelDateSerial('13/2/2025')).toBe(45701)
    expect(textToExcelDateSerial('2/13/2025')).toBe(45701)
  })

  it('leaves ambiguous or non-date text alone', () => {
    expect(textToExcelDateSerial('1/2/2025')).toBeUndefined()
    expect(textToExcelDateSerial('ORD-2025')).toBeUndefined()
    expect(textToExcelDateSerial('')).toBeUndefined()
  })
})

describe('textToExcelTimeFraction', () => {
  it('parses clock times', () => {
    expect(textToExcelTimeFraction('12:00')).toBeCloseTo(0.5, 5)
    expect(textToExcelTimeFraction('00:00:00')).toBe(0)
    expect(textToExcelTimeFraction('6:00 PM')).toBeCloseTo(0.75, 5)
  })
})
