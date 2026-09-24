import { describe, expect, it } from 'vitest'

import { textToNumber } from '../src/domain/excel-number-text'

describe('textToNumber', () => {
  it('parses plain and currency amounts', () => {
    expect(textToNumber('598')).toBe(598)
    expect(textToNumber('1,299.5')).toBe(1299.5)
    expect(textToNumber('¥100')).toBe(100)
    expect(textToNumber('￥100')).toBe(100)
    expect(textToNumber('50%')).toBe(0.5)
    expect(textToNumber('(1,299)')).toBe(-1299)
    expect(textToNumber('1，299.5')).toBe(1299.5)
  })

  it('keeps identifiers as text', () => {
    expect(textToNumber('13800138000')).toBeUndefined()
    expect(textToNumber('007')).toBeUndefined()
    expect(textToNumber('ORD-1')).toBeUndefined()
  })
})
