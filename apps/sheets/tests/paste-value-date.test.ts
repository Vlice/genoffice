import { describe, expect, it } from 'vitest'

import { temporalPatternToCarry } from '../src/renderer/paste-value-date'

describe('temporalPatternToCarry', () => {
  it('keeps the pattern when paste-values parses a date string into a serial', () => {
    expect(temporalPatternToCarry({ v: '2026-09-06' }, { external: true })).toBe('yyyy-mm-dd')
    expect(temporalPatternToCarry({ v: '2026-09-01' }, { external: true })).toBe('yyyy-mm-dd')
    expect(temporalPatternToCarry({ v: '2026/9/6' }, { external: true })).toBe('yyyy/m/d')
    expect(temporalPatternToCarry({ v: '2026-09-01 12:00' }, { external: true })).toBe(
      'yyyy-mm-dd hh:mm',
    )
  })

  it('keeps time patterns the same way', () => {
    expect(temporalPatternToCarry({ v: '12:30:00' }, { external: true })).toBe('hh:mm:ss')
    expect(temporalPatternToCarry({ v: '12:30' }, { external: true })).toBe('hh:mm')
  })

  it('does not format plain numbers or identifiers', () => {
    expect(temporalPatternToCarry({ v: '299' }, { external: true })).toBeNull()
    expect(temporalPatternToCarry({ v: 'ORD-20260901-001' }, { external: true })).toBeNull()
    expect(temporalPatternToCarry({ v: 46271 }, { external: true })).toBeNull()
  })

  it('carries an existing date format on a numeric cell, including in-app paste', () => {
    expect(
      temporalPatternToCarry(
        { v: 46271, s: { n: { pattern: 'yyyy-mm-dd' } } },
        { external: false },
      ),
    ).toBe('yyyy-mm-dd')
  })

  it('leaves in-app date text as text', () => {
    expect(temporalPatternToCarry({ v: '2026-09-06' }, { external: false })).toBeNull()
  })

  it('resolves a style-pool id to its date pattern', () => {
    expect(
      temporalPatternToCarry(
        { v: 46271, s: 'style-1' },
        { external: false, stylePatternById: (id) => (id === 'style-1' ? 'yyyy-mm-dd' : undefined) },
      ),
    ).toBe('yyyy-mm-dd')
  })

  it('ignores a non-date format on the source cell', () => {
    expect(
      temporalPatternToCarry({ v: 299, s: { n: { pattern: '0.00' } } }, { external: false }),
    ).toBeNull()
  })
})
