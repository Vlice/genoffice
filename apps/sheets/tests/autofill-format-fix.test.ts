import { describe, expect, it } from 'vitest'

import { stripAutofillStylePayload } from '../src/renderer/autofill-format-fix'

describe('stripAutofillStylePayload', () => {
  it('drops style from a CLEAR then SET autofill payload', () => {
    const cleared = stripAutofillStylePayload({
      2: { 7: { v: null, s: null, p: null, f: null, si: null, custom: null } },
    }) as Record<string, Record<string, Record<string, unknown>>>
    expect(cleared[2]?.[7]).toEqual({
      v: null,
      p: null,
      f: null,
      si: null,
      custom: null,
    })

    const copied = stripAutofillStylePayload({
      2: { 7: { v: 12033, t: 2, s: { bl: 0, tb: 1, bg: { rgb: '' } } } },
    }) as Record<string, Record<string, Record<string, unknown>>>
    expect(copied[2]?.[7]).toEqual({ v: 12033, t: 2 })
  })

  it('leaves a value-only cell untouched', () => {
    const payload = { 0: { 0: { v: 'keep' } } }
    expect(stripAutofillStylePayload(payload)).toEqual({ 0: { 0: { v: 'keep' } } })
  })
})
