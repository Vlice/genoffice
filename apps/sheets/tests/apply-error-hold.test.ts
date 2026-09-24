import { describe, expect, it } from 'vitest'
import { holdApplyOutcome } from '../src/renderer/ai/apply-error-hold'

const fallback = 'transaction failed'

describe('holdApplyOutcome', () => {
  it('holds a failure while the agent is still busy', () => {
    expect(
      holdApplyOutcome({
        applyGen: 2,
        latestGen: 2,
        busy: true,
        pending: null,
        ok: false,
        reason: 'Invalid horizontal alignment: right',
        fallbackReason: fallback,
      }),
    ).toEqual({
      pending: 'Invalid horizontal alignment: right',
      surfaceNow: null,
    })
  })

  it('replaces the held failure with a later one', () => {
    expect(
      holdApplyOutcome({
        applyGen: 3,
        latestGen: 3,
        busy: true,
        pending: 'earlier',
        ok: false,
        reason: 'later',
        fallbackReason: fallback,
      }),
    ).toEqual({ pending: 'later', surfaceNow: null })
  })

  it('clears a held failure when a later apply succeeds', () => {
    expect(
      holdApplyOutcome({
        applyGen: 4,
        latestGen: 4,
        busy: true,
        pending: 'Invalid horizontal alignment: right',
        ok: true,
        fallbackReason: fallback,
      }),
    ).toEqual({ pending: null, surfaceNow: null })
  })

  it('surfaces the failure immediately when idle', () => {
    expect(
      holdApplyOutcome({
        applyGen: 1,
        latestGen: 1,
        busy: false,
        pending: null,
        ok: false,
        reason: 'workbook changed',
        fallbackReason: fallback,
      }),
    ).toEqual({ pending: 'workbook changed', surfaceNow: 'workbook changed' })
  })

  it('ignores a stale earlier batch so it cannot overwrite a later outcome', () => {
    expect(
      holdApplyOutcome({
        applyGen: 1,
        latestGen: 2,
        busy: true,
        pending: 'kept',
        ok: false,
        reason: 'stale',
        fallbackReason: fallback,
      }),
    ).toEqual({ pending: 'kept', surfaceNow: null })
  })

  it('falls back when the failure has no reason', () => {
    expect(
      holdApplyOutcome({
        applyGen: 1,
        latestGen: 1,
        busy: false,
        pending: null,
        ok: false,
        fallbackReason: fallback,
      }),
    ).toEqual({ pending: fallback, surfaceNow: fallback })
  })
})
