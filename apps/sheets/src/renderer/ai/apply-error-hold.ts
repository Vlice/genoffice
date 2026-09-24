/**
 * Mid-run propose_operations failures are retried by the model. Chat/status
 * should not flash those intermediate errors — only the last failure of a
 * finished run (or an idle/deterministic apply) is shown.
 */
export type ApplyErrorHold = {
  readonly pending: string | null
  readonly surfaceNow: string | null
}

export function holdApplyOutcome(args: {
  applyGen: number
  latestGen: number
  busy: boolean
  pending: string | null
  ok: boolean
  reason?: string | undefined
  fallbackReason: string
}): ApplyErrorHold {
  if (args.applyGen !== args.latestGen) {
    return { pending: args.pending, surfaceNow: null }
  }
  if (args.ok) return { pending: null, surfaceNow: null }
  const reason = args.reason?.trim() ? args.reason : args.fallbackReason
  return { pending: reason, surfaceNow: args.busy ? null : reason }
}
