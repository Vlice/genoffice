import type { AgentSkill, ExecutedToolCall } from './skill'
import type {
  AgentImage,
  AgentMessage,
  AgentStreamHandle,
  AgentToolCall,
  AgentToolResult,
  AgentTransport,
  ToolExecution,
} from './types'

export interface ToolExecutedEvent<TSnapshot> {
  call: AgentToolCall
  execution: ToolExecution
  /**
   * Snapshot captured just before this tool ran; present only on the first
   * mutating tool of a run (hook for one-click rollback UIs).
   */
  snapshotBefore?: TSnapshot | undefined
}

export interface AgentRunResult {
  /** final assistant text of the run ('' when cut off) */
  text: string
  cancelled: boolean
  /** true when maxTurns was reached; text is the partial answer from the no-tools finalizing turn */
  turnLimit: boolean
  /** the final turn hit the token limit (stop_reason max_tokens): text is incomplete; set only when true */
  truncated?: boolean
}

export interface AgentLoopEvents<TSnapshot> {
  /** cumulative assistant text of the current turn (call per delta) */
  onText?(text: string): void
  /** a tool is about to execute (UI shows a live "running" indicator; onToolExecuted always follows) */
  onToolStart?(call: AgentToolCall): void
  onToolExecuted?(event: ToolExecutedEvent<TSnapshot>): void
  /** a turn requested tools and they ran; the loop is going back to the model */
  onTurnEnd?(): void
  onDone?(result: AgentRunResult): void
  onError?(error: string): void
}

/** Context compaction config (budget tracked in UTF-8 bytes rather than message count) */
export interface CompactionOptions {
  /** History size that triggers compaction (UTF-8 bytes, default 256KB) */
  maxBytes?: number
  /** Size of recent messages kept after compaction (bytes, default 96KB, cut at a user boundary) */
  keepRecentBytes?: number
  /** Disable LLM summarization and use only the mechanical digest (for tests/offline) */
  disableLlmSummary?: boolean
}

export interface AgentLoopOptions<TSnapshot = unknown> {
  transport: AgentTransport
  skill: AgentSkill
  events?: AgentLoopEvents<TSnapshot>
  /** hard cap on model round-trips per run (default DEFAULT_MAX_TURNS) */
  maxTurns?: number
  /** history cap in messages, trimmed at user-turn boundaries (default 40) */
  maxHistory?: number
  /** Context compaction; false disables it (enabled by default with default thresholds) */
  compaction?: CompactionOptions | false
  /** capture rollback state; invoked right before tools run (see snapshotBefore) */
  captureSnapshot?(): TSnapshot
  /** wrap instruction + skill context into the user message text */
  formatUserMessage?(instruction: string, context: string): string
  /** appended to the system prompt each turn (e.g. reply-language directive following the UI language) */
  systemSuffix?(): string
}

const COMPACT_MAX_BYTES = 256 * 1024
const COMPACT_KEEP_RECENT_BYTES = 96 * 1024
/** Pre-truncation of each tool output in the summary request (the compaction request itself must not blow up on huge outputs) */
const SUMMARIZE_TOOL_OUTPUT_MAX = 2_000
const SUMMARIZE_TIMEOUT_MS = 30_000
/** When over budget mid-run, keep the last N tool messages verbatim and truncate earlier outputs to this length */
const STALE_TOOL_KEEP_RECENT = 2
const STALE_TOOL_OUTPUT_MAX = 1_000

/** Unified turn budget across the suite's chat panels (apps may still override per loop) */
export const DEFAULT_MAX_TURNS = 100

/** Cap on consecutive tool-input parse failures (a successful parse resets it); abort beyond it (keeps the model from burning turns on bad JSON) */
const MAX_INPUT_PARSE_RETRIES = 3

/**
 * Degenerate-loop guards. Weak models (BYOK/local endpoints especially) can
 * repeat the exact same turn forever or keep issuing failing tool calls; with
 * a large turn budget these must abort early instead of burning it.
 */
const MAX_IDENTICAL_TURNS = 3
/** Hard abort after this many consecutive all-error turns *after* the recovery nudge. */
const MAX_ALL_ERROR_TURNS = 3
/** Inject a recovery nudge after this many consecutive all-error turns (once per run). */
const ALL_ERROR_NUDGE_AFTER = 2

/**
 * Backoff schedule for in-place same-turn retries on empty-stream errors.
 * The "(empty stream)" suffix is a cross-layer contract with the ai-provider
 * protocols: the gateway closed the SSE stream without content, tool calls, or
 * message framing — a transient soft-failure. The turn produced nothing and
 * history is untouched, so re-sending the identical request is idempotent;
 * retrying here keeps one gateway hiccup from killing a long multi-tool run.
 */
const EMPTY_STREAM_RETRY_DELAYS_MS = [1_000, 3_000]

const TURN_LIMIT_NOTE =
  '[System] The tool-call turn limit for this request has been reached; no more tools may be called this turn. ' +
  'Answer directly from the information already gathered; if the task is unfinished, briefly state what is done and what remains.'

/**
 * One soft recovery before aborting an identical non-mutating loop. Read-only
 * audit/analysis runs (sheets AI 校验 / AI 分析) often re-issue the same
 * get_workbook_context + read_range turn; aborting immediately looks like a
 * hard product failure. Force one no-tools turn so the model must answer from
 * data already in history (weak endpoints ignore a text-only nudge while tools
 * remain available).
 */
const IDENTICAL_TURN_NUDGE =
  '[System] You repeated the exact same tool calls and got the same results without changing the workbook. ' +
  'No more tools may be called this turn. Reply to the user now in plain text using the information already gathered ' +
  '(for audit/analysis: list findings and fix suggestions; if nothing is wrong, say so clearly).'

/**
 * Soft recovery when every tool keeps failing (typical: image_search/web_search
 * backend down, or a weak model retrying a hallucinated tool). Unlike the
 * identical-turn nudge this still allows tools — the model should switch to a
 * working alternative instead of retrying the same failure.
 */
const ALL_ERROR_NUDGE =
  '[System] Every tool call in the last turns failed. Do not retry the same failing tools. ' +
  'Switch to a different tool that can complete the task, or answer the user in plain text from information already gathered. ' +
  'If you still cannot complete the task, you will be asked to tell the user what failed at the end — do not dump tool errors mid-run.'

/**
 * After the all-error recovery still fails, strip tools and force a user-facing
 * answer. Intermediate tool errors stay in history for the model; the UI should
 * only surface them with this final reply (not as a hard abort).
 */
const ALL_ERROR_FINALIZE_NOTE =
  '[System] Tools kept failing. No more tools may be called this turn. ' +
  'Reply to the user now in plain text, in the same language as the user, using any information already gathered. ' +
  'Mention tool failures only at the end of the answer, briefly. Do not retry tools.'

/** Same no-tools finish after unusable tool JSON, instead of aborting the run. */
const PARSE_FAIL_NOTE =
  '[System] Tool arguments were unusable (unparseable or truncated) repeatedly. ' +
  'No more tools may be called this turn. Reply to the user now in plain text: say what you already know and what could not be done.'

function isLoopSystemNote(text: string): boolean {
  return (
    text === TURN_LIMIT_NOTE ||
    text === IDENTICAL_TURN_NUDGE ||
    text === ALL_ERROR_NUDGE ||
    text.startsWith(ALL_ERROR_FINALIZE_NOTE) ||
    text.startsWith(PARSE_FAIL_NOTE)
  )
}

function toolErrorDigest(results: readonly { name: string; output: string; isError?: boolean }[]): string {
  const lines = results
    .filter((r) => r.isError)
    .map((r) => `- ${r.name}: ${String(r.output).replace(/\s+/g, ' ').slice(0, 300)}`)
  if (lines.length === 0) return ''
  return `Failed tools this turn:\n${lines.join('\n')}`.slice(0, 2_000)
}

/**
 * Terminal assistant text when tools mutated the artifact (or an edits-only
 * turn was restored) and the model returned no prose. Must be non-empty so
 * provider message converters never emit empty assistant content, which breaks
 * multi-turn follow-ups (see finishTurn / restore).
 * Exported so apps can substitute a localized / tool-derived summary in the UI.
 */
export const COMPLETED_VIA_TOOLS_TEXT = '(completed tool actions; no text reply)'

const SUMMARIZE_SYSTEM =
  'You are a conversation compressor. Compress this editing session between the user and the AI assistant into a concise summary so later turns can continue with context. ' +
  "Keep: the user's goals and key instructions, completed changes (which files/pages/elements were modified), important facts and data, and outstanding items. " +
  'For specific figures/statistics, mark their provenance: figures from the user or from tool results (e.g. web_search) keep their source; figures the assistant produced without a source must be marked "(unverified)" so later turns do not treat them as established facts. ' +
  'Omit: pleasantries, tool-call details, and intermediate trial and error. Use a bullet list of at most 400 words. Write the summary in the same language as the conversation. Output only the summary body, with no preamble.'

/** Prefix of the synthetic user message that carries the compacted-history summary */
const COMPACT_SUMMARY_PREFIX = '[Summary of earlier conversation'
const COMPACT_SUMMARY_HEADER = '[Summary of earlier conversation (auto-compacted)]'
const COMPACT_SUMMARY_ACK = 'Understood, continuing from the progress so far.'

/** Approximate UTF-8 byte count (ASCII 1 byte, CJK etc. 3; surrogate pairs count as 6 — slight overestimate is harmless) */
function utf8Size(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    n += c < 0x80 ? 1 : c < 0x800 ? 2 : 3
  }
  return n
}

/** Approximate byte cost of one message (text + tool inputs/outputs + image base64) */
function messageSize(m: AgentMessage): number {
  if (m.role === 'tool') {
    return m.results.reduce((n, r) => n + utf8Size(r.output) + 40, 0)
  }
  let n = utf8Size(m.text)
  if (m.role === 'user' && m.images) {
    n += m.images.reduce((s, img) => s + img.base64.length, 0)
  }
  if (m.role === 'assistant' && m.toolCalls) {
    for (const c of m.toolCalls) {
      try {
        n += utf8Size(JSON.stringify(c.input)) + 40
      } catch {
        n += 40
      }
    }
  }
  return n
}

function historySize(messages: readonly AgentMessage[]): number {
  return messages.reduce((n, m) => n + messageSize(m), 0)
}

/** Mechanical digest when LLM summarization is unavailable: bullet list of user instructions + final replies */
function mechanicalDigest(dropped: readonly AgentMessage[]): string {
  const lines: string[] = []
  for (const m of dropped) {
    if (m.role === 'user' && !m.text.startsWith(COMPACT_SUMMARY_PREFIX)) {
      lines.push(`- User: ${m.text.slice(0, 200)}`)
    } else if (m.role === 'assistant' && m.text && !m.toolCalls?.length) {
      lines.push(`  Reply: ${m.text.slice(0, 200)}`)
    }
  }
  return lines.join('\n').slice(0, 4_000) || '(earlier conversation omitted)'
}

/**
 * Generic ReAct loop: user message -> model turn (text + tool calls) ->
 * execute tools -> feed results back -> repeat until the model answers with
 * plain text. History persists across runs, so follow-up questions work.
 */
export class AgentLoop<TSnapshot = unknown> {
  private readonly options: AgentLoopOptions<TSnapshot>
  private history: AgentMessage[] = []
  private handle: AgentStreamHandle | null = null
  private running = false
  private cancelled = false
  private turns = 0
  /** Finalizing turn: no tools, let the model answer from what it has read */
  private finalizing = false
  /** True when finalizing because maxTurns was hit (drives onDone.turnLimit). Identical-nudge finalize stays false. */
  private finalizingFromTurnLimit = false
  private mutationSeen = false
  private inputParseFails = 0
  /** signature (text + tool calls) of the previous turn, for the identical-turn guard */
  private lastTurnSig = ''
  private identicalTurns = 0
  /** identical-turn nudge may fire once per run before a hard abort */
  private identicalNudgeUsed = false
  private allErrorTurns = 0
  /** all-error nudge may fire once per run before a hard abort */
  private allErrorNudgeUsed = false
  private turnStopReason: string | null = null
  private turnText = ''
  private turnReasoning = ''
  private toolCalls: AgentToolCall[] = []
  /** tools actually executed during this run, fed to skill.verifyResponse */
  private executedCalls: ExecutedToolCall[] = []
  /** verifyResponse may force one extra corrective turn per run — never more */
  private verifyRetryUsed = false
  /** user message of the in-flight run; a failed run rolls it (and everything after) back out of history */
  private runUserMsg: AgentMessage | null = null
  /** invalidates stale transport callbacks after cancel/reset */
  private generation = 0
  /** per-run abort: aborted on cancel(); long tools (e.g. generate_deck) use it to break internal loops */
  private abortController: AbortController | null = null

  constructor(options: AgentLoopOptions<TSnapshot>) {
    this.options = options
  }

  get busy(): boolean {
    return this.running
  }

  get messages(): readonly AgentMessage[] {
    return this.history
  }

  /**
   * Seed the conversation with restored history (e.g. transcript reloaded from
   * disk when a document reopens), so follow-up instructions keep their context.
   * No-op unless the loop is idle with an empty history.
   * Old messages over the compaction budget fold into a mechanical digest
   * (no LLM request on restore, guaranteeing zero latency).
   */
  restore(messages: readonly AgentMessage[]): void {
    if (this.running || this.history.length > 0 || messages.length === 0) return
    // Edits-only runs persist an assistant message with no text; give it a placeholder
    // so the turn stays paired and providers never see an empty assistant content block.
    // Turn-limit notes persisted by older builds are stripped: they are stale
    // directives ("no more tools may be called") that poison every later run.
    const normalized = messages
      .filter((m) => !(m.role === 'user' && isLoopSystemNote(m.text)))
      .map((m) =>
        m.role === 'assistant' && !m.text ? { ...m, text: COMPLETED_VIA_TOOLS_TEXT } : m,
      )
    // Unanswered user messages (a failed or interrupted run persisted them without a
    // reply) must not re-enter the model context: trailing ones would pair with the
    // next instruction as one turn, adjacent ones read as a combined instruction
    this.history = normalized.filter(
      (m, i) => m.role !== 'user' || (normalized[i + 1] && normalized[i + 1]!.role !== 'user'),
    )
    if (this.history.length === 0) return
    if (this.compactionEnabled()) {
      const { maxBytes, keepRecentBytes } = this.compactBudget()
      if (historySize(this.history) > maxBytes) {
        const cut = this.findCompactCut(keepRecentBytes)
        if (cut > 0) {
          const digest = mechanicalDigest(this.history.slice(0, cut))
          this.history = [
            { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${digest}` },
            { role: 'assistant', text: COMPACT_SUMMARY_ACK },
            ...this.history.slice(cut),
          ]
        }
      }
    }
    this.trimHistory()
  }

  /** images: inline attachments for this user turn (vision input; see AgentImage) */
  run(instruction: string, images?: AgentImage[]): void {
    if (this.running || !instruction) return
    this.running = true
    this.cancelled = false
    this.turns = 0
    this.finalizing = false
    this.finalizingFromTurnLimit = false
    this.mutationSeen = false
    this.inputParseFails = 0
    this.lastTurnSig = ''
    this.identicalTurns = 0
    this.identicalNudgeUsed = false
    this.allErrorTurns = 0
    this.allErrorNudgeUsed = false
    this.executedCalls = []
    this.verifyRetryUsed = false
    this.abortController = new AbortController()
    const context = this.options.skill.buildContext?.() ?? ''
    const format =
      this.options.formatUserMessage ??
      ((instr: string, ctx: string) => (ctx ? `${instr}\n\n${ctx}` : instr))
    const userMsg: AgentMessage = {
      role: 'user',
      text: format(instruction, context),
      ...(images?.length ? { images } : {}),
    }
    void this.beginRun(userMsg)
  }

  /** Compact (if needed), push the user message, then start the turn. Compaction failure doesn't block the run. */
  private async beginRun(userMsg: AgentMessage): Promise<void> {
    const generation = this.generation
    try {
      await this.maybeCompact()
    } catch {
      // Proceed with the run even if compaction fails (an over-budget history only costs more, it's still correct)
    }
    if (generation !== this.generation) return // reset during compaction
    if (this.cancelled) {
      this.running = false
      this.options.events?.onDone?.({ text: '', cancelled: true, turnLimit: false })
      return
    }
    // Leftover unanswered user message (a previous run failed before replying):
    // drop it so the model never sees two adjacent user turns as one combined instruction
    while (this.history.at(-1)?.role === 'user') this.history.pop()
    this.trimHistory()
    // Reasoning echo only matters inside a run's own tool loop; drop it from
    // finished runs so it stops costing tokens on every later request.
    this.history = this.history.map((m) =>
      m.role === 'assistant' && m.reasoning ? { ...m, reasoning: undefined } : m,
    )
    if (userMsg.role === 'user') {
      userMsg = { ...userMsg, text: sanitizeAgentPayload(userMsg.text) }
    }
    this.runUserMsg = userMsg
    this.history.push(userMsg)
    this.startTurn()
  }

  /**
   * A run failed: remove its user message and every message after it, so the
   * failed instruction can't be silently re-executed by the next run.
   */
  private rollbackFailedRun(): void {
    const msg = this.runUserMsg
    this.runUserMsg = null
    if (!msg) return
    const i = this.history.lastIndexOf(msg)
    if (i >= 0) this.history.splice(i)
  }

  // ── Context compaction: fold old conversation into a summary, keep recent messages verbatim ──

  private compactionEnabled(): boolean {
    return this.options.compaction !== false
  }

  private compactBudget(): { maxBytes: number; keepRecentBytes: number } {
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    return {
      maxBytes: opt?.maxBytes ?? COMPACT_MAX_BYTES,
      keepRecentBytes: opt?.keepRecentBytes ?? COMPACT_KEEP_RECENT_BYTES,
    }
  }

  /**
   * Find the compaction cut at a user boundary: accumulate from the tail up to keepRecentBytes.
   * Returns the start index of the kept segment; if no suitable boundary exists,
   * fall back to keeping the last user turn.
   */
  private findCompactCut(keepRecentBytes: number): number {
    let kept = 0
    let cut = -1
    for (let i = this.history.length - 1; i >= 0; i--) {
      kept += messageSize(this.history[i]!)
      if (kept > keepRecentBytes && cut >= 0) break
      if (this.history[i]!.role === 'user') cut = i
    }
    if (cut < 0) {
      for (let i = this.history.length - 1; i >= 0; i--) {
        if (this.history[i]!.role === 'user') return i
      }
    }
    return cut
  }

  private async maybeCompact(): Promise<void> {
    if (!this.compactionEnabled()) return
    const { maxBytes, keepRecentBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    const cut = this.findCompactCut(keepRecentBytes)
    if (cut <= 0) return // no foldable prefix
    const dropped = this.history.slice(0, cut)
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    let summary: string | null = null
    if (!opt?.disableLlmSummary) summary = await this.summarizeViaLlm(dropped)
    if (!summary) summary = mechanicalDigest(dropped)
    this.history = [
      { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${summary}` },
      { role: 'assistant', text: COMPACT_SUMMARY_ACK },
      ...this.history.slice(cut),
    ]
  }

  /** Hand the folded conversation to the model for a summary; returns null on failure/timeout (falls back to the mechanical digest). */
  private summarizeViaLlm(dropped: readonly AgentMessage[]): Promise<string | null> {
    // Slim down the summary request itself: pre-truncate tool outputs, strip images
    const slim: AgentMessage[] = dropped.map((m) => {
      if (m.role === 'tool') {
        return {
          role: 'tool' as const,
          results: m.results.map((r) => ({
            ...r,
            output: r.output.slice(0, SUMMARIZE_TOOL_OUTPUT_MAX),
          })),
        }
      }
      if (m.role === 'user' && m.images?.length) return { role: 'user' as const, text: m.text }
      return m
    })
    return new Promise((resolve) => {
      let text = ''
      let settled = false
      const finish = (v: string | null) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        resolve(v)
      }
      const timer = setTimeout(() => finish(null), SUMMARIZE_TIMEOUT_MS)
      try {
        // Attach to this.handle so cancel() can abort the summary request when the user clicks stop
        this.handle = this.options.transport.stream(
          {
            system: SUMMARIZE_SYSTEM,
            messages: [
              ...slim,
              { role: 'user', text: 'Compress the conversation above as instructed.' },
            ],
            tools: [],
          },
          {
            onDelta: (t) => {
              text += t
            },
            onToolCall: () => {
              /* the summary turn gets no tools */
            },
            onDone: () => finish(text.trim() || null),
            onError: () => finish(null),
          },
        )
      } catch {
        finish(null)
      }
    })
  }

  /**
   * When over budget mid-run (between tool turns), truncate stale tool outputs:
   * keep structure (tool_use/tool_result pairs intact), cut content only,
   * and keep the most recent N verbatim.
   */
  private squashStaleToolOutputs(): void {
    if (!this.compactionEnabled()) return
    const { maxBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    let recent = 0
    for (let i = this.history.length - 1; i >= 0; i--) {
      const m = this.history[i]!
      if (m.role !== 'tool') continue
      recent++
      if (recent <= STALE_TOOL_KEEP_RECENT) continue
      m.results = m.results.map((r) =>
        r.output.length > STALE_TOOL_OUTPUT_MAX
          ? {
              ...r,
              output: `${r.output.slice(0, STALE_TOOL_OUTPUT_MAX)}\n…(output truncated: too long)`,
            }
          : r,
      )
    }
  }

  cancel(): void {
    if (!this.running) return
    this.cancelled = true
    // abort lets long tools mid-execution (internal LLM loops etc.) stop promptly
    this.abortController?.abort()
    // the transport emits onDone after aborting, which finalizes the run
    this.handle?.cancel()
  }

  /** drop the conversation (e.g. when a different document is opened) */
  reset(): void {
    this.generation++
    this.abortController?.abort()
    this.handle?.cancel()
    this.handle = null
    this.running = false
    this.cancelled = false
    this.history = []
    this.runUserMsg = null
  }

  /** Runs at run boundaries only (restore / before a new user message): a long run's tail is all assistant/tool messages, and cutting mid-run would empty the request. */
  private trimHistory(): void {
    const max = this.options.maxHistory ?? 40
    if (this.history.length <= max) return
    // cut only at a user message so tool_use/tool_result pairs stay intact
    let i = this.history.length - max
    while (i < this.history.length && this.history[i]!.role !== 'user') i++
    if (i >= this.history.length) return // no user boundary in the window: keep history over budget
    const next = this.history.slice(i)
    if (this.runUserMsg && !next.includes(this.runUserMsg)) return
    this.history = next
  }

  /** Strip tools and ask the model for a user-facing answer (turn-limit / loop guards). */
  private beginNoToolsFinalize(note: string): void {
    this.finalizing = true
    this.finalizingFromTurnLimit = false
    this.lastTurnSig = ''
    this.identicalTurns = 0
    this.allErrorTurns = 0
    this.history.push({ role: 'user', text: note })
    this.squashStaleToolOutputs()
    this.options.events?.onTurnEnd?.()
    this.startTurn()
  }

  private startTurn(retriesUsed = 0): void {
    const generation = this.generation
    this.turnText = ''
    this.turnReasoning = ''
    this.toolCalls = []
    this.turnStopReason = null
    // Some transports emit an extra onDone after cancel — this turn may finalize only once
    let settled = false
    this.handle = this.options.transport.stream(
      {
        system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? ''),
        messages: [...this.history],
        tools: this.finalizing ? [] : this.options.skill.tools,
      },
      {
        onDelta: (text) => {
          if (generation !== this.generation || settled) return
          this.turnText += text
          this.options.events?.onText?.(this.turnText)
        },
        onReasoning: (text) => {
          if (generation !== this.generation || settled) return
          this.turnReasoning += text
        },
        onToolCall: (call) => {
          if (generation !== this.generation || settled) return
          this.toolCalls.push(call)
        },
        onStopReason: (reason) => {
          if (generation !== this.generation || settled) return
          this.turnStopReason = reason
        },
        onDone: () => {
          if (generation !== this.generation || settled) return
          settled = true
          void this.finishTurn()
        },
        onError: (error) => {
          if (generation !== this.generation || settled) return
          settled = true
          const delay = EMPTY_STREAM_RETRY_DELAYS_MS[retriesUsed]
          // The no-partial-output guard keeps the retry idempotent (an empty
          // stream never emits deltas, but a mislabeled error must not replay
          // a turn whose text/tool calls the UI already saw)
          if (
            delay !== undefined &&
            error.includes('(empty stream)') &&
            !this.cancelled &&
            !this.turnText &&
            this.toolCalls.length === 0
          ) {
            setTimeout(() => {
              if (generation !== this.generation) return
              // Stopped during the backoff window: finalize like a normal cancel
              if (this.cancelled) {
                void this.finishTurn()
                return
              }
              this.startTurn(retriesUsed + 1)
            }, delay)
            return
          }
          this.running = false
          this.rollbackFailedRun()
          this.options.events?.onError?.(error)
        },
      },
    )
  }

  private async finishTurn(): Promise<void> {
    const { events, skill, captureSnapshot } = this.options
    const toolCalls = this.toolCalls

    // Claimed-action guard: before accepting a final text turn, let the skill
    // check the claims in it against the tools that actually ran this run.
    // A returned correction forces one more model turn (tools stay available,
    // so the model can perform the missing action or reword its claim).
    if (toolCalls.length === 0 && !this.cancelled && !this.finalizing) {
      // snapshot copy: the live array keeps growing if the corrective turn
      // runs more tools, and the hook must see the state at check time
      const correction =
        !this.verifyRetryUsed && this.turnText && skill.verifyResponse
          ? skill.verifyResponse(this.turnText, [...this.executedCalls])
          : null
      if (correction) {
        this.verifyRetryUsed = true
        this.history.push({ role: 'assistant', text: this.turnText })
        this.history.push({ role: 'user', text: correction })
        // No onTurnEnd here: UIs use it to seal the current assistant bubble,
        // which would keep the rejected claim visible. Without it, the
        // corrective turn's cumulative onText overwrites the bubble in place.
        this.startTurn()
        return
      }
    }

    // final turn: no tools requested, the user stopped the run, or the
    // no-tools finalizing turn after hitting the limit
    // (a cancelled turn drops its tool calls — no results would follow)
    if (toolCalls.length === 0 || this.cancelled || this.finalizing) {
      // Strip one-shot system notes so they do not poison the next user run.
      if (this.finalizing) {
        for (let i = this.history.length - 1; i >= 0; i--) {
          const m = this.history[i]!
          if (m.role === 'user' && isLoopSystemNote(m.text)) {
            this.history.splice(i, 1)
            break
          }
        }
      }
      // Models often end a tool-using run with an empty text turn ("I'm done").
      // Leaving assistant text empty in history then poisons the next user
      // prompt: Anthropic rejects empty content arrays, Gemini rejects empty
      // parts, and OpenAI-compatible routes send content:null with no tool_calls —
      // all of which make follow-up turns fail or return empty again (see
      // genoffice#12 / #22: first prompt works, second shows "no summary").
      // Same normalization as restore(), applied unconditionally: cancelled and
      // read-only empty turns poison follow-ups just the same. onDone still
      // reports the raw turn text so app UIs keep their localized fallbacks
      // instead of surfacing this English placeholder.
      this.history.push({ role: 'assistant', text: this.turnText || COMPLETED_VIA_TOOLS_TEXT })
      this.running = false
      this.runUserMsg = null
      events?.onDone?.({
        text: this.turnText,
        cancelled: this.cancelled,
        turnLimit: this.finalizingFromTurnLimit,
        // set only when true so exact-shape consumers/tests stay unaffected
        ...(this.turnStopReason === 'max_tokens' && !this.cancelled ? { truncated: true } : {}),
      })
      return
    }

    // Strip turn-local execution hints (inputError/truncated) from the stored
    // history: they are not model context, and transports with strict message
    // schemas (the Electron IPC bridge) reject unknown tool-call keys when the
    // history is echoed back on the next turn. The OpenAI-compatible stream
    // paths attach `inputError: undefined` on every parsed call, so without
    // this the second turn of any custom-provider agent run fails validation.
    this.history.push({
      role: 'assistant',
      text: this.turnText,
      toolCalls: toolCalls.map(({ id, name, input }) => ({ id, name, input })),
      // interleaved-thinking models degrade in tool loops unless their reasoning is echoed back
      ...(this.turnReasoning ? { reasoning: this.turnReasoning } : {}),
    })
    const generation = this.generation
    const results: AgentToolResult[] = []
    let turnMutated = false
    for (const call of toolCalls) {
      // The user hit stop while an earlier tool was running: skip remaining tools,
      // but fill in paired error results to keep tool_use/tool_result pairs valid for the next request
      if (this.cancelled) {
        results.push({
          id: call.id,
          name: call.name,
          output: '(the user stopped the run; this tool was not executed)',
          isError: true,
        })
        continue
      }
      // Unusable input (truncated by the token limit, or JSON that failed to parse):
      // don't execute; feed a targeted error back so the model retries correctly
      if (call.truncated || call.inputError) {
        this.inputParseFails++
        const output = call.truncated
          ? 'Tool arguments were cut off by the output length limit; the tool was not executed. Split this operation into several smaller tool calls (less content per call) and try again.'
          : `Tool input JSON failed to parse; the tool was not executed: ${call.inputError}\nFix the arguments (make sure quotes inside strings are escaped) and call again.`
        results.push({ id: call.id, name: call.name, output, isError: true })
        events?.onToolExecuted?.({
          call,
          execution: { output, isError: true, summary: call.name },
        })
        continue
      }
      this.inputParseFails = 0
      events?.onToolStart?.(call)
      const snapshot = !this.mutationSeen ? captureSnapshot?.() : undefined
      let execution: ToolExecution
      try {
        execution = await skill.executeTool(call, this.abortController?.signal)
      } catch (e) {
        execution = {
          output: e instanceof Error ? e.message : String(e),
          isError: true,
          summary: call.name,
        }
      }
      if (generation !== this.generation) return // reset while a tool was running
      this.executedCalls.push({ name: call.name, ok: !execution.isError })
      const firstMutation = !!execution.mutated && !this.mutationSeen
      if (execution.mutated) {
        this.mutationSeen = true
        turnMutated = true
      }
      results.push({
        id: call.id,
        name: call.name,
        output: execution.output,
        isError: execution.isError,
      })
      events?.onToolExecuted?.({
        call,
        execution,
        snapshotBefore: firstMutation ? snapshot : undefined,
      })
    }
    this.history.push({ role: 'tool', results })

    // Cancelled while tools were executing: finish immediately, no further model request
    if (this.cancelled) {
      this.running = false
      this.runUserMsg = null
      events?.onDone?.({ text: this.turnText, cancelled: true, turnLimit: false })
      return
    }

    // Bad-input retries hit the cap: finish with a text answer instead of aborting
    if (this.inputParseFails >= MAX_INPUT_PARSE_RETRIES) {
      this.beginNoToolsFinalize(PARSE_FAIL_NOTE)
      return
    }

    // A turn where every tool call failed makes no progress; a long streak
    // (unknown-tool loops from malformed BYOK streams, hallucinated tools)
    // would otherwise burn the whole turn budget re-erroring.
    this.allErrorTurns = results.every((r) => r.isError) ? this.allErrorTurns + 1 : 0
    if (this.allErrorTurns >= ALL_ERROR_NUDGE_AFTER && !this.allErrorNudgeUsed) {
      this.allErrorNudgeUsed = true
      this.allErrorTurns = 0
      this.lastTurnSig = ''
      this.history.push({ role: 'user', text: ALL_ERROR_NUDGE })
      this.squashStaleToolOutputs()
      events?.onTurnEnd?.()
      this.startTurn()
      return
    }
    if (this.allErrorTurns >= MAX_ALL_ERROR_TURNS) {
      const digest = toolErrorDigest(results)
      this.beginNoToolsFinalize(
        digest ? `${ALL_ERROR_FINALIZE_NOTE}\n${digest}` : ALL_ERROR_FINALIZE_NOTE,
      )
      return
    }

    // Identical-turn guard: a model (typically a weak BYOK/local endpoint)
    // re-emitting the same text, tool calls AND tool outputs is looping, not
    // progressing. Turns that mutated the artifact are exempt (repeating an
    // identical edit is legitimate progress), and changing outputs break the
    // streak (so poll-style tools survive).
    const turnSig = JSON.stringify([
      this.turnText,
      toolCalls.map(({ name, input }) => [name, input]),
      results.map((r) => r.output),
    ])
    if (turnSig === this.lastTurnSig && !turnMutated) {
      if (++this.identicalTurns >= MAX_IDENTICAL_TURNS) {
        // Soft recovery once: strip tools and force a text answer from data
        // already in history (sheets AI 校验/分析). Hard-abort only if that
        // recovery path is somehow skipped (identicalNudgeUsed already set).
        if (!this.identicalNudgeUsed) {
          this.identicalNudgeUsed = true
          this.identicalTurns = 0
          this.lastTurnSig = ''
          this.finalizing = true
          this.finalizingFromTurnLimit = false
          this.history.push({ role: 'user', text: IDENTICAL_TURN_NUDGE })
          this.squashStaleToolOutputs()
          events?.onTurnEnd?.()
          this.startTurn()
          return
        }
        this.beginNoToolsFinalize(IDENTICAL_TURN_NUDGE)
        return
      }
    } else {
      this.lastTurnSig = turnSig
      this.identicalTurns = 0
    }

    this.turns++
    if (this.turns >= (this.options.maxTurns ?? DEFAULT_MAX_TURNS)) {
      // Don't throw away the context already gathered: append one no-tools turn for a partial answer
      this.finalizing = true
      this.finalizingFromTurnLimit = true
      this.history.push({ role: 'user', text: TURN_LIMIT_NOTE })
    }
    // Long runs (e.g. page-by-page generation) over budget mid-way: truncate stale tool outputs so each turn doesn't resend a huge payload
    this.squashStaleToolOutputs()
    events?.onTurnEnd?.()
    this.startTurn()
  }
}

/**
 * Redact secret-looking tokens from an outgoing user message so accidentally
 * pasted API keys, URL credentials, and password assignments don't reach
 * remote model APIs verbatim.
 *
 * Imported from public PR #32 (BuiltByHarshil), with the credential pattern
 * narrowed to URL userinfo (scheme://user:pass@host) so ordinary "a:b@c"
 * prose is never rewritten.
 */
export function sanitizeAgentPayload(payload: string): string {
  return payload
    .replace(/\b(?:sk-|AIza|ghp_|secret_)[A-Za-z0-9_-]{16,}/g, '[REDACTED_API_KEY]')
    .replace(/([a-z][a-z0-9+.-]*:\/\/[^\s:@/]+):[^\s@/]+@/gi, '$1:[REDACTED_CREDENTIALS]@')
    .replace(
      /(password|passwd|secret_key|private_key)(\s*[:=]\s*)["'][^"']+["']/gi,
      '$1$2"[REDACTED_SECURE_TOKEN]"',
    )
}
