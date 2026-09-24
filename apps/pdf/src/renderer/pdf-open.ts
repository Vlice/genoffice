import { getDocument, type PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf.mjs'

/** Default open budget — module workers that never emit ready/error hang forever without this. */
export const PDF_OPEN_TIMEOUT_MS = 20_000

/**
 * MoreAI embeds `?mode=tab` inside a parent sandboxed iframe. Chromium module Workers
 * can stall there (Worker constructed, neither `error` nor pdf.js `test`/`ready`), so
 * `getDocument().promise` never settles and the UI stays on「正在打开…」.
 *
 * Fail Worker construction for pdf.worker so pdf.js falls through to FakeWorker
 * (main-thread MessageHandler) — still full GenOffice PDF rendering, not a native fallback.
 */
export function ensurePdfWorkerForEmbed(): void {
  if (typeof window === 'undefined') return
  const params = new URLSearchParams(window.location.search)
  if (params.get('mode') !== 'tab') return
  if ((window as unknown as { __MOREAI_PDF_WORKER_PATCHED__?: boolean }).__MOREAI_PDF_WORKER_PATCHED__) {
    return
  }
  ;(window as unknown as { __MOREAI_PDF_WORKER_PATCHED__?: boolean }).__MOREAI_PDF_WORKER_PATCHED__ = true

  const OrigWorker = window.Worker
  window.Worker = class MoreAiTabEmbedPdfWorker extends OrigWorker {
    constructor(scriptURL: string | URL, options?: WorkerOptions) {
      const src = String(scriptURL)
      if (options?.type === 'module' && /pdf\.worker/i.test(src)) {
        throw new Error('pdf.js module worker disabled in MoreAI tab embed (use FakeWorker)')
      }
      super(scriptURL, options)
    }
  } as typeof Worker
}

/** getDocument with destroy-on-timeout so open never hangs unboundedly. */
export async function getDocumentWithTimeout(
  src: Parameters<typeof getDocument>[0],
  timeoutMs = PDF_OPEN_TIMEOUT_MS,
): Promise<PDFDocumentProxy> {
  const task = getDocument(src)
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    return await Promise.race([
      task.promise,
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          void task.destroy().catch(() => undefined)
          reject(new Error(`PDF open timed out after ${timeoutMs}ms`))
        }, timeoutMs)
      }),
    ])
  } finally {
    if (timer !== undefined) clearTimeout(timer)
  }
}
