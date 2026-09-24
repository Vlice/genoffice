/**
 * Host contracts for embedding GenOffice editors in a non-Electron SPA (e.g. MoreAI).
 * Electron preload remains a valid implementation of the same shapes via window.* bridges.
 */

export type OfficeKind = 'doc' | 'sheet' | 'ppt' | 'markdown' | 'pdf'

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx' | 'md' | 'pdf'

export interface OfficeHost {
  kind: OfficeKind
  /** Logical document id (artifact id or path alias). */
  documentId: string
  readBytes(): Promise<Uint8Array>
  writeBytes(data: Uint8Array): Promise<void>
  /** Optional UTF-8 helpers for markdown. */
  readText?(): Promise<string>
  writeText?(text: string): Promise<void>
  pickFiles?(accept: string[]): Promise<File[]>
  getLanguage?(): Promise<string>
  getTheme?(): Promise<'light' | 'dark' | 'system'>
}

/** Aligns with xlsx-sidecar JSON-lines command surface used by sheets-main. */
export interface SheetsBackend {
  open(input: {
    bytesBase64: string
    locale?: string
  }): Promise<{ sessionId: string; meta?: unknown }>
  readRange(input: {
    sessionId: string
    sheetId: string
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
  }): Promise<unknown>
  readFormulaCells?(input: { sessionId: string; sheetId: string }): Promise<unknown>
  recalc?(input: { sessionId: string }): Promise<unknown>
  applyEdits?(input: { sessionId: string; edits: unknown[] }): Promise<unknown>
  save(input: { sessionId: string }): Promise<{ bytesBase64: string }>
  close(input: { sessionId: string }): Promise<void>
}

declare global {
  interface Window {
    /** Injected by MoreAI (or tests) before GenOffice renderer bootstrap. */
    __OFFICE_HOST__?: unknown
    __resolveOfficeHost?: () => void
  }
}

/** Wait until parent/host assigned window bridges (iframe or SPA inject). */
export async function waitForOfficeHostInjection(timeoutMs = 15_000): Promise<void> {
  if (typeof window === 'undefined') return
  const w = window as Window & {
    desktop?: unknown
    desktopApi?: unknown
    slidesApi?: unknown
    markdownApi?: unknown
    pdfApi?: unknown
  }
  if (
    window.__OFFICE_HOST__ ||
    w.desktop ||
    w.desktopApi ||
    w.slidesApi ||
    w.markdownApi ||
    w.pdfApi
  ) {
    return
  }
  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Office host injection timeout')), timeoutMs)
    window.__resolveOfficeHost = () => {
      clearTimeout(timer)
      resolve()
    }
  })
}
