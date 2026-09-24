/**
 * Delete / Backspace on a multi-cell selection must clear every selected
 * cell (Excel / Google Sheets). Univer 0.25 binds Backspace to
 * "delete-and-start-editing" (SetCellEditVisibleOperation) which only
 * affects the active cell, and on some hosts that same binding races
 * Delete. Intercept at window-capture — before Univer's shortcut
 * dispatcher — and route to sheet.command.clear-selection-content.
 *
 * Univer parks grid focus on a hidden contenteditable host, so a naive
 * "skip every contentEditable" check would never intercept the grid.
 * Native inputs (including Univer find/replace and rule panels) always
 * skip; remaining contenteditables skip unless they sit in the sheet
 * canvas rather than Univer chrome or app chrome.
 *
 * Sheet rename is a contenteditable span inside the sheet bar (still
 * under #univer-container). Backspace/Delete must edit that label, not
 * clear the grid selection.
 */

export const CLEAR_SELECTION_CONTENT_COMMAND = 'sheet.command.clear-selection-content'

const NATIVE_FIELD_SELECTOR = 'input, textarea, select'
const CONTENT_EDITABLE_SELECTOR = '[contenteditable="true"]'
const SHEET_CONTAINER_SELECTOR = '#univer-container'
export const SKIP_HOST_SELECTOR = [
  '[data-u-comp="formula-bar"]',
  '[data-u-comp="input"]',
  '[data-u-comp="textarea"]',
  '[data-u-comp="panel"]',
  '[data-u-comp="panel-field"]',
  '[data-u-comp="cell-popup"]',
  '[data-u-comp="defined-name"]',
  '[data-u-comp="defined-name-container"]',
  '[data-u-comp="select"]',
  '[data-u-comp="multiple-select"]',
  '[data-u-comp="sheets-dropdown-list"]',
  '[data-u-comp="gallery"]',
  '.shape-editable',
  '.chart-editor',
  '.dialog-backdrop',
  '[role="dialog"]',
  '[data-u-comp="slide-tab-item"]',
].join(', ')

export function isClearSelectionHotkey(
  event: Pick<ClearSelectionKeyEvent, 'key' | 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey'>,
): boolean {
  if (event.key !== 'Delete' && event.key !== 'Backspace') return false
  return !event.metaKey && !event.ctrlKey && !event.altKey && !event.shiftKey
}

export interface ClearSelectionKeyEvent {
  readonly key: string
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly shiftKey: boolean
  readonly defaultPrevented: boolean
  readonly isComposing: boolean
  readonly target: { closest(selector: string): unknown } | EventTarget | null
}

export function shouldInterceptClearSelection(
  event: ClearSelectionKeyEvent,
  isCellEditing: boolean,
): boolean {
  if (!isClearSelectionHotkey(event)) return false
  if (event.defaultPrevented) return false
  if (event.isComposing) return false
  if (isCellEditing) return false
  const target = event.target
  if (!hasClosest(target)) return true
  if (target.closest(SKIP_HOST_SELECTOR)) return false
  // Find/replace, data-validation, CF, and app chrome all use native fields.
  if (target.closest(NATIVE_FIELD_SELECTOR)) return false
  // App chrome (AI composer) is contenteditable outside the sheet container.
  if (target.closest(CONTENT_EDITABLE_SELECTOR) && !target.closest(SHEET_CONTAINER_SELECTOR)) {
    return false
  }
  return true
}

function hasClosest(value: ClearSelectionKeyEvent['target']): value is Element {
  return value != null && typeof (value as Element).closest === 'function'
}

/** Invisible character that keeps a caret line box after the label is cleared. */
export const SHEET_RENAME_CARET = '\u200b'

export function sheetRenameVisibleText(text: string): string {
  return text.replaceAll(SHEET_RENAME_CARET, '').replaceAll('\n', '')
}

/** An empty rename label has no line box, so the caret disappears. */
export function sheetRenameNeedsCaret(text: string): boolean {
  return sheetRenameVisibleText(text) === ''
}

const SHEET_RENAME_LABEL = '[data-u-comp="slide-tab-item"] [contenteditable="true"]'

function sheetRenameLabel(target: EventTarget | null): HTMLElement | null {
  if (target == null || typeof (target as Element).closest !== 'function') return null
  const el = (target as Element).closest(SHEET_RENAME_LABEL)
  return el instanceof HTMLElement ? el : null
}

function parkSheetRenameCaret(doc: Document, el: HTMLElement): void {
  if (!sheetRenameNeedsCaret(el.textContent ?? '')) return
  if (el.textContent !== SHEET_RENAME_CARET) el.textContent = SHEET_RENAME_CARET
  const node = el.firstChild
  if (node == null) return
  const range = doc.createRange()
  range.setStart(node, node.textContent?.length ?? 0)
  range.collapse(true)
  const selection = doc.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)
}

/** Keep a caret in a cleared sheet tab, and strip that placeholder before the name is saved. */
export function bindSheetRenameCaret(doc: Document): () => void {
  const onBeforeInput = (event: Event): void => {
    const input = event as InputEvent
    if (input.isComposing) return
    const el = sheetRenameLabel(event.target)
    if (el == null || !sheetRenameNeedsCaret(el.textContent ?? '')) return
    if (input.inputType !== 'deleteContentBackward' && input.inputType !== 'deleteContentForward') return
    event.preventDefault()
    parkSheetRenameCaret(doc, el)
  }
  const onInput = (event: Event): void => {
    if ((event as InputEvent).isComposing) return
    const el = sheetRenameLabel(event.target)
    if (el == null || !sheetRenameNeedsCaret(el.textContent ?? '')) return
    parkSheetRenameCaret(doc, el)
  }
  const onFocusOut = (event: FocusEvent): void => {
    const el = sheetRenameLabel(event.target)
    if (el == null) return
    const name = sheetRenameVisibleText(el.textContent ?? '')
    if ((el.textContent ?? '') !== name) el.textContent = name
    if (name !== '') return
    setTimeout(() => {
      if (el.getAttribute('contenteditable') === 'true' && sheetRenameNeedsCaret(el.textContent ?? '')) {
        parkSheetRenameCaret(doc, el)
      }
    }, 0)
  }
  doc.addEventListener('beforeinput', onBeforeInput, true)
  doc.addEventListener('input', onInput, true)
  doc.addEventListener('focusout', onFocusOut, true)
  return () => {
    doc.removeEventListener('beforeinput', onBeforeInput, true)
    doc.removeEventListener('input', onInput, true)
    doc.removeEventListener('focusout', onFocusOut, true)
  }
}

export type SheetTabRenameCapture = 'release' | 'passthrough'

/**
 * Sheet rename is a contenteditable span. Univer's shortcut service listens
 * on the same window-capture phase and cancels Backspace plus character
 * keys. `release` stops the rest of that phase without preventDefault, so
 * the browser edits the label. Enter stays a passthrough: the tab commits
 * by blurring on Enter.
 */
export function sheetTabRenameCapture(
  event: Pick<ClearSelectionKeyEvent, 'key' | 'target'>,
): SheetTabRenameCapture | null {
  const target = event.target
  if (!hasClosest(target)) return null
  if (!target.closest('[data-u-comp="slide-tab-item"]')) return null
  if (!target.closest(CONTENT_EDITABLE_SELECTOR)) return null
  if (event.key === 'Enter') return 'passthrough'
  return 'release'
}
