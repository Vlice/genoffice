/**
 * Unified dropdown/popover dismissal (the slides PR #505 model, generalized):
 *
 *  1. a press anywhere outside the popover closes it
 *  2. window blur closes it (app/window switch)
 *  3. a press on the shell tab strip — a sibling WebContentsView whose input
 *     never reaches this document — closes it via the app:chrome-pressed IPC
 *     relay (the shell main process also broadcasts it on window drag)
 *
 * Apps whose ribbon tab row doubles as the frameless-window drag region must
 * additionally suspend that region while a popover is open (drag regions
 * swallow mouse events, so no listener can see presses on the blank band).
 * While any popover installed here is open, `<html>` carries the
 * `genoffice-popover-open` class — suspend the drag region with
 * `html.genoffice-popover-open .your-drag-row { -webkit-app-region: no-drag; }`.
 */
import { useEffect, useRef } from 'react'

type ChromePressedApi = { onChromePressed?: (handler: () => void) => () => void }

/** Each app's preload exposes the app:chrome-pressed subscription under its
 * own namespace; probe the known ones so callers never need to care. */
function subscribeChromePressed(handler: () => void): (() => void) | undefined {
  const w = window as unknown as Record<string, ChromePressedApi | undefined>
  for (const name of [
    'slidesApi',
    'desktopApi',
    'desktop',
    'pdfApi',
    'markdownApi',
    'aiOfficeTabs',
  ]) {
    const sub = w[name]?.onChromePressed
    if (typeof sub === 'function') return sub.call(w[name], handler)
  }
  return undefined
}

/** open-popover refcount driving the html-level drag-region suspension class */
let openPopovers = 0
function bumpOpenPopovers(delta: 1 | -1): void {
  openPopovers = Math.max(0, openPopovers + delta)
  document.documentElement.classList.toggle('genoffice-popover-open', openPopovers > 0)
}

/** Flag while a native <input type="color"> dialog may be open (window blur
 *  / outside pointerdown must not dismiss the host palette — that would unmount
 *  the input and kill the OS color disk mid-pick). Chromium fires `change` on
 *  the first disk click (live), so we keep a timestamp grace instead of
 *  disarming on every change — disarming there is what folded「其他颜色…」. */
let nativeColorDialogArm = 0
let nativeColorDialogAt = 0
const NATIVE_COLOR_GRACE_MS = 750

/** Call from color-input focus / pointerdown / input before the OS picker opens. */
export function armNativeColorDialog(): void {
  nativeColorDialogArm = 1
  nativeColorDialogAt = performance.now()
}

/** Call when the native color picker session ends (host palette closed, or the
 *  window regained focus after the OS dialog dismissed). */
export function disarmNativeColorDialog(): void {
  nativeColorDialogArm = 0
}

/** True while the OS/browser color dialog is open, about to open, or a click on
 *  the disk just leaked through to the page (Electron/Chromium click-through). */
export function nativeColorPickerActive(): boolean {
  if (nativeColorDialogArm > 0) return true
  if (performance.now() - nativeColorDialogAt < NATIVE_COLOR_GRACE_MS) return true
  const active = document.activeElement
  return active instanceof HTMLInputElement && active.type === 'color'
}

export interface PopoverDismissOptions {
  /**
   * Roots the press may land in without dismissing (the popover panel and the
   * trigger that toggles it). When provided, the outside-press listener runs
   * on capture-phase pointerdown with this containment guard. When omitted,
   * it runs on bubble-phase mousedown and the popover must protect itself
   * with `onMouseDown={(e) => e.stopPropagation()}` (the slides convention).
   */
  inside?: () => ReadonlyArray<Element | null | undefined>
}

/** Install the three dismissal listeners; returns a teardown function. */
export function installPopoverDismiss(
  close: () => void,
  options?: PopoverDismissOptions,
): () => void {
  const inside = options?.inside
  const onPress = (e: Event) => {
    // Clicks on the native color disk land outside the DOM tree of the ribbon
    // panel — without this guard the first palette click closes「其他颜色…」.
    if (nativeColorPickerActive()) return
    if (inside) {
      const target = e.target as Node | null
      if (target) {
        for (const root of inside()) if (root && root.contains(target)) return
      }
    }
    close()
  }
  // Native <input type="color"> opens an OS/browser dialog that blurs the
  // window. Closing on that blur unmounts the input and kills the picker
  // (docs "其他颜色…" / sheets "更多颜色" — sheets already skips blur close
  // for ColorDropdown; this guard covers shared ribbon dismiss).
  const onBlur = () => {
    requestAnimationFrame(() => {
      if (nativeColorPickerActive()) return
      close()
    })
  }
  const onChrome = () => {
    if (nativeColorPickerActive()) return
    close()
  }
  if (inside) window.addEventListener('pointerdown', onPress, true)
  else window.addEventListener('mousedown', onPress)
  window.addEventListener('blur', onBlur)
  const offChrome = subscribeChromePressed(onChrome)
  bumpOpenPopovers(1)
  return () => {
    if (inside) window.removeEventListener('pointerdown', onPress, true)
    else window.removeEventListener('mousedown', onPress)
    window.removeEventListener('blur', onBlur)
    offChrome?.()
    bumpOpenPopovers(-1)
    disarmNativeColorDialog()
  }
}

/** React binding: listeners live only while `open` is true. */
export function useDismissablePopover(
  open: boolean,
  close: () => void,
  options?: PopoverDismissOptions,
): void {
  const closeRef = useRef(close)
  closeRef.current = close
  const insideRef = useRef(options?.inside)
  insideRef.current = options?.inside
  const guarded = options?.inside != null
  useEffect(() => {
    if (!open) return
    return installPopoverDismiss(
      () => closeRef.current(),
      guarded ? { inside: () => insideRef.current?.() ?? [] } : undefined,
    )
  }, [open, guarded])
}
