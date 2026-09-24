/**
 * PPT text-box keyboard helpers.
 *
 * Chinese IMEs (Pinyin / Wubi / Sogou / Microsoft Pinyin) often report Enter as
 * key="Process" or "Unidentified" with code="Enter", and may clear shiftKey
 * because Shift is also the 中/英 toggle. Match the physical keys, not the
 * English `e.key` string.
 */

export function isEnterKey(e: { key: string; code: string; keyCode?: number }): boolean {
  return (
    e.key === 'Enter' ||
    e.code === 'Enter' ||
    e.code === 'NumpadEnter' ||
    e.keyCode === 13
  )
}

export function isShiftKey(e: { key: string; code: string }): boolean {
  return e.key === 'Shift' || e.code === 'ShiftLeft' || e.code === 'ShiftRight'
}

export function modifierShiftDown(e: {
  shiftKey: boolean
  getModifierState?: (key: string) => boolean
}): boolean {
  return e.shiftKey || e.getModifierState?.('Shift') === true
}

let physicalShift = false

/** Track Shift even when the IME swallows shiftKey on the following Enter. */
export function notePhysicalShift(e: {
  type: string
  key: string
  code: string
  shiftKey: boolean
  getModifierState?: (key: string) => boolean
}): void {
  if (isShiftKey(e)) {
    physicalShift = e.type !== 'keyup'
    return
  }
  // IME may swallow Shift keyup after using it as a language toggle. Don't leave
  // the latch stuck on ordinary keys; Enter is the combo we still honor via it.
  if (!modifierShiftDown(e) && !isEnterKey(e)) physicalShift = false
}

export function isPhysicalShiftHeld(): boolean {
  return physicalShift
}

export function resetPhysicalShift(): void {
  physicalShift = false
}

export function isSoftLineBreakShortcut(
  e: {
    key: string
    code: string
    keyCode?: number
    shiftKey: boolean
    metaKey?: boolean
    ctrlKey?: boolean
    altKey?: boolean
    isComposing?: boolean
    getModifierState?: (key: string) => boolean
  },
  shiftHeld = isPhysicalShiftHeld(),
): boolean {
  // Candidate-confirm Enter must reach the IME; a following Enter is a real edit key.
  if (e.isComposing) return false
  if (e.metaKey || e.ctrlKey || e.altKey) return false
  if (!isEnterKey(e)) return false
  return modifierShiftDown(e) || shiftHeld
}

/** beforeinput backup: browsers emit insertLineBreak for Shift+Enter; some IMEs
 * mis-label it as insertParagraph while Shift is still physically down. */
export function isSoftLineBreakInput(
  inputType: string,
  shiftHeld = isPhysicalShiftHeld(),
): boolean {
  return inputType === 'insertLineBreak' || (inputType === 'insertParagraph' && shiftHeld)
}

/** IME may deliver both keydown and beforeinput for the same Shift+Enter. */
let softBreakConsumed = false

export function claimSoftBreak(): boolean {
  if (softBreakConsumed) return false
  softBreakConsumed = true
  return true
}

export function releaseSoftBreak(): void {
  softBreakConsumed = false
}

let paraBreakConsumed = false

export function claimParaBreak(): boolean {
  if (paraBreakConsumed) return false
  paraBreakConsumed = true
  return true
}

export function releaseParaBreak(): void {
  paraBreakConsumed = false
}
