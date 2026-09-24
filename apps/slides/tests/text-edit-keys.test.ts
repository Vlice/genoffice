import { afterEach, describe, expect, it } from 'vitest'
import {
  claimSoftBreak,
  isEnterKey,
  isSoftLineBreakInput,
  isSoftLineBreakShortcut,
  notePhysicalShift,
  releaseSoftBreak,
  resetPhysicalShift,
} from '../src/renderer/text-edit-keys'

afterEach(() => {
  resetPhysicalShift()
  releaseSoftBreak()
})

function key(
  over: Partial<{
    key: string
    code: string
    keyCode: number
    shiftKey: boolean
    metaKey: boolean
    ctrlKey: boolean
    altKey: boolean
    isComposing: boolean
  }> = {},
) {
  return {
    key: 'Enter',
    code: 'Enter',
    keyCode: 13,
    shiftKey: false,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    ...over,
  }
}

describe('isEnterKey', () => {
  it('accepts English Enter, numpad Enter, and IME Process/Unidentified on the Enter key', () => {
    expect(isEnterKey(key())).toBe(true)
    expect(isEnterKey(key({ code: 'NumpadEnter' }))).toBe(true)
    expect(isEnterKey(key({ key: 'Process', keyCode: 13 }))).toBe(true)
    expect(isEnterKey(key({ key: 'Unidentified', code: 'Enter' }))).toBe(true)
    expect(isEnterKey({ key: 'a', code: 'KeyA', keyCode: 65 })).toBe(false)
  })
})

describe('isSoftLineBreakShortcut', () => {
  it('matches English Shift+Enter', () => {
    expect(isSoftLineBreakShortcut(key({ shiftKey: true }))).toBe(true)
  })

  it('matches Chinese IME Process/Unidentified Shift+Enter via code, not e.key', () => {
    expect(isSoftLineBreakShortcut(key({ key: 'Process', shiftKey: true, keyCode: 229 }))).toBe(
      true,
    )
    expect(
      isSoftLineBreakShortcut(key({ key: 'Unidentified', shiftKey: true, keyCode: 229 })),
    ).toBe(true)
  })

  it('matches Enter while Shift is physically held even if the IME cleared shiftKey', () => {
    notePhysicalShift({ type: 'keydown', key: 'Shift', code: 'ShiftLeft', shiftKey: true })
    expect(isSoftLineBreakShortcut(key({ key: 'Process', shiftKey: false, keyCode: 229 }))).toBe(
      true,
    )
  })

  it('does not steal IME candidate-confirm Enter', () => {
    expect(isSoftLineBreakShortcut(key({ shiftKey: true, isComposing: true }))).toBe(false)
  })

  it('does not treat plain Enter or ⌘/Ctrl+Enter as a soft break', () => {
    expect(isSoftLineBreakShortcut(key())).toBe(false)
    expect(isSoftLineBreakShortcut(key({ shiftKey: true, metaKey: true }))).toBe(false)
    expect(isSoftLineBreakShortcut(key({ shiftKey: true, ctrlKey: true }))).toBe(false)
  })

  it('releases the physical-Shift latch on a non-Enter key so a later Enter is a paragraph', () => {
    notePhysicalShift({ type: 'keydown', key: 'Shift', code: 'ShiftLeft', shiftKey: true })
    notePhysicalShift({ type: 'keyup', key: 'Shift', code: 'ShiftLeft', shiftKey: false })
    expect(isSoftLineBreakShortcut(key())).toBe(false)
  })
})

describe('isSoftLineBreakInput', () => {
  it('treats insertLineBreak as a soft break', () => {
    expect(isSoftLineBreakInput('insertLineBreak')).toBe(true)
    expect(isSoftLineBreakInput('insertParagraph')).toBe(false)
  })

  it('treats insertParagraph as a soft break only while Shift is held (IME mis-label)', () => {
    notePhysicalShift({ type: 'keydown', key: 'Shift', code: 'ShiftLeft', shiftKey: true })
    expect(isSoftLineBreakInput('insertParagraph')).toBe(true)
  })
})

describe('claimSoftBreak', () => {
  it('is single-shot until released so keydown + beforeinput do not insert twice', () => {
    expect(claimSoftBreak()).toBe(true)
    expect(claimSoftBreak()).toBe(false)
    releaseSoftBreak()
    expect(claimSoftBreak()).toBe(true)
  })
})
