/**
 * Superscript while editing: vertical-align does not move glyphs inside a focused
 * contentEditable (the raise used to appear only after blur, when the canvas
 * redrew). The ribbon shift is a position offset that paints immediately.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { extractParagraphs, toggleSelectionBaseline } from '../src/renderer/TextEditOverlay'

afterEach(() => {
  document.body.innerHTML = ''
  window.getSelection()?.removeAllRanges()
})

function editorWithFragment(text: string): { root: HTMLDivElement; frag: HTMLSpanElement } {
  const root = document.createElement('div')
  root.contentEditable = 'true'
  root.tabIndex = 0
  Object.defineProperty(root, 'isContentEditable', { value: true })
  const frag = document.createElement('span')
  frag.dataset.layoutFragment = 'true'
  frag.style.display = 'inline-block'
  frag.style.verticalAlign = 'top'
  frag.style.width = '80px'
  frag.textContent = text
  root.appendChild(frag)
  document.body.appendChild(root)
  root.focus()
  return { root, frag }
}

function selectText(node: Text, start: number, end: number) {
  const range = document.createRange()
  range.setStart(node, start)
  range.setEnd(node, end)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
}

describe('superscript while the editor is focused', () => {
  it('raises the selection immediately and commits baseline 30', () => {
    const { root, frag } = editorWithFragment('abc')
    selectText(frag.firstChild as Text, 1, 2)
    toggleSelectionBaseline('super')

    const marked = root.querySelector<HTMLElement>('span[data-baseline]')
    expect(marked).toBeTruthy()
    expect(marked!.dataset.baseline).toBe('30')
    expect(marked!.style.top).toBe('-0.3em')
    expect(marked!.textContent).toBe('b')
    // The rest of the line stays pinned to the canvas advances.
    expect(frag.style.verticalAlign).toBe('top')
    expect(frag.style.width).toBe('80px')

    const runs = extractParagraphs(root, 1)[0]!.runs
    expect(runs.map((r) => r.text).join('')).toBe('abc')
    expect(runs.find((r) => r.text === 'b')!.baseline).toBe(30)
    expect(runs.find((r) => r.text === 'a')!.baseline).toBe(0)
  })

  it('toggles superscript off on the next click', () => {
    const { root, frag } = editorWithFragment('abc')
    selectText(frag.firstChild as Text, 0, 3)
    toggleSelectionBaseline('super')
    toggleSelectionBaseline('super')

    const marked = root.querySelector<HTMLElement>('span[data-baseline]')
    expect(marked!.dataset.baseline).toBe('0')
    expect(marked!.style.top).toBe('')
    expect(extractParagraphs(root, 1)[0]!.runs.map((r) => r.text).join('')).toBe('abc')
    expect(extractParagraphs(root, 1)[0]!.runs.every((r) => !r.baseline)).toBe(true)
  })

  it('lowers a subscript by 25% of the font size', () => {
    const { root, frag } = editorWithFragment('x2')
    selectText(frag.firstChild as Text, 1, 2)
    toggleSelectionBaseline('sub')
    const marked = root.querySelector<HTMLElement>('span[data-baseline]')
    expect(marked!.dataset.baseline).toBe('-25')
    expect(marked!.style.top).toBe('0.25em')
    expect(extractParagraphs(root, 1)[0]!.runs.find((r) => r.text === '2')!.baseline).toBe(-25)
  })
})
