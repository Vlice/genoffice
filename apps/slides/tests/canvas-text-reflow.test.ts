import { describe, expect, it } from 'vitest'
import type { GlyphRun, RenderTextLayout, TextLine } from '@genoffice/pptx-render'
import { reflowTextLayoutToCanvas } from '../src/renderer/canvas-text-reflow'

function run(over: Partial<GlyphRun> & { text: string }): GlyphRun {
  return {
    x: 0,
    baselineY: 16,
    fontFamily: 'PingFang SC',
    fontSizePx: 20,
    color: '#000',
    bold: false,
    italic: false,
    underline: false,
    widthPx: 40,
    ...over,
  }
}

function line(runs: GlyphRun[], over: Partial<TextLine> = {}): TextLine {
  return { runs, top: 0, height: 24, paraStart: true, align: 'left', ...over }
}

function layout(lines: TextLine[]): RenderTextLayout {
  return {
    lines,
    insets: { l: 0, t: 0, r: 0, b: 0 },
    anchor: 'top',
    fontScale: 1,
    contentHeight: 24,
    wrap: true,
  }
}

describe('reflowTextLayoutToCanvas', () => {
  it('re-packs wider measured fragments so they no longer share x (overlap)', () => {
    const a = run({ text: 'Multi-Agent', x: 0, widthPx: 40 })
    const b = run({ text: '协同架构', x: 40, widthPx: 40 })
    const text = layout([line([a, b])])
    const changed = reflowTextLayoutToCanvas(text, 400, (r) => (r.text === 'Multi-Agent' ? 120 : 80))
    expect(changed).toBe(true)
    expect(text.lines[0]!.runs[0]!.widthPx).toBe(120)
    expect(text.lines[0]!.runs[1]!.x).toBe(120)
    expect(text.lines[0]!.runs[1]!.widthPx).toBe(80)
  })

  it('wraps to the next line when measured fragments no longer fit the box', () => {
    const a = run({ text: 'Multi-Agent', x: 0, widthPx: 50 })
    const b = run({ text: '协同架构', x: 50, widthPx: 50 })
    const text = layout([line([a, b])])
    reflowTextLayoutToCanvas(text, 100, (r) => (r.text === 'Multi-Agent' ? 90 : 80))
    expect(text.lines).toHaveLength(2)
    expect(text.lines[0]!.runs.map((r) => r.text)).toEqual(['Multi-Agent'])
    expect(text.lines[1]!.runs.map((r) => r.text)).toEqual(['协同架构'])
    expect(text.lines[1]!.top).toBe(24)
    expect(text.lines[1]!.paraStart).toBe(false)
  })

  it('leaves matching metrics untouched (PowerPoint-accurate layouts)', () => {
    const a = run({ text: 'Hello', x: 0, widthPx: 50 })
    const text = layout([line([a])])
    expect(reflowTextLayoutToCanvas(text, 400, () => 50)).toBe(false)
    expect(text.lines[0]!.runs[0]!.x).toBe(0)
  })
})
