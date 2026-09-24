import { describe, expect, it } from 'vitest'

import { clip } from '../src/renderer/ai/protocol'

describe('clip document-context previews', () => {
  it('does not split a surrogate pair when truncating at an emoji', () => {
    // 🟢 is U+1F7E2 → UTF-16 [0xD83D, 0xDFE2]. slice(0, 1) would leave \uD83D.
    const text = `${'x'.repeat(59)}🟢more`
    const out = clip(text, 60)
    expect(out.includes('\uD83D') && !out.includes('\uDFE2')).toBe(false)
    expect(JSON.stringify(out)).not.toMatch(/\\ud83d(?![0-9a-fA-F]{4})/i)
    expect(out.endsWith('…')).toBe(true)
  })

  it('still truncates ordinary BMP text', () => {
    expect(clip('abcdefghij', 4)).toBe('abcd…')
  })
})
