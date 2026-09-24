import { describe, expect, it } from 'vitest'
import { decodeToolNewlines } from '../src/shared/tool-text'

describe('decodeToolNewlines', () => {
  it('keeps real newlines', () => {
    expect(decodeToolNewlines('a\nb')).toBe('a\nb')
  })

  it('turns the two-character sequence \\n into a newline', () => {
    expect(decodeToolNewlines('a\\nb')).toBe('a\nb')
  })

  it('collapses stacked backslashes before n', () => {
    expect(decodeToolNewlines('a\\\\nb')).toBe('a\nb')
    expect(decodeToolNewlines('a\\\\\\\\nb')).toBe('a\nb')
  })

  it('accepts HTML line breaks', () => {
    expect(decodeToolNewlines('a<br>b<br/>c')).toBe('a\nb\nc')
  })
})
