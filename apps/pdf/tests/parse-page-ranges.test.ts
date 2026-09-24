import { describe, expect, it } from 'vitest'
import { parsePageRanges } from '../src/renderer/view-config'

describe('parsePageRanges', () => {
  it('parses singles and ascending ranges', () => {
    expect(parsePageRanges('1-3,5', 10)).toEqual([1, 2, 3, 5])
    expect(parsePageRanges('46', 62)).toEqual([46])
  })

  it('normalizes reverse ranges such as 46-1', () => {
    expect(parsePageRanges('46-1', 62)).toEqual(Array.from({ length: 46 }, (_, i) => i + 1))
    expect(parsePageRanges('3-1,5', 10)).toEqual([1, 2, 3, 5])
  })

  it('rejects out-of-bounds and empty input', () => {
    expect(parsePageRanges('1-99', 62)).toBeNull()
    expect(parsePageRanges('0-2', 10)).toBeNull()
    expect(parsePageRanges('abc', 10)).toBeNull()
    expect(parsePageRanges('  ', 10)).toBeNull()
  })
})
