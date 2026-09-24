import { describe, expect, it } from 'vitest'
import {
  decodeDataImageUrl,
  isHttpImageUrl,
  isSlideImageRef,
  parseAttachmentIndex,
} from '../src/shared/image-refs'

describe('slide image refs', () => {
  it('accepts http(s), attachment:N, and data:image urls', () => {
    expect(isHttpImageUrl('https://cdn.example/a.png')).toBe(true)
    expect(isSlideImageRef('https://cdn.example/a.png')).toBe(true)
    expect(isSlideImageRef('attachment:0')).toBe(true)
    expect(isSlideImageRef('attachment:12')).toBe(true)
    expect(isSlideImageRef('data:image/png;base64,AAAA')).toBe(true)
    expect(isSlideImageRef('file:///etc/passwd')).toBe(false)
    expect(isSlideImageRef('ftp://nope')).toBe(false)
    expect(isSlideImageRef('attachment:cat')).toBe(false)
  })

  it('parses attachment indexes', () => {
    expect(parseAttachmentIndex('attachment:0')).toBe(0)
    expect(parseAttachmentIndex(' attachment:3 ')).toBe(3)
    expect(parseAttachmentIndex('https://x')).toBeNull()
  })

  it('decodes data:image payloads', () => {
    const r = decodeDataImageUrl('data:image/jpeg;base64,QQ==')
    expect(r).toEqual({ base64: 'QQ==', ext: 'jpg' })
  })
})
