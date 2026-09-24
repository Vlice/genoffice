import { describe, expect, it } from 'vitest'

import {
  isVisualDragActive,
  restoreHostVisual,
  unstickVisualDragIfStale,
} from '../src/renderer/WorkbookVisuals'

describe('restoreHostVisual', () => {
  it('clears hidden host styles so the chart is hittable again', () => {
    const host = { style: { visibility: 'hidden', opacity: '0' } } as HTMLElement
    restoreHostVisual(host)
    expect(host.style.visibility).toBe('')
    expect(host.style.opacity).toBe('')
  })

  it('ignores a missing host', () => {
    expect(() => restoreHostVisual(null)).not.toThrow()
    expect(() => restoreHostVisual(undefined)).not.toThrow()
  })
})

describe('unstickVisualDragIfStale', () => {
  it('is a no-op when no drag is active', () => {
    expect(isVisualDragActive()).toBe(false)
    expect(unstickVisualDragIfStale(0)).toBe(false)
    expect(isVisualDragActive()).toBe(false)
  })
})
