import { describe, expect, it } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import {
  allPathsSelected,
  filterEntriesByExt,
  listFilterExt,
  nextSelectAll,
  paginateList,
} from '../src/renderer/src/home-list'

const entry = (path: string, ext = 'pdf'): RecentEntry => ({
  path,
  name: path,
  ext,
  mtimeMs: 0,
  sizeBytes: 0,
  starred: false,
})

describe('listFilterExt', () => {
  it('applies type chips in every browse mode, including AI 产物', () => {
    expect(listFilterExt('pdf', true)).toBe('pdf')
    expect(listFilterExt('pdf', false)).toBe('pdf')
    expect(listFilterExt('all', false)).toBeUndefined()
    expect(listFilterExt('all', true)).toBeUndefined()
  })
})

describe('paginateList', () => {
  it('returns the requested page and clamps past the last page', () => {
    const items = Array.from({ length: 23 }, (_, i) => i)
    expect(paginateList(items, 1, 10)).toEqual(items.slice(0, 10))
    expect(paginateList(items, 3, 10)).toEqual(items.slice(20, 23))
    expect(paginateList(items, 9, 10)).toEqual(items.slice(20, 23))
  })
})

describe('nextSelectAll', () => {
  it('selects every filtered path, not just the loaded page', () => {
    const all = Array.from({ length: 147 }, (_, i) => `/f/${i}.pdf`)
    const visible = all.slice(0, 50)
    const selected = nextSelectAll(all, new Set())
    expect(selected.size).toBe(147)
    expect([...selected].slice(0, 50)).toEqual(visible)
    expect(allPathsSelected(visible, selected)).toBe(true)
    expect(allPathsSelected(all, selected)).toBe(true)
  })

  it('clears when the full list is already selected', () => {
    const all = ['/a.pdf', '/b.pdf']
    expect(nextSelectAll(all, new Set(all)).size).toBe(0)
  })

  it('fills in the rest when only the loaded page is checked', () => {
    const all = ['/a.pdf', '/b.pdf', '/c.pdf']
    const next = nextSelectAll(all, new Set(['/a.pdf']))
    expect([...next].sort()).toEqual(['/a.pdf', '/b.pdf', '/c.pdf'])
  })
})

describe('filterEntriesByExt', () => {
  it('keeps the full unpaged list so select-all can use it', () => {
    const all = [
      ...Array.from({ length: 60 }, (_, i) => entry(`/d/${i}.docx`, 'docx')),
      ...Array.from({ length: 87 }, (_, i) => entry(`/p/${i}.pdf`, 'pdf')),
    ]
    expect(filterEntriesByExt(all, undefined)).toHaveLength(147)
    expect(filterEntriesByExt(all, 'pdf')).toHaveLength(87)
  })
})
