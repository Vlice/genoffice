import type { RecentEntry } from '../../shared/home-api'

export function entryMatchesExt(entry: RecentEntry, ext?: string): boolean {
  if (!ext) return true
  const want = ext.toLowerCase()
  if (want === 'xlsx') return entry.ext === 'xlsx' || entry.ext === 'xlsm' || entry.ext === 'xls'
  if (want === 'md') return entry.ext === 'md' || entry.ext === 'markdown'
  if (want === 'table') return entry.ext === 'table' || entry.ext === 'bitable.json' || entry.ext === 'bitable'
  if (want === 'board') {
    return entry.ext === 'board' || entry.ext === 'excalidraw' || entry.ext === 'excalidraw.json'
  }
  if (want === 'code') {
    return (
      entry.ext === 'code' ||
      [
        'py',
        'js',
        'mjs',
        'cjs',
        'ts',
        'tsx',
        'jsx',
        'vue',
        'css',
        'sh',
        'bash',
        'sql',
        'go',
        'rs',
        'java',
        'c',
        'cpp',
        'h',
      ].includes(entry.ext)
    )
  }
  return entry.ext === want
}

export function filterEntriesByExt(all: RecentEntry[], ext?: string): RecentEntry[] {
  return ext ? all.filter((e) => entryMatchesExt(e, ext)) : all
}

/** Type-chip filter for the current view (最近 / 收藏 / 项目 / AI 产物). */
export function listFilterExt(filter: string, _inAiBrowse?: boolean): string | undefined {
  return !filter || filter === 'all' ? undefined : filter
}

export function paginateList<T>(items: readonly T[], page: number, pageSize: number): T[] {
  if (pageSize <= 0) return []
  const pages = Math.max(1, Math.ceil(items.length / pageSize) || 1)
  const safePage = Math.min(pages, Math.max(1, page))
  const start = (safePage - 1) * pageSize
  return items.slice(start, start + pageSize)
}

export function allPathsSelected(
  paths: readonly string[],
  selected: ReadonlySet<string>,
): boolean {
  return paths.length > 0 && paths.every((p) => selected.has(p))
}

/**
 * Header 全选 covers the full filtered list (`allPaths`), not just the loaded
 * page. Toggles off when every path is already selected.
 */
export function nextSelectAll(
  allPaths: readonly string[],
  selected: ReadonlySet<string>,
): Set<string> {
  return allPathsSelected(allPaths, selected) ? new Set() : new Set(allPaths)
}
