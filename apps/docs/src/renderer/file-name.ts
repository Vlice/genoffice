/**
 * Basename for UI / FILENAME fields. Decodes percent-encoding used by some
 * virtual paths (e.g. historical MoreAI moreai://…/%E6%9C%AA…).
 */
export function fileNameFromPath(path: string | null | undefined, fallback = ''): string {
  const base = path?.split(/[\\/]/).pop() || fallback
  if (!base) return fallback
  if (!/%[0-9A-Fa-f]{2}/.test(base)) return base
  try {
    return decodeURIComponent(base)
  } catch {
    return base
  }
}
