/**
 * Detect plain-text Markdown on paste so "# Title" / "**bold**" become real
 * structure instead of literal characters (same signals as the docs app).
 */

const STRONG_SIGNALS: readonly RegExp[] = [
  /^#{1,6}\s+\S/m,
  /^\s{0,3}```/m,
  /\[[^\]\n]+\]\([^\s)]+\)/,
  /(?:^|\W)\*\*[^*\n]+\*\*(?:\W|$)/,
  /^\s{0,3}\|.+\|\s*$\n^\s{0,3}\|[\s:|-]+\|\s*$/m,
]

const LIST_LINE = /^\s{0,3}(?:[-*+]|\d{1,3}[.)])\s+\S/
const QUOTE_LINE = /^\s{0,3}>\s?\S/

export function looksLikeMarkdown(text: string): boolean {
  if (STRONG_SIGNALS.some((signal) => signal.test(text))) return true
  const lines = text.split('\n')
  const listLines = lines.filter((line) => LIST_LINE.test(line)).length
  const quoteLines = lines.filter((line) => QUOTE_LINE.test(line)).length
  return listLines >= 2 || quoteLines >= 2
}
