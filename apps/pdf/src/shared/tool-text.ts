/** Models often emit the two-character sequence \\n (or stacked backslashes)
 *  instead of a real newline, and sometimes HTML <br>. Collapse those to '\\n'
 *  so insert_text / rasterize never draw a literal backslash-n. */
export function decodeToolNewlines(text: string): string {
  let s = String(text ?? '')
  s = s.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  s = s.replace(/<br\s*\/?>/gi, '\n')
  // \\n, \\\\n, … — one or more backslashes before n/r
  s = s.replace(/(?:\\)+n/g, '\n').replace(/(?:\\)+r/g, '\n')
  s = s.replace(/\uFF3Cn/g, '\n')
  s = s.replace(/[\u2028\u2029\u0085]/g, '\n')
  return s
}
