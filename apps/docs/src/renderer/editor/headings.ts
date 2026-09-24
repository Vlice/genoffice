import type { Node as PmNode } from '@tiptap/pm/model'
import type { StyleInfo } from '@genoffice/docx-engine'

export interface HeadingRef {
  text: string
  level: number
  /** top-level position of the docHeading node */
  pos: number
}

/**
 * w:pStyle to stamp when the user (or AI) changes outline role.
 * TipTap's setNode copies the current block attrs, so a heading saved as
 * Heading1 keeps that styleId unless we overwrite it — [data-style] CSS then
 * keeps showing Heading 1 after a switch to Heading 2 / Normal.
 * level 0 = body paragraph (the document default, usually Normal).
 */
export function styleIdForOutlineLevel(
  styles: Map<string, StyleInfo> | undefined,
  level: number,
): string {
  if (level <= 0) {
    if (styles) {
      for (const info of styles.values()) {
        if (info.isDefault && info.type === 'paragraph') return info.styleId
      }
      if (styles.has('Normal')) return 'Normal'
    }
    return 'Normal'
  }
  if (styles) {
    for (const info of styles.values()) {
      if (info.type === 'paragraph' && info.headingLevel === level) return info.styleId
    }
    const id = `Heading${level}`
    if (styles.has(id)) return id
  }
  return `Heading${level}`
}

/** Single heading predicate shared by TOC, nav pane and TOC page backfill (document order) */
export function collectHeadings(doc: PmNode): HeadingRef[] {
  const out: HeadingRef[] = []
  doc.forEach((node, offset) => {
    if (node.type.name === 'docHeading' && node.textContent.trim()) {
      out.push({ text: node.textContent, level: Number(node.attrs.level) || 1, pos: offset })
    }
  })
  return out
}
