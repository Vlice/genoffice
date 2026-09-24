import type { ChainedCommands, Editor } from '@tiptap/core'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'

const TEXT_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

export type StyleAwareMark = 'bold' | 'italic'

const OFF_ATTR: Record<StyleAwareMark, 'boldOff' | 'italicOff'> = {
  bold: 'boldOff',
  italic: 'italicOff',
}

function styleSize(styleId: unknown, styles: Map<string, StyleInfo> | undefined): number | null {
  if (typeof styleId !== 'string' || !styleId) return null
  const size = styles?.get(styleId)?.display?.sizeHalfPoints
  return typeof size === 'number' && size > 0 ? size : null
}

function styleFlag(
  styleId: unknown,
  flag: StyleAwareMark,
  styles: Map<string, StyleInfo> | undefined,
): boolean | undefined {
  if (typeof styleId !== 'string' || !styleId) return undefined
  const v = styles?.get(styleId)?.display?.[flag]
  return typeof v === 'boolean' ? v : undefined
}

function resolveStyleMaps(
  editor: Editor,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): { styles?: Map<string, StyleInfo>; docDefaults?: DocDefaults } {
  const store = editor.storage.listNumbering as
    | { styles?: Map<string, StyleInfo>; docDefaults?: DocDefaults }
    | undefined
  return {
    styles: styles ?? store?.styles,
    docDefaults: docDefaults ?? store?.docDefaults,
  }
}

/** Resolve the font size Word displays at the caret without adding direct formatting. */
export function effectiveSizeHalfPoints(
  editor: Editor,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): number | null {
  const textAttrs = editor.getAttributes('docTextStyle')
  const directSize = Number(textAttrs.sizeHalfPoints)
  if (Number.isFinite(directSize) && directSize > 0) return directSize

  const characterStyleSize = styleSize(textAttrs.styleId, styles)
  if (characterStyleSize !== null) return characterStyleSize

  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth)
    if (!TEXT_BLOCK_TYPES.has(node.type.name)) continue
    const paragraphStyleSize = styleSize(node.attrs.styleId, styles)
    if (paragraphStyleSize !== null) return paragraphStyleSize
    break
  }

  const defaultSize = docDefaults?.sizeHalfPoints
  return typeof defaultSize === 'number' && defaultSize > 0 ? defaultSize : null
}

/**
 * Whether paragraph/character style (or docDefaults) contributes bold/italic —
 * the CSS that `[data-style=…]` / `.doc-page` already applies, before any run mark.
 */
export function inheritedTextFlag(
  editor: Editor,
  flag: StyleAwareMark,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): boolean {
  const maps = resolveStyleMaps(editor, styles, docDefaults)
  const textAttrs = editor.getAttributes('docTextStyle')
  const char = styleFlag(textAttrs.styleId, flag, maps.styles)
  if (char === true) return true
  if (char === false) return false

  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth)
    if (!TEXT_BLOCK_TYPES.has(node.type.name)) continue
    const para = styleFlag(node.attrs.styleId, flag, maps.styles)
    if (para === true) return true
    if (para === false) return false
    break
  }

  return !!(flag === 'bold' ? maps.docDefaults?.bold : maps.docDefaults?.italic)
}

/** Effective bold/italic as Word shows it (mark, or style-inherited unless explicitly off). */
export function isEffectivelyMarked(
  editor: Editor,
  flag: StyleAwareMark,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): boolean {
  if (editor.isActive(flag)) return true
  if (editor.getAttributes('docTextStyle')[OFF_ATTR[flag]]) return false
  return inheritedTextFlag(editor, flag, styles, docDefaults)
}

/** Append Word-style bold/italic set onto an existing command chain (one undo step). */
export function applyStyleAwareMarkToChain(
  chain: ChainedCommands,
  editor: Editor,
  flag: StyleAwareMark,
  wantOn: boolean,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): ChainedCommands {
  const inherited = inheritedTextFlag(editor, flag, styles, docDefaults)
  const offKey = OFF_ATTR[flag]
  if (wantOn) {
    if (inherited) {
      return chain.setMark('docTextStyle', { [offKey]: null }).unsetMark(flag)
    }
    // drop a leftover explicit-off without inventing an empty docTextStyle mark
    if (editor.getAttributes('docTextStyle')[offKey]) {
      chain = chain.setMark('docTextStyle', { [offKey]: null })
    }
    return chain.setMark(flag)
  }
  chain = chain.unsetMark(flag)
  return inherited ? chain.setMark('docTextStyle', { [offKey]: true }) : chain
}

/**
 * Set bold/italic the way Word does: when the look comes from a style, turning
 * off writes `boldOff`/`italicOff` (w:b/w:i w:val="0") instead of only dropping
 * the TipTap mark — otherwise style CSS keeps the text looking bold/italic.
 */
export function setStyleAwareMark(
  editor: Editor,
  flag: StyleAwareMark,
  wantOn: boolean,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): boolean {
  return applyStyleAwareMarkToChain(
    editor.chain().focus(),
    editor,
    flag,
    wantOn,
    styles,
    docDefaults,
  ).run()
}

/** Toggle bold/italic with style-inherited override (Word w:val="0" semantics). */
export function toggleStyleAwareMark(
  editor: Editor,
  flag: StyleAwareMark,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): boolean {
  const on = isEffectivelyMarked(editor, flag, styles, docDefaults)
  return setStyleAwareMark(editor, flag, !on, styles, docDefaults)
}
