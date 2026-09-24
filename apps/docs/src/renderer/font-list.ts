import type { Lang } from '@genoffice/i18n'

/**
 * Font dropdown candidates grouped by script, ordered per UI language so the
 * fonts a user is most likely to want appear at the top (e.g. Western fonts
 * for the US market, Japanese fonts for the Japanese market).
 */
const LATIN = [
  'Calibri',
  'Arial',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Cambria',
  'Garamond',
  'Trebuchet MS',
  'Segoe UI',
  'Courier New',
  'Impact',
]
// GB/T 9704 official-document fonts included so government documents can be
// authored from scratch (values are free-typed either way; the combobox accepts any name)
const SIMPLIFIED_CHINESE = [
  '等线',
  '宋体',
  '黑体',
  '微软雅黑',
  '楷体',
  '仿宋',
  '仿宋_GB2312',
  '楷体_GB2312',
  '方正小标宋简体',
]
const JAPANESE = ['Yu Gothic', 'Yu Mincho', 'Meiryo', 'MS Gothic', 'MS Mincho']
const KOREAN = ['Malgun Gothic', 'Batang', 'Gulim', 'Dotum']
const TRADITIONAL_CHINESE = ['Microsoft JhengHei', 'PMingLiU']

export const BUILTIN_FONT_FAMILIES: readonly string[] = [
  ...LATIN,
  ...SIMPLIFIED_CHINESE,
  ...JAPANESE,
  ...KOREAN,
  ...TRADITIONAL_CHINESE,
]

const EAST_ASIAN_FONT_RE =
  /[⺀-鿿豈-﫿぀-ヿㇰ-ㇿ가-힯]|sim(sun|hei)|nsimsun|kaiti|fangsong|dengxian|yahei|songti|heiti|xingkai|lisu|youyuan|st(zhongsong|song|kai|fangsong|xihei|hupo|liti|caiyun)|pingfang|hiragino|meiryo|osaka|kozuka|yu (gothic|mincho)|yugoth|ms (ui )?p?(gothic|mincho)|biz ud|malgun|batang|gulim|dotum|gungsuh|m(ye|yu)ngjo|nanum|apple (sd )?gothic|applemyungjo|jhenghei|p?mingliu|biaukai|dfkai|kaiu|source han|noto (sans|serif) (cjk|sc|tc|hk|jp|kr)|wenquanyi/i

/**
 * Which rFonts slot a font-box pick should target: East Asian names go to
 * w:eastAsia, everything else to w:ascii/w:hAnsi.
 */
export function isEastAsianFontName(name: string): boolean {
  return EAST_ASIAN_FONT_RE.test(name.normalize('NFKC'))
}

/** Patch for a user/AI font-box pick. CJK faces cover ASCII glyphs, so they
 *  write both rFonts slots — otherwise digits/letters keep the previous Latin
 *  face (Impact stays Impact while 仿宋 only paints CJK). Latin faces usually
 *  lack CJK glyphs, so they still write only ascii/hAnsi. */
export function fontPickPatch(name: string | null): { font?: string | null; fontAscii?: string | null } {
  if (!name) return { font: null, fontAscii: null }
  if (isEastAsianFontName(name)) return { font: name, fontAscii: name }
  return { fontAscii: name }
}

export function fontFamiliesFor(lang: Lang): readonly string[] {
  switch (lang) {
    case 'zh':
      return [...SIMPLIFIED_CHINESE, ...LATIN, ...JAPANESE, ...KOREAN, ...TRADITIONAL_CHINESE]
    case 'zh-TW':
      return [...TRADITIONAL_CHINESE, ...LATIN, ...SIMPLIFIED_CHINESE, ...JAPANESE, ...KOREAN]
    case 'ja':
      return [...JAPANESE, ...LATIN, ...SIMPLIFIED_CHINESE, ...KOREAN, ...TRADITIONAL_CHINESE]
    case 'ko':
      return [...KOREAN, ...LATIN, ...JAPANESE, ...SIMPLIFIED_CHINESE, ...TRADITIONAL_CHINESE]
    default:
      return [...LATIN, ...JAPANESE, ...SIMPLIFIED_CHINESE, ...KOREAN, ...TRADITIONAL_CHINESE]
  }
}

/**
 * docDefaults w:eastAsia font for new blank documents, matching what Word
 * ships per market (zh → SimSun, ja → Yu Mincho, ko → Malgun Gothic,
 * zh-TW → PMingLiU). English and every other language return undefined —
 * like en-US Word, whose theme leaves the East Asian slot empty and lets
 * per-script substitution kick in only when CJK text actually appears.
 * Latin default stays Calibri for every language.
 */
export function defaultEastAsiaFontFor(lang: Lang): string | undefined {
  switch (lang) {
    case 'zh':
      return '宋体'
    case 'ja':
      return 'Yu Mincho'
    case 'ko':
      return 'Malgun Gothic'
    case 'zh-TW':
      return 'PMingLiU'
    default:
      return undefined
  }
}
