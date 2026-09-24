/**
 * CSS font stacks used by both Konva paint and the edit overlay / canvas reflow.
 * Kept out of konva-adapter so canvas-text-reflow can measure with the same
 * family list without a circular import.
 */
import { classifyCjkScript } from '../shared/cjk-script'

const JA_SANS = "'Yu Gothic', 'Hiragino Sans', Meiryo, 'Noto Sans JP', sans-serif"
const JA_SERIF = "'Yu Mincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP', serif"
const KO_SANS = "'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif"
const KO_SERIF = "Batang, AppleMyungjo, 'Noto Serif KR', serif"
const TC_SANS = "'Microsoft JhengHei', 'PingFang TC', 'Heiti TC', 'Noto Sans TC', sans-serif"
const TC_SERIF = "PMingLiU, 'Songti TC', 'Noto Serif TC', serif"
const SERIF_HINT_RE =
  /mincho|明朝|batang|바탕|myeongjo|명조|gungsuh|궁서|mingliu|細明|標楷|宋|song/i

/**
 * Display font stack: font names in the file may not be installed locally (Microsoft YaHei
 * on mac / PingFang on win), so append cross-platform equivalents as CSS-level fallbacks.
 * Metrics are handled by the main process FontMetricsProvider's alias table.
 */
const FONT_STACK: Record<string, string> = {
  'microsoft yahei': "'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', sans-serif",
  微软雅黑: "'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', sans-serif",
  'pingfang sc': "'PingFang SC', 'Microsoft YaHei', 'Noto Sans SC', sans-serif",
  苹方: "'PingFang SC', 'Microsoft YaHei', 'Noto Sans SC', sans-serif",
  宋体: "SimSun, 'Songti SC', serif",
  simsun: "SimSun, 'Songti SC', serif",
  黑体: "SimHei, 'Heiti SC', sans-serif",
  simhei: "SimHei, 'Heiti SC', sans-serif",
  楷体: "KaiTi, 'Kaiti SC', serif",
  kaiti: "KaiTi, 'Kaiti SC', serif",
  仿宋: "FangSong, 'Songti SC', serif",
  等线: "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif",
  dengxian: "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif",
  // Western: consistent with the main-process metrics alias chain (calibri→Carlito→Arial etc.),
  // otherwise metrics use Arial while drawing falls back to the system default font, misaligning word spacing/line breaks.
  calibri: 'Calibri, Carlito, Arial, sans-serif',
  'calibri light': "'Calibri Light', Carlito, Arial, sans-serif",
  helvetica: 'Helvetica, Arial, sans-serif',
  'helvetica neue': "'Helvetica Neue', Helvetica, Arial, sans-serif",
  cambria: 'Cambria, Georgia, serif',
  // Japanese (win family names <-> mac Hiragino back each other up; Japanese fonts first, then Chinese fallback, so kanji don't render with Chinese glyph shapes)
  'yu gothic': JA_SANS,
  游ゴシック: "'游ゴシック', " + JA_SANS,
  meiryo: 'Meiryo, ' + JA_SANS,
  メイリオ: "'メイリオ', Meiryo, " + JA_SANS,
  'ms gothic': "'MS Gothic', 'MS PGothic', " + JA_SANS,
  'ms pgothic': "'MS PGothic', 'MS Gothic', " + JA_SANS,
  'ms ui gothic': "'MS UI Gothic', 'MS PGothic', " + JA_SANS,
  'ms ゴシック': "'ＭＳ ゴシック', 'MS Gothic', " + JA_SANS,
  'ms pゴシック': "'ＭＳ Ｐゴシック', 'MS PGothic', " + JA_SANS,
  'hiragino sans': "'Hiragino Sans', " + JA_SANS,
  'hiragino kaku gothic pron': "'Hiragino Kaku Gothic ProN', " + JA_SANS,
  ヒラギノ角ゴシック: "'Hiragino Sans', " + JA_SANS,
  'noto sans jp': "'Noto Sans JP', " + JA_SANS,
  'yu mincho': JA_SERIF,
  游明朝: "'游明朝', " + JA_SERIF,
  'ms mincho': "'MS Mincho', 'MS PMincho', " + JA_SERIF,
  'ms pmincho': "'MS PMincho', 'MS Mincho', " + JA_SERIF,
  'ms 明朝': "'ＭＳ 明朝', 'MS Mincho', " + JA_SERIF,
  'ms p明朝': "'ＭＳ Ｐ明朝', 'MS PMincho', " + JA_SERIF,
  'hiragino mincho pron': "'Hiragino Mincho ProN', " + JA_SERIF,
  ヒラギノ明朝: "'Hiragino Mincho ProN', " + JA_SERIF,
  'noto serif jp': "'Noto Serif JP', " + JA_SERIF,
  // Korean
  'malgun gothic': KO_SANS,
  '맑은 고딕': "'맑은 고딕', " + KO_SANS,
  'apple sd gothic neo': "'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif",
  gulim: 'Gulim, Dotum, ' + KO_SANS,
  굴림: "'굴림', Gulim, Dotum, " + KO_SANS,
  dotum: 'Dotum, Gulim, ' + KO_SANS,
  돋움: "'돋움', Dotum, Gulim, " + KO_SANS,
  'noto sans kr': "'Noto Sans KR', " + KO_SANS,
  batang: KO_SERIF,
  바탕: "'바탕', " + KO_SERIF,
  gungsuh: 'Gungsuh, ' + KO_SERIF,
  궁서: "'궁서', Gungsuh, " + KO_SERIF,
  // Traditional Chinese
  'microsoft jhenghei': TC_SANS,
  微軟正黑體: "'微軟正黑體', " + TC_SANS,
  'pingfang tc': "'PingFang TC', 'Microsoft JhengHei', 'Heiti TC', 'Noto Sans TC', sans-serif",
  'pingfang hk': "'PingFang HK', 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans TC', sans-serif",
  pmingliu: TC_SERIF,
  新細明體: "'新細明體', " + TC_SERIF,
  mingliu: "MingLiU, 'PMingLiU', 'Songti TC', serif",
  細明體: "'細明體', MingLiU, 'Songti TC', serif",
  'dfkai-sb': "'DFKai-SB', BiauKai, 'Kaiti TC', serif",
  標楷體: "'標楷體', 'DFKai-SB', BiauKai, 'Kaiti TC', serif",
}

export function displayFontFamily(name: string): string {
  const stack = FONT_STACK[name.normalize('NFKC').toLowerCase()]
  if (stack) return stack
  // Unknown fonts first get script detection by family name and same-script fallback, so Japanese/Korean/Traditional glyphs don't render as Simplified Chinese shapes
  const script = classifyCjkScript(name)
  if (script === 'ja') return `'${name}', ${SERIF_HINT_RE.test(name) ? JA_SERIF : JA_SANS}`
  if (script === 'ko') return `'${name}', ${SERIF_HINT_RE.test(name) ? KO_SERIF : KO_SANS}`
  if (script === 'tc') return `'${name}', ${SERIF_HINT_RE.test(name) ? TC_SERIF : TC_SANS}`
  return `'${name}', 'PingFang SC', 'Microsoft YaHei', sans-serif`
}
