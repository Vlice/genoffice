import { escapeXmlAttr } from './xml-utils'

/**
 * Text watermark support. Word implements watermarks as VML shapes with a
 * v:textpath inside the page header; the shapes float behind the body text
 * on every page that uses that header.
 *
 * We tile several diagonal copies across the page (page-relative) so the
 * mark fills the sheet instead of a single centered ghost.
 */

/** namespaces the header part root needs when it carries a VML watermark */
export const WATERMARK_NS =
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:w10="urn:schemas-microsoft-com:office:word"'

/** read the watermark text from a header part; null when it has none */
export function readWatermarkText(headerXml: string): string | null {
  if (!headerXml.includes('<v:textpath')) return null
  const m = /<v:textpath[^>]*\bstring="([^"]*)"/.exec(headerXml)
  if (!m) return null
  const text = m[1]
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
  return text || null
}

const TILE_COLS = 3
const TILE_ROWS = 4
/** Cover both Letter (612×792) and A4 (595×842) with a little overflow. */
const PAGE_W_PT = 612
const PAGE_H_PT = 842
const SHAPE_W_PT = 168
const SHAPE_H_PT = 78

function watermarkShapeXml(text: string, index: number, xPt: number, yPt: number): string {
  const n = index + 1
  return (
    `<v:shape id="PowerPlusWaterMarkObject${n}" o:spid="_x0000_s${2048 + n}" type="#_x0000_t136"` +
    ` style="position:absolute;left:0;text-align:left;margin-left:${xPt.toFixed(2)}pt;margin-top:${yPt.toFixed(2)}pt;` +
    `width:${SHAPE_W_PT}pt;height:${SHAPE_H_PT}pt;rotation:315;z-index:-251656192;` +
    `mso-position-horizontal:absolute;mso-position-horizontal-relative:page;` +
    `mso-position-vertical:absolute;mso-position-vertical-relative:page"` +
    ` o:allowincell="f" fillcolor="silver" stroked="f">` +
    '<v:fill opacity=".5"/>' +
    `<v:textpath style="font-family:&quot;DengXian&quot;;font-size:1pt" string="${escapeXmlAttr(text)}"/>` +
    '</v:shape>'
  )
}

/**
 * Diagonal gray text watermark paragraph. Lives as the first paragraph of the
 * header part; a grid of VML textpath shapes tiles the page.
 */
export function watermarkParagraphXml(text: string): string {
  const shapetype =
    '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"' +
    ' path="m@7,l@8,m@5,21600l@6,21600e">' +
    '<v:formulas>' +
    '<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>' +
    '<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>' +
    '<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>' +
    '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>' +
    '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>' +
    '</v:formulas>' +
    '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"' +
    ' o:connectangles="270,180,90,0"/>' +
    '<v:textpath on="t" fitshape="t"/>' +
    '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
    '<o:lock v:ext="edit" text="t" shapetype="t"/>' +
    '</v:shapetype>'
  const cellW = PAGE_W_PT / TILE_COLS
  const cellH = PAGE_H_PT / TILE_ROWS
  let shapes = ''
  let i = 0
  for (let r = 0; r < TILE_ROWS; r++) {
    for (let c = 0; c < TILE_COLS; c++) {
      const x = cellW * (c + 0.5) - SHAPE_W_PT / 2
      const y = cellH * (r + 0.5) - SHAPE_H_PT / 2
      shapes += watermarkShapeXml(text, i++, x, y)
    }
  }
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:pict>${shapetype}${shapes}</w:pict></w:r></w:p>`
}
