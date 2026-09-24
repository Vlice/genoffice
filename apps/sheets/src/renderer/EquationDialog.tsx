import { useMemo, useRef, useState } from 'react'

import { latexToOmml, ommlToMathML } from '@genoffice/docx-engine/math'
import { useI18n } from './i18n/locale'

/// Excel's Insert → Equation: LaTeX input with a live MathML preview (docs'
/// LaTeX→OMML pipeline, rendered natively by Chromium). Inserting rasterizes
/// the painted preview to a PNG so it rides the same picture pipeline
/// (journal + xl/media) as any inserted image.
///
/// Do NOT wrap MathML in SVG `<foreignObject>` and load it as `<img>`:
/// Chrome's SVG-as-image renderer skips MathML, so OK looked like a no-op
/// (blank/transparent picture, or onerror with no toast).

const PRESETS = [
  'A = \\pi r^2',
  'a^2 + b^2 = c^2',
  'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
  '(x+a)^n = \\sum_{k=0}^{n} \\binom{n}{k} x^k a^{n-k}',
  'e^x = 1 + \\frac{x}{1!} + \\frac{x^2}{2!} + \\frac{x^3}{3!} + \\cdots',
  'f(x) = a_0 + \\sum_{n=1}^{\\infty} \\left( a_n \\cos \\frac{n\\pi x}{L} + b_n \\sin \\frac{n\\pi x}{L} \\right)',
]
const MATH_FONT = "'STIX Two Math', 'Cambria Math', 'Latin Modern Math', serif"
const FONT_PX = 30
/// Rasterize at 2× and anchor at the measured CSS size, so the PNG stays
/// sharp on retina displays.
const RASTER_SCALE = 2
/// Breathing room around the measured MathML box; glyphs (radicals, large
/// operators) can paint slightly outside their layout bounds.
const PADDING_PX = 4

function mathmlOf(latex: string): { mathml: string } | { error: string } | null {
  if (!latex.trim()) return null
  try {
    const omml = latexToOmml(latex)
    const mathml = ommlToMathML(`<m:oMath>${omml}</m:oMath>`)
    return mathml ? { mathml } : { error: latex }
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function paintSvg(svg: string, width: number, height: number): Promise<string> {
  const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = width * RASTER_SCALE
      canvas.height = height * RASTER_SCALE
      const context = canvas.getContext('2d')
      if (!context) {
        URL.revokeObjectURL(url)
        reject(new Error('canvas 2d context unavailable'))
        return
      }
      context.drawImage(image, 0, 0, canvas.width, canvas.height)
      URL.revokeObjectURL(url)
      resolve(canvas.toDataURL('image/png'))
    }
    image.onerror = () => {
      URL.revokeObjectURL(url)
      reject(new Error('equation rasterization failed'))
    }
    image.src = url
  })
}

/// Snapshot the live HTML MathML preview (which Chromium *does* paint) into
/// an SVG of positioned `<text>`/`<line>` — no foreignObject, so `<img>` and
/// canvas accept it in MoreAI's iframe embed.
export function rasterizePreviewNode(node: HTMLElement): Promise<{
  dataUrl: string
  width: number
  height: number
}> {
  const rect = node.getBoundingClientRect()
  const width = Math.max(1, Math.ceil(rect.width) + PADDING_PX * 2)
  const height = Math.max(1, Math.ceil(rect.height) + PADDING_PX * 2)
  const parts: string[] = [
    `<rect width="${width}" height="${height}" fill="#ffffff"/>`,
  ]

  const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT)
  let current: Node | null
  while ((current = walker.nextNode())) {
    const text = current as Text
    const value = text.data
    if (!value) continue
    const parent = text.parentElement
    if (!parent) continue
    const range = document.createRange()
    range.selectNodeContents(text)
    const box = range.getBoundingClientRect()
    if (box.width === 0 && box.height === 0) continue
    const cs = getComputedStyle(parent)
    const x = box.left - rect.left + PADDING_PX
    const y = box.top - rect.top + PADDING_PX + box.height * 0.82
    const family = cs.fontFamily || MATH_FONT
    parts.push(
      `<text x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${cs.fontSize || `${FONT_PX}px`}" ` +
        `font-family="${escapeXml(family)}" font-style="${cs.fontStyle}" ` +
        `font-weight="${cs.fontWeight}" fill="${cs.color || '#000'}">${escapeXml(value)}</text>`,
    )
  }

  for (const el of node.querySelectorAll('mfrac')) {
    const num = el.firstElementChild
    if (!num) continue
    const box = el.getBoundingClientRect()
    const nb = num.getBoundingClientRect()
    const y = nb.bottom - rect.top + PADDING_PX + 1
    const x1 = box.left - rect.left + PADDING_PX
    const x2 = box.right - rect.left + PADDING_PX
    const color = getComputedStyle(el).color || '#000'
    parts.push(
      `<line x1="${x1.toFixed(2)}" y1="${y.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y.toFixed(2)}" ` +
        `stroke="${color}" stroke-width="1.2"/>`,
    )
  }

  if (parts.length <= 1) {
    return Promise.reject(new Error('equation preview has no painted glyphs'))
  }

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width * RASTER_SCALE}" ` +
    `height="${height * RASTER_SCALE}" viewBox="0 0 ${width} ${height}">` +
    parts.join('') +
    `</svg>`
  return paintSvg(svg, width, height).then((dataUrl) => ({ dataUrl, width, height }))
}

export function EquationDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (dataUrl: string, width: number, height: number) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [latex, setLatex] = useState('')
  const [busy, setBusy] = useState(false)
  const [insertError, setInsertError] = useState<string | null>(null)
  const previewRef = useRef<HTMLSpanElement>(null)
  const preview = useMemo(() => mathmlOf(latex), [latex])
  const canInsert = preview !== null && 'mathml' in preview

  const insert = (): void => {
    const node = previewRef.current
    if (busy || !canInsert || !node) return
    setBusy(true)
    setInsertError(null)
    // Wait one frame so the preview's layout matches what the user sees.
    requestAnimationFrame(() => {
      rasterizePreviewNode(node)
        .then(({ dataUrl, width, height }) => {
          onInsert(dataUrl, width, height)
          onClose()
        })
        .catch((error: unknown) => {
          setBusy(false)
          setInsertError(error instanceof Error ? error.message : String(error))
        })
    })
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog equation-dialog"
        role="dialog"
        aria-label={t('dlgEquationTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgEquationTitle')}</header>
        <section className="dialog-body">
          <div className="equation-presets">
            {PRESETS.map((preset) => {
              const rendered = mathmlOf(preset)
              return (
                <button
                  key={preset}
                  type="button"
                  data-tip={preset}
                  aria-label={preset}
                  onClick={() => {
                    setLatex(preset)
                    setInsertError(null)
                  }}
                >
                  {rendered !== null && 'mathml' in rendered && (
                    <span dangerouslySetInnerHTML={{ __html: rendered.mathml }} />
                  )}
                </button>
              )
            })}
          </div>
          <input
            className="equation-latex"
            value={latex}
            placeholder={t('dlgEquationPlaceholder')}
            autoFocus
            onChange={(event) => {
              setLatex(event.target.value)
              setInsertError(null)
            }}
            onKeyDown={(event) => event.key === 'Enter' && insert()}
          />
          <div className="equation-preview">
            {preview === null && <span className="dialog-note">{t('dlgEquationHint')}</span>}
            {preview !== null && 'error' in preview && (
              <span className="equation-error">{preview.error}</span>
            )}
            {preview !== null && 'mathml' in preview && (
              <span
                ref={previewRef}
                style={{
                  fontFamily: MATH_FONT,
                  fontSize: FONT_PX,
                  display: 'inline-block',
                }}
                dangerouslySetInnerHTML={{ __html: preview.mathml }}
              />
            )}
          </div>
          {insertError && <span className="equation-error">{insertError}</span>}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary" disabled={!canInsert || busy} onClick={insert}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
