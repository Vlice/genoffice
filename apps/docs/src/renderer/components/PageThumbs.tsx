import { useCallback, useEffect, useRef, useState } from 'react'

const THUMB_W = 120

function topLevelPageGaps(pm: HTMLElement): HTMLElement[] {
  return [...pm.querySelectorAll<HTMLElement>('.page-gap')].filter(
    (el) =>
      !el.classList.contains('page-gap-inline') &&
      !el.classList.contains('page-gap-cell') &&
      !el.closest('.page-gap-inline, .page-gap-cell, .page-gap-table'),
  )
}

/** Page vertical ranges in ProseMirror local Y (layout px, includes current CSS zoom). */
export function docsPageBounds(pm: HTMLElement, total: number): { top: number; bottom: number }[] {
  const n = Math.max(1, total)
  const pmTop = pm.getBoundingClientRect().top
  const gaps = topLevelPageGaps(pm)
  const bounds: { top: number; bottom: number }[] = []
  let top = 0
  for (let i = 0; i < n; i++) {
    if (i < n - 1 && gaps[i]) {
      const gapRect = gaps[i]!.getBoundingClientRect()
      const bottom = Math.max(top + 8, gapRect.top - pmTop)
      bounds.push({ top, bottom })
      top = gapRect.bottom - pmTop
    } else {
      bounds.push({ top, bottom: Math.max(top + 8, pm.scrollHeight) })
    }
  }
  return bounds
}

/** Scroll the docs canvas so `pageNo` (1-based) is near the top of the scroller. */
export function scrollDocsToPage(pageNo: number): void {
  const scroller = document.querySelector('.editor-scroll') as HTMLElement | null
  const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
  if (!scroller || !pm) return
  if (pageNo <= 1) {
    scroller.scrollTo({ top: 0, behavior: 'smooth' })
    return
  }
  const gaps = topLevelPageGaps(pm)
  const gap = gaps[pageNo - 2]
  if (gap) {
    // Prefer the first content node after the inter-page band (start of that page).
    let target: HTMLElement = gap
    let sib = gap.nextElementSibling as HTMLElement | null
    while (sib && (sib.classList.contains('page-gap') || sib.classList.contains('page-float-host'))) {
      sib = sib.nextElementSibling as HTMLElement | null
    }
    if (sib) target = sib
    const top =
      target.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop
    scroller.scrollTo({ top: Math.max(0, top - 6), behavior: 'smooth' })
    return
  }
  // Fallback before page-gaps finish painting: estimate from --page-h + CSS zoom.
  const pageH = parseFloat(getComputedStyle(pm).getPropertyValue('--page-h')) || 0
  const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
  const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
  if (pageH > 0) {
    scroller.scrollTo({ top: (pageNo - 1) * pageH * zoom, behavior: 'smooth' })
  }
}

function buildPageCloneHtml(
  pm: HTMLElement,
  bound: { top: number; bottom: number },
  pageW: number,
): string {
  const wrap = document.createElement('div')
  wrap.className = 'doc-thumb-page-inner'
  wrap.style.width = `${pageW}px`
  wrap.style.minHeight = `${Math.max(40, bound.bottom - bound.top)}px`
  const pmTop = pm.getBoundingClientRect().top
  for (const child of Array.from(pm.children)) {
    if (!(child instanceof HTMLElement)) continue
    if (child.classList.contains('page-gap') || child.classList.contains('page-float-host')) continue
    const r = child.getBoundingClientRect()
    const top = r.top - pmTop
    const bottom = r.bottom - pmTop
    if (bottom < bound.top + 2 || top > bound.bottom - 2) continue
    const clone = child.cloneNode(true) as HTMLElement
    clone.querySelectorAll('script, .page-gap, .page-float-host').forEach((el) => el.remove())
    wrap.appendChild(clone)
    if (wrap.childElementCount >= 24) break
  }
  return wrap.outerHTML
}

type Props = {
  current: number
  total: number
  /** Paper width at 100% zoom (px). */
  pageWidthPx: number
  /** Paper height at 100% zoom (px). */
  pageHeightPx: number
  /** Bump when pagination / doc content settles so clones refresh. */
  revision: number
  onGoToPage: (page: number) => void
}

/**
 * PDF-style page thumbnail rail for KB embed (`?readOnly=1`).
 * Visuals are CSS-scaled clones of blocks that intersect each page band.
 */
export function PageThumbs({
  current,
  total,
  pageWidthPx,
  pageHeightPx,
  revision,
  onGoToPage,
}: Props) {
  const listRef = useRef<HTMLDivElement>(null)
  const [htmls, setHtmls] = useState<string[]>([])

  const refresh = useCallback(() => {
    const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    if (!pm || total < 1 || pageWidthPx <= 0) {
      setHtmls([])
      return
    }
    const bounds = docsPageBounds(pm, total)
    const next: string[] = []
    for (let i = 0; i < total; i++) {
      const b = bounds[i] ?? { top: 0, bottom: pageHeightPx }
      next.push(buildPageCloneHtml(pm, b, pageWidthPx))
    }
    setHtmls(next)
  }, [total, pageWidthPx, pageHeightPx, revision])

  useEffect(() => {
    const t1 = window.setTimeout(refresh, 160)
    // Page-gaps paint after the first pagination pass — refresh again so thumbs match bands.
    const t2 = window.setTimeout(refresh, 700)
    return () => {
      window.clearTimeout(t1)
      window.clearTimeout(t2)
    }
  }, [refresh])

  useEffect(() => {
    const el = listRef.current?.querySelector<HTMLElement>(`[data-page="${current}"]`)
    el?.scrollIntoView({ block: 'nearest' })
  }, [current])

  const scale = THUMB_W / Math.max(1, pageWidthPx)
  const thumbH = Math.max(40, pageHeightPx * scale)
  const count = Math.max(1, total)

  return (
    <aside className="doc-page-thumbs" aria-label="Page thumbnails">
      <div ref={listRef} className="doc-page-thumbs-list">
        {Array.from({ length: count }, (_, i) => {
          const pageNo = i + 1
          const active = pageNo === current
          return (
            <button
              key={pageNo}
              type="button"
              data-page={pageNo}
              className={`doc-thumb${active ? ' doc-thumb-active' : ''}`}
              onClick={() => onGoToPage(pageNo)}
            >
              <div className="doc-thumb-box" style={{ width: THUMB_W, height: thumbH }}>
                <div
                  className="doc-thumb-scale"
                  style={{
                    width: pageWidthPx,
                    height: pageHeightPx,
                    transform: `scale(${scale})`,
                    transformOrigin: 'top left',
                  }}
                  // Cloned page fragment — not user HTML; stripped of scripts above.
                  dangerouslySetInnerHTML={{ __html: htmls[i] ?? '' }}
                />
              </div>
              <span className="doc-thumb-no">{pageNo}</span>
            </button>
          )
        })}
      </div>
    </aside>
  )
}
