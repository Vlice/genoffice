/**
 * Univer float DOMs clamp their screen box into the grid viewport
 * (`transformBound2DOMBound` Math.max's against row/column header sizes). That is
 * fine for in-cell chrome (dropdowns, editors) but wrong for charts/shapes:
 * scrolling shrinks or pins them to the header edge and collapses the hit box
 * so they can no longer be selected.
 *
 * After `addFloatDomToRange`, loosen `boundsOfViewArea` so offsets may go
 * off the viewport; the host's `overflow: hidden` still clips paint, but the
 * hit box must keep the drawing's true size (see `.shape-editable` overflow).
 */
import { SheetCanvasFloatDomManagerService } from '@univerjs/sheets-drawing-ui'

import type { createUniver } from './create-univer'

type UniverRuntime = ReturnType<typeof createUniver>

const OFF_VIEWPORT = Number.NEGATIVE_INFINITY
const PAST_VIEWPORT = Number.POSITIVE_INFINITY

export function relaxFloatDomScrollClamp(runtime: UniverRuntime, floatId: string): void {
  try {
    const service = runtime.univer.__getInjector().get(SheetCanvasFloatDomManagerService)
    const info = service.getFloatDomInfo(floatId)
    const bounds = info?.boundsOfViewArea
    if (!bounds) return
    bounds.left = OFF_VIEWPORT
    bounds.top = OFF_VIEWPORT
    bounds.right = PAST_VIEWPORT
    bounds.bottom = PAST_VIEWPORT
  } catch {
    /* injector / service unavailable during teardown */
  }
}
