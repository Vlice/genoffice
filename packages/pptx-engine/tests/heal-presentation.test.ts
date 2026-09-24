import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createBlankPptx } from '../src/blank'
import { openPptx, savePptx } from '../src/index'

describe('heal missing sldMasterIdLst', () => {
  it('injects sldMasterIdLst on open so WPS-valid save', async () => {
    const good = await createBlankPptx()
    const z = await JSZip.loadAsync(good)
    let pres = await z.file('ppt/presentation.xml')!.async('string')
    pres = pres.replace(/<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>/, '')
    expect(pres).not.toMatch(/sldMasterIdLst/)
    z.file('ppt/presentation.xml', pres)
    const broken = await z.generateAsync({ type: 'uint8array' })

    const opened = await openPptx(broken)
    const out = await savePptx(opened)
    const z2 = await JSZip.loadAsync(out)
    const healed = await z2.file('ppt/presentation.xml')!.async('string')
    expect(healed).toMatch(/<p:sldMasterIdLst>/)
    expect(healed).toMatch(/r:id="rId1"/)
  })
})
