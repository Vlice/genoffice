import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { CompareDialog } from '../src/renderer/components/CompareDialog'
import type { CompareDocList, OpenFileResult } from '../src/shared/ipc'

const sampleFile = (name: string): OpenFileResult => ({
  path: `/tmp/${name}`,
  name,
  data: new ArrayBuffer(8),
  hash: 'abc',
})

describe('CompareDialog', () => {
  let host: HTMLDivElement
  let root: Root
  const onPicked = vi.fn<(file: OpenFileResult) => void>()
  const onCancel = vi.fn()
  const listCompareDocs = vi.fn(async (): Promise<CompareDocList> => ({
    source: 'cloud',
    items: [
      { id: '1', name: '规格.docx' },
      { id: '2', name: '需求.docx' },
    ],
  }))
  const readCompareDocx = vi.fn(async (id: string) => sampleFile(`cloud-${id}.docx`))
  const pickCompareDocx = vi.fn(async () => sampleFile('local.docx'))

  beforeEach(async () => {
    onPicked.mockReset()
    onCancel.mockReset()
    listCompareDocs.mockClear()
    readCompareDocx.mockClear()
    pickCompareDocx.mockClear()
    window.desktop = {
      ...(window.desktop as object),
      listCompareDocs,
      readCompareDocx,
      pickCompareDocx,
    } as typeof window.desktop
    host = document.createElement('div')
    document.body.appendChild(host)
    await act(async () => {
      root = createRoot(host)
      root.render(createElement(CompareDialog, { excludeId: '1', onPicked, onCancel }))
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
  })

  afterEach(async () => {
    await act(async () => root.unmount())
    host.remove()
  })

  it('lists cloud documents except the open one', () => {
    const names = [...host.querySelectorAll('.compare-doc-name')].map((el) => el.textContent)
    expect(names).toEqual(['需求.docx'])
    expect(host.querySelector('.btn-primary')).toHaveProperty('disabled', true)
  })

  it('compares the selected cloud document', async () => {
    await act(async () => {
      host.querySelector<HTMLInputElement>('input[type="radio"]')!.click()
    })
    expect(host.querySelector('.btn-primary')).toHaveProperty('disabled', false)
    await act(async () => {
      host.querySelector<HTMLButtonElement>('.btn-primary')!.click()
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    expect(readCompareDocx).toHaveBeenCalledWith('2')
    expect(onPicked).toHaveBeenCalledWith(expect.objectContaining({ name: 'cloud-2.docx' }))
  })

  it('compares a locally picked file', async () => {
    await act(async () => {
      host.querySelector<HTMLButtonElement>('.compare-pick-local')!.click()
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    await act(async () => {
      host.querySelector<HTMLButtonElement>('.btn-primary')!.click()
    })
    expect(pickCompareDocx).toHaveBeenCalled()
    expect(onPicked).toHaveBeenCalledWith(expect.objectContaining({ name: 'local.docx' }))
  })
})
