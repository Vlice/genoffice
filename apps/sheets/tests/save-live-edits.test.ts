import { describe, expect, it } from 'vitest'

import {
  applyPaintedRunStyle,
  cellPlainValue,
  collectLiveValueEdits,
  liveCellStyle,
  mergeEditsPreferLive,
  univerStyleToEdit,
} from '../src/renderer/save-live-edits'
import type { WorkbookCellEdit } from '../src/shared/desktop-api'

const cell = (
  row: number,
  column: number,
  extra: Partial<WorkbookCellEdit> = {},
): WorkbookCellEdit => ({
  sheetId: 'sheet-1',
  row,
  column,
  writeValue: true,
  value: extra.value ?? 'x',
  ...extra,
})

describe('univerStyleToEdit', () => {
  it('treats boolean and numeric bold flags the same', () => {
    expect(univerStyleToEdit({ bl: true })?.bold).toBe(true)
    expect(univerStyleToEdit({ bl: 1 })?.bold).toBe(true)
    expect(univerStyleToEdit({ bl: false })?.bold).toBe(false)
    expect(univerStyleToEdit({ bl: 0 })?.bold).toBe(false)
  })

  it('ignores the empty-rgb fill sentinel used to block column styles', () => {
    expect(univerStyleToEdit({ bg: { rgb: '' } })).toBeUndefined()
  })
})

describe('liveCellStyle', () => {
  it('resolves hashed and numeric style ids from the snapshot table', () => {
    const table = {
      s1: { bl: 1, ff: 'Aptos' },
      '2': { bl: 0, ff: 'Arial' },
    }
    expect(liveCellStyle('s1', table)).toEqual({ bold: true, fontFamily: 'Aptos' })
    expect(liveCellStyle(2, table)?.bold).toBe(false)
    expect(liveCellStyle(2, table)?.fontFamily).toBe('Arial')
  })
})

describe('cellPlainValue', () => {
  it('reads typed text from a rich document when v is empty', () => {
    expect(
      cellPlainValue({
        p: { body: { dataStream: '景区名称\r\n' } },
      }),
    ).toBe('景区名称')
  })

  it('prefers a real v over the document stream', () => {
    expect(
      cellPlainValue({
        v: 'plain',
        p: { body: { dataStream: 'ignored\r\n' } },
      }),
    ).toBe('plain')
  })
})

describe('mergeEditsPreferLive', () => {
  it('keeps a journaled header bold when the live snapshot still says not-bold', () => {
    const merged = mergeEditsPreferLive(
      [cell(0, 0, { value: '景区名称', style: { bold: true } })],
      [cell(0, 0, { value: '景区名称', style: { bold: false, fontFamily: 'Aptos' } })],
    )
    expect(merged).toHaveLength(1)
    expect(merged[0]?.style).toEqual({ bold: true, fontFamily: 'Aptos' })
  })

  it('fills live-only styles onto cells the journal never formatted', () => {
    const merged = mergeEditsPreferLive(
      [cell(1, 0, { value: 'AAAAA' })],
      [cell(1, 0, { value: 'AAAAA', style: { fontFamily: 'Aptos' } })],
    )
    expect(merged[0]?.style).toEqual({ fontFamily: 'Aptos' })
  })

  it('uses live style only when the journal never recorded one', () => {
    const merged = mergeEditsPreferLive(
      [cell(1, 0, { value: '九寨沟' })],
      [cell(1, 0, { value: '九寨沟' })],
    )
    expect(merged[0]?.style).toBeUndefined()
  })

  it('lets painted runs un-bold data when the shared xf says bold', () => {
    const run = (bold: boolean) => ({
      text: '九寨沟',
      bold,
      italic: false,
      underline: false,
      strikethrough: false,
      family: 'Aptos',
    })
    const merged = mergeEditsPreferLive(
      [],
      [cell(1, 0, { value: '九寨沟', style: { bold: true, fontFamily: 'Aptos' }, rich: [run(false)] })],
    )
    expect(merged[0]?.style?.bold).toBe(false)
  })

  it('keeps a journaled header bold over stale regular runs', () => {
    const run = (bold: boolean) => ({
      text: '景区名称',
      bold,
      italic: false,
      underline: false,
      strikethrough: false,
      family: 'Aptos',
    })
    const merged = mergeEditsPreferLive(
      [cell(0, 0, { value: '景区名称', style: { bold: true }, rich: [run(true)] })],
      [cell(0, 0, { value: '景区名称', style: { bold: false }, rich: [run(false)] })],
    )
    expect(merged[0]?.style?.bold).toBe(true)
  })

  it('keeps a pasted run face over a journaled theme font', () => {
    const run = (family: string) => ({
      text: '张伟',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      family,
    })
    const merged = mergeEditsPreferLive(
      [cell(1, 0, { value: '张伟', style: { fontFamily: 'Aptos', bold: true }, rich: [run('Aptos')] })],
      [cell(1, 0, { value: '张伟', style: { fontFamily: '微软雅黑' }, rich: [run('微软雅黑')] })],
    )
    expect(merged[0]?.style?.fontFamily).toBe('微软雅黑')
    expect(merged[0]?.style?.bold).toBe(true)
    expect(merged[0]?.rich?.[0]?.family).toBe('微软雅黑')
  })
})

describe('collectLiveValueEdits', () => {
  it('dumps composed getCellStyleData and document text', () => {
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSnapshot: () => ({
            styles: { shared: { bl: 0, ff: 'Calibri' } },
            sheets: {
              'sheet-1': {
                cellData: {
                  0: {
                    0: { s: 'shared', p: { body: { dataStream: '景区名称\r\n' } } },
                  },
                  1: {
                    0: { v: '九寨沟', s: 'shared' },
                  },
                },
              },
            },
          }),
          getSheetBySheetId: () => ({
            getRange: (row: number) => ({
              getCellStyleData: () => (row === 0 ? { bl: 1, ff: 'Aptos', fs: 10 } : { bl: 0, ff: 'Aptos', fs: 10 }),
            }),
          }),
        }),
      },
    }
    const edits = collectLiveValueEdits(runtime as never)
    const header = edits.find((e) => e.row === 0)
    const data = edits.find((e) => e.row === 1)
    expect(header?.value).toBe('景区名称')
    expect(header?.style).toMatchObject({ bold: true, fontFamily: 'Aptos', fontSize: 10 })
    expect(data?.value).toBe('九寨沟')
    expect(data?.style?.bold).toBe(false)
  })

  it('lifts a pasted rich-run face over the composed theme font', () => {
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSnapshot: () => ({
            styles: { theme: { ff: 'Aptos' } },
            sheets: {
              'sheet-1': {
                cellData: {
                  0: {
                    0: {
                      v: '张伟',
                      s: 'theme',
                      p: {
                        body: {
                          dataStream: '张伟\r\n',
                          textRuns: [{ st: 0, ed: 2, ts: { ff: '微软雅黑' } }],
                        },
                      },
                    },
                  },
                },
              },
            },
          }),
          getSheetBySheetId: () => ({
            getRange: () => ({
              getCellStyleData: (type?: string) =>
                type === 'cell' ? { ff: 'Aptos' } : { ff: 'Aptos', fs: 11 },
            }),
          }),
        }),
      },
    }
    const edits = collectLiveValueEdits(runtime as never)
    expect(edits[0]?.style?.fontFamily).toBe('微软雅黑')
    expect(edits[0]?.rich?.[0]?.family).toBe('微软雅黑')
  })

  it('follows painted runs when a shared style id inverts header bold onto data', () => {
    const textRun = (text: string, bold: boolean) => ({
      st: 0,
      ed: text.length,
      ts: { bl: bold ? 1 : 0, ff: 'Aptos' },
    })
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSnapshot: () => ({
            styles: { shared: { bl: 1, ff: 'Arial' } },
            sheets: {
              'sheet-1': {
                cellData: {
                  0: {
                    0: {
                      v: '景区名称',
                      s: 'shared',
                      p: { body: { dataStream: '景区名称\r\n', textRuns: [textRun('景区名称', true)] } },
                    },
                  },
                  1: {
                    0: {
                      v: '九寨沟',
                      s: 'shared',
                      p: { body: { dataStream: '九寨沟\r\n', textRuns: [textRun('九寨沟', false)] } },
                    },
                  },
                },
              },
            },
          }),
          getSheetBySheetId: () => ({
            getRange: (row: number) => ({
              // Cell-own style still says the shared xf is bold for both rows.
              getCellStyleData: () => ({ bl: 1, ff: 'Arial' }),
              _row: row,
            }),
          }),
        }),
      },
    }
    const edits = collectLiveValueEdits(runtime as never)
    const header = edits.find((e) => e.row === 0)
    const data = edits.find((e) => e.row === 1)
    expect(header?.style?.bold).toBe(true)
    expect(data?.style?.bold).toBe(false)
    expect(header?.style?.fontFamily).toBe('Aptos')
    expect(data?.style?.fontFamily).toBe('Aptos')
    expect(applyPaintedRunStyle(cell(1, 0, { style: { bold: true }, rich: data?.rich }))?.style?.bold).toBe(
      false,
    )
  })
})
