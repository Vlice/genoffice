import { describe, expect, it } from 'vitest'
import { fontPickPatch, isEastAsianFontName } from '../src/renderer/font-list'

describe('isEastAsianFontName', () => {
  it('classifies CJK families across naming conventions', () => {
    for (const f of [
      'KaiTi',
      'SimSun',
      '宋体',
      '楷体_GB2312',
      '仿宋',
      '微软雅黑',
      '等线',
      '方正小标宋简体',
      'PingFang SC',
      'Microsoft YaHei',
      'Yu Mincho',
      'MS Gothic',
      'Malgun Gothic',
      'Batang',
      'Microsoft JhengHei',
      'PMingLiU',
      'Noto Sans CJK SC',
      'Source Han Serif',
    ]) {
      expect(isEastAsianFontName(f), f).toBe(true)
    }
  })

  it('classifies Latin families as non-East-Asian', () => {
    for (const f of [
      'Times New Roman',
      'Arial',
      'Calibri',
      'Cambria',
      'Georgia',
      'Courier New',
      'Segoe UI',
      'Consolas',
      'Garamond',
    ]) {
      expect(isEastAsianFontName(f), f).toBe(false)
    }
  })
})

describe('fontPickPatch', () => {
  it('writes both rFonts slots for a CJK face so digits and letters follow', () => {
    expect(fontPickPatch('仿宋')).toEqual({ font: '仿宋', fontAscii: '仿宋' })
    expect(fontPickPatch('KaiTi')).toEqual({ font: 'KaiTi', fontAscii: 'KaiTi' })
  })

  it('writes only the ascii slot for a Latin face', () => {
    expect(fontPickPatch('Impact')).toEqual({ fontAscii: 'Impact' })
    expect(fontPickPatch('Arial')).toEqual({ fontAscii: 'Arial' })
  })

  it('clears both slots when the pick is empty', () => {
    expect(fontPickPatch(null)).toEqual({ font: null, fontAscii: null })
  })
})
