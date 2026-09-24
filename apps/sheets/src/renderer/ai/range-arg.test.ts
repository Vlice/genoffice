/**
 * Run with: npx tsx src/renderer/ai/range-arg.test.ts
 */
import assert from 'node:assert/strict'

import { resolveToolRange } from './range-arg'

const sheets = [
  { id: 'sheet-1', name: 'Sheet1' },
  { id: 'sheet-2', name: 'Sheet2' },
]

assert.equal(resolveToolRange({ range: 'A1:H20' }, undefined, sheets).range, 'A1:H20')

assert.deepEqual(resolveToolRange({ range: { start: 'A1', end: 'H1' } }, undefined, sheets), {
  range: 'A1:H1',
})

assert.equal(resolveToolRange({ parameters: { range: 'B2:D4' } }, undefined, sheets).range, 'B2:D4')

assert.deepEqual(resolveToolRange({}, 'Sheet2!A1:H2', sheets), {
  range: 'A1:H2',
  sheetIdFromQualifier: 'sheet-2',
})

assert.equal(resolveToolRange({ range: '' }, 'C3:E10', sheets).range, 'C3:E10')
assert.equal(resolveToolRange({}, undefined, sheets).range, '')
assert.equal(resolveToolRange({ range: '$A$1:$C$3' }, undefined, sheets).range, 'A1:C3')

console.log('range-arg.test.ts ok')
