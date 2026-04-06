const assert = require('assert');

// Simulate Google Apps Script SpreadsheetApp interactions
class MockSheet {
  constructor() {
    this.apiCalls = 0;
    this.rows = [];
    this.maxRows = 100;
  }
  getLastRow() {
    this.apiCalls++;
    return this.rows.length;
  }
  getMaxRows() {
    this.apiCalls++;
    return this.maxRows;
  }
  insertRowsAfter(row, numRows) {
    this.apiCalls++;
    this.maxRows += numRows;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    this.apiCalls++;
    return new MockRange(this);
  }
  resetCalls() {
    this.apiCalls = 0;
  }
}

class MockRange {
  constructor(sheet) {
    this.sheet = sheet;
  }
  setValues(vals) {
    this.sheet.apiCalls++;
    this.sheet.rows.push(...vals);
    return this;
  }
}

// Variables setup
const batchSize = 5;
const newGradeRows = [
  ['sess1', 'stu1', 'Bob', 'q1', 1, 1, 'good', 'ans1', 'date', false, '', '', ''],
  ['sess1', 'stu2', 'Alice', 'q1', 1, 1, 'good', 'ans2', 'date', false, '', '', ''],
  ['sess1', 'stu3', 'Charlie', 'q1', 1, 1, 'good', 'ans3', 'date', false, '', '', ''],
  ['sess1', 'stu4', 'Dave', 'q1', 1, 1, 'good', 'ans4', 'date', false, '', '', ''],
  ['sess1', 'stu5', 'Eve', 'q1', 1, 1, 'good', 'ans5', 'date', false, '', '', '']
];

// Helper from Grader
function _batchAppendRows(sheet, rows) {
  if (!rows || !rows.length) return;
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  const neededRows = (lastRow + rows.length) - maxRows;
  if (neededRows > 0) {
    sheet.insertRowsAfter(maxRows, neededRows);
  }
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// 1. Original Code Benchmark (Inside Fallback Loop)
const sheetOriginal = new MockSheet();
for (const r of newGradeRows) {
  _batchAppendRows(sheetOriginal, [r]);
}
const originalCalls = sheetOriginal.apiCalls;

// 2. Optimized Code Benchmark (Outside Fallback Loop)
const sheetOptimized = new MockSheet();
_batchAppendRows(sheetOptimized, newGradeRows);
const optimizedCalls = sheetOptimized.apiCalls;

console.log("=== Fallback Loop Append Benchmark ===");
console.log(`Original API Calls (${batchSize} fallbacks): ${originalCalls}`);
console.log(`Optimized API Calls (${batchSize} fallbacks): ${optimizedCalls}`);
console.log(`Improvement: ${((originalCalls - optimizedCalls) / originalCalls * 100).toFixed(2)}% reduction in Google Sheets API roundtrips.`);
