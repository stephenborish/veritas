const assert = require('assert');

// Simulate Google Apps Script SpreadsheetApp interactions
class MockRange {
  constructor(row, col, numRows = 1, numCols = 1, sheet) {
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
    this.sheet = sheet;
  }
  setNumberFormat(fmt) {
    this.sheet.apiCalls++;
    return this;
  }
  setValue(val) {
    this.sheet.apiCalls++;
    return this;
  }
  setValues(vals) {
    this.sheet.apiCalls++;
    return this;
  }
}

class MockSheet {
  constructor() {
    this.apiCalls = 0;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    this.apiCalls++;
    return new MockRange(row, col, numRows, numCols, this);
  }
  resetCalls() {
    this.apiCalls = 0;
  }
}

// Variables setup
const i = 5;
const ansStr = "my answer";
const isCorrect = true;
const points = 1;
const dateStr = "2023-01-01T00:00:00Z";
const partialCredit = false;

// Mock values from getDataRange().getValues()
// rd[i][8] corresponds to column 9 (MaxPoints)
const rd = Array(10).fill(0).map(() => Array(15).fill(0));
rd[i][8] = 1; // existing MaxPoints

// 1. Original Code Benchmark
const rSheetOriginal = new MockSheet();
rSheetOriginal.getRange(i+1,6).setNumberFormat('@').setValue(ansStr);
rSheetOriginal.getRange(i+1,7).setValue(isCorrect);
rSheetOriginal.getRange(i+1,8).setValue(points);
rSheetOriginal.getRange(i+1,10).setValue(dateStr);
rSheetOriginal.getRange(i+1,11).setValue(partialCredit);
const originalCalls = rSheetOriginal.apiCalls;

// 2. Optimized Code Benchmark
const rSheetOptimized = new MockSheet();
rSheetOptimized.getRange(i+1, 6).setNumberFormat('@');
rSheetOptimized.getRange(i+1, 6, 1, 6).setValues([[
  ansStr,           // Col 6
  isCorrect,        // Col 7
  points,           // Col 8
  rd[i][8],         // Col 9 (Preserved MaxPoints)
  dateStr,          // Col 10
  partialCredit     // Col 11
]]);
const optimizedCalls = rSheetOptimized.apiCalls;

console.log("=== Performance Benchmark ===");
console.log(`Original API Calls: ${originalCalls}`);
console.log(`Optimized API Calls: ${optimizedCalls}`);
console.log(`Improvement: ${(originalCalls - optimizedCalls) / originalCalls * 100}% reduction in Google Sheets API roundtrips.`);
