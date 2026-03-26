const assert = require('assert');

// Simulate Google Apps Script
class MockRange {
  constructor(row, col, numRows = 1, numCols = 1, sheet) {
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
    this.sheet = sheet;
  }
  getValues() {
    this.sheet.apiCalls++;
    this.sheet.bytesTransferred += this.numRows * this.numCols;

    // Return dummy data
    const data = [];
    for (let i = 0; i < this.numRows; i++) {
      const rowArr = [];
      for (let j = 0; j < this.numCols; j++) {
        if (j === 0) rowArr.push('sess_123');
        else if (j === 1) rowArr.push(`stu_${i}`);
        else rowArr.push('data');
      }
      data.push(rowArr);
    }
    return data;
  }
  setValue(val) {
    this.sheet.apiCalls++;
    return this;
  }
}

class MockSheet {
  constructor(totalRows, totalCols) {
    this.totalRows = totalRows;
    this.totalCols = totalCols;
    this.apiCalls = 0;
    this.bytesTransferred = 0;
  }
  getDataRange() {
    return new MockRange(1, 1, this.totalRows, this.totalCols, this);
  }
  getLastRow() {
    return this.totalRows;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange(row, col, numRows, numCols, this);
  }
  resetCalls() {
    this.apiCalls = 0;
    this.bytesTransferred = 0;
  }
}

const TOTAL_ROWS = 1000;
const TOTAL_COLS = 20;
const TARGET_SESS = 'sess_123';
const TARGET_STU = 'stu_999';
const Q_INDEX = 5;

// 1. Original
const sheetOrg = new MockSheet(TOTAL_ROWS, TOTAL_COLS);
const dataOrg = sheetOrg.getDataRange().getValues();
for (let i = 1; i < dataOrg.length; i++) {
  if (dataOrg[i][0] === TARGET_SESS && dataOrg[i][1] === TARGET_STU) {
    sheetOrg.getRange(i + 1, 17).setValue(Q_INDEX);
    break;
  }
}
const orgCalls = sheetOrg.apiCalls;
const orgTransferred = sheetOrg.bytesTransferred;

// 2. Optimized
const sheetOpt = new MockSheet(TOTAL_ROWS, TOTAL_COLS);
const lastRow = sheetOpt.getLastRow();
if (lastRow > 1) {
  const dataOpt = sheetOpt.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < dataOpt.length; i++) {
    if (dataOpt[i][0] === TARGET_SESS && dataOpt[i][1] === TARGET_STU) {
      sheetOpt.getRange(i + 2, 17).setValue(Q_INDEX);
      break;
    }
  }
}
const optCalls = sheetOpt.apiCalls;
const optTransferred = sheetOpt.bytesTransferred;

console.log("=== Benchmark Results ===");
console.log(`Original API Calls: ${orgCalls}, Cells Transferred: ${orgTransferred}`);
console.log(`Optimized API Calls: ${optCalls}, Cells Transferred: ${optTransferred}`);
const reduction = ((orgTransferred - optTransferred) / orgTransferred) * 100;
console.log(`Improvement: ${reduction.toFixed(2)}% reduction in data payload size.`);
