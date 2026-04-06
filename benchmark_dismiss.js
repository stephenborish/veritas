const assert = require('assert');

class MockFinder {
  constructor(vals, text) {
    this.vals = vals;
    this.text = text;
    this.currIdx = 0;
  }
  matchEntireCell() { return this; }
  findNext() {
    for (let i = this.currIdx; i < this.vals.length; i++) {
      if (String(this.vals[i][4]) === String(this.text)) {
        this.currIdx = i + 1; // Start from next row on next call
        return { getRow: () => i + 1 };
      }
    }
    return null;
  }
}

class MockRange {
  constructor(vals) { this.vals = vals; }
  getValues() { return this.vals; }
  getValue() { return this.vals[0][0]; }
  createTextFinder(text) {
    return new MockFinder(this.vals, text);
  }
}

class MockSheet {
  constructor(data) {
    this.data = data;
    this.apiCalls = 0;
    this.transfers = 0;
  }
  getDataRange() {
    this.apiCalls++;
    this.transfers += this.data.length * this.data[0].length;
    return new MockRange(this.data);
  }
  getRange(row, col) {
    if (col === 5) {
      this.apiCalls++;
      // Return a range encompassing all of col 5 (using full data for mock simplicity)
      return new MockRange(this.data);
    }
    this.apiCalls++;
    this.transfers += 1; // getValue overhead
    return new MockRange([[this.data[row-1][col-1]]]);
  }
  deleteRow() {
    this.apiCalls++;
  }
}

// 10,000 rows, simulating a large Violations sheet
// Let's insert multiple matches for the same timestamp, but only the 3rd one has the correct session
const largeData = Array(10000).fill(0).map((_, i) => ['sess1', 'stu1', 'John', 'tab_switch', `2023-01-01T${i}`, false]);

// Target is at the end, and we'll insert a few mock duplicates before it
const targetTimestamp = `2023-01-01T9999`;
const sessionId = 'target_sess';

// Insert duplicate timestamps that will fail the sessionId check
largeData[5000] = ['wrong_sess1', 'stu2', 'Jane', 'tab_switch', targetTimestamp, false];
largeData[7000] = ['wrong_sess2', 'stu3', 'Bob', 'tab_switch', targetTimestamp, false];
// The correct match
largeData[9000] = [sessionId, 'stu4', 'Alice', 'tab_switch', targetTimestamp, false];


// Benchmark 1: Original
const sheet1 = new MockSheet(largeData);
const start1 = process.hrtime.bigint();
const violData = sheet1.getDataRange().getValues();
for (let i = 1; i < violData.length; i++) {
  if (violData[i][0] === sessionId && String(violData[i][4]) === String(targetTimestamp)) {
    sheet1.deleteRow(i + 1);
    break;
  }
}
const end1 = process.hrtime.bigint();
const time1 = Number(end1 - start1) / 1e6;

// Benchmark 2: Optimized (with while loop logic exactly as implemented in Data.gs)
const sheet2 = new MockSheet(largeData);
const start2 = process.hrtime.bigint();

const timestampCol = sheet2.getRange(1, 5); // Mock getRange('E:E')
const textFinder = timestampCol.createTextFinder(String(targetTimestamp)).matchEntireCell(true);
let searchResult = textFinder.findNext();
let findAttempts = 0;
while (searchResult) {
  findAttempts++;
  const rowIndex = searchResult.getRow();
  if (rowIndex > 1 && String(sheet2.getRange(rowIndex, 1).getValue()) === String(sessionId)) {
    sheet2.deleteRow(rowIndex);
    break;
  }
  searchResult = textFinder.findNext();
}
const end2 = process.hrtime.bigint();
const time2 = Number(end2 - start2) / 1e6;

console.log("=== dismissViolation Optimization Benchmark ===");
console.log(`Original Time (simulated): ${time1.toFixed(3)} ms, Cells Transferred: ${sheet1.transfers}`);
console.log(`Optimized Time (simulated): ${time2.toFixed(3)} ms, Cells Transferred: ${sheet2.transfers}`);
console.log(`findNext() iterations required to match target: ${findAttempts}`);
console.log(`Cell Transfer Reduction: ${((sheet1.transfers - sheet2.transfers) / sheet1.transfers * 100).toFixed(2)}%`);
