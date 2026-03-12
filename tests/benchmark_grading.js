const { performance } = require('perf_hooks');

// Simulating Google Apps Script's SpreadsheetApp behavior and network latency
class MockRange {
  setValue(val) {
    // Simulated network delay for a single cell write
    return new Promise(resolve => setTimeout(resolve, 50));
  }
  setValues(vals) {
    // Simulated network delay for a bulk write (slightly larger but mostly constant time overhead)
    return new Promise(resolve => setTimeout(resolve, 75));
  }
}

class MockSheet {
  constructor() {
    this.rows = [];
    this.lastRow = 0;
  }

  // N+1 Anti-pattern
  async appendRow(row) {
    this.rows.push(row);
    this.lastRow++;
    // Simulated network delay for appending a single row
    await new Promise(resolve => setTimeout(resolve, 60));
  }

  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange();
  }

  getLastRow() {
    return this.lastRow;
  }
}

async function runBenchmark() {
  const NUM_STUDENTS = 20;
  console.log(`\n📊 Running benchmark for grading ${NUM_STUDENTS} students...`);

  const mockData = Array.from({ length: NUM_STUDENTS }, (_, i) => [
    'sess123', `stu${i}`, `Student ${i}`, 'q1',
    5, 5, 'Great job!',
    'This is a good answer.', new Date().toISOString(),
    false, '', '', ''
  ]);

  // --- BASELINE (N+1 Write Operations) ---
  const sheetBaseline = new MockSheet();
  const startBaseline = performance.now();

  for (const row of mockData) {
    await sheetBaseline.appendRow(row);
    await sheetBaseline.getRange(sheetBaseline.getLastRow(), 8).setValue(row[4]);
  }

  const endBaseline = performance.now();
  const durationBaseline = endBaseline - startBaseline;
  console.log(`⏱️  Baseline (N+1 Writes): ${durationBaseline.toFixed(2)} ms`);

  // --- OPTIMIZED (Batch Writes) ---
  const sheetOptimized = new MockSheet();
  const startOptimized = performance.now();

  // In the real code, we collect these in the loop
  let newRows = [];

  for (const row of mockData) {
    newRows.push(row);
    // Note: We still simulate the individual .setValue() calls on Responses sheet
    // as bulk updating sparse, non-contiguous rows natively in standard GAS is tricky
    // without Sheets API Advanced Service or rewriting the entire sheet.
    // However, we avoid the .appendRow() delay entirely inside the loop.
    await sheetOptimized.getRange(sheetOptimized.getLastRow() + newRows.length, 8).setValue(row[4]);
  }

  // Then do a single bulk append outside the loop
  if (newRows.length > 0) {
    await sheetOptimized.getRange(sheetOptimized.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    sheetOptimized.lastRow += newRows.length;
  }

  const endOptimized = performance.now();
  const durationOptimized = endOptimized - startOptimized;
  console.log(`⏱️  Optimized (Batched Appends): ${durationOptimized.toFixed(2)} ms`);

  const improvement = durationBaseline - durationOptimized;
  const percentage = (improvement / durationBaseline) * 100;

  console.log(`\n🚀 Improvement: ${improvement.toFixed(2)} ms (${percentage.toFixed(2)}% faster)\n`);
}

runBenchmark().catch(console.error);
