const { performance } = require('perf_hooks');

class MockRange {
  setValue(val) {
    return new Promise(resolve => setTimeout(resolve, 50));
  }
  setValues(vals) {
    return new Promise(resolve => setTimeout(resolve, 75));
  }
}

class MockSheet {
  constructor(numRows) {
    this.rows = Array.from({ length: numRows }, (_, i) => [
      'sess123', `stu${i % 20}`, `Student ${i}`, 'q1',
      0, 5, '', 'Old answer', new Date().toISOString(), false, '', '', ''
    ]);
  }

  getDataRange() {
    return {
      getValues: () => JSON.parse(JSON.stringify(this.rows))
    };
  }

  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange();
  }
}

const DB = {
  withLock: async (cb) => {
    // Simulate lock overhead
    await new Promise(resolve => setTimeout(resolve, 10));
    return cb();
  }
};

async function runBenchmark() {
  const NUM_ROWS_IN_SHEET = 500;
  const BATCH_SIZE = 10;

  console.log(`\n📊 Running regrade benchmark for ${BATCH_SIZE} students in a sheet with ${NUM_ROWS_IN_SHEET} rows...`);

  const mockBatch = Array.from({ length: BATCH_SIZE }, (_, i) => ({
    studentId: `stu${i}`,
    answer: 'New answer'
  }));

  const resultMap = {};
  for (const r of mockBatch) {
    resultMap[r.studentId] = { score: 5, feedback: 'Good' };
  }

  // --- BASELINE (N+1 Write Operations with O(N) Scans) ---
  const sheetBaseline = new MockSheet(NUM_ROWS_IN_SHEET);
  const startBaseline = performance.now();

  for (const r of mockBatch) {
    const result = resultMap[r.studentId];
    await DB.withLock(async () => {
      const rows = sheetBaseline.getDataRange().getValues();
      let found = false;
      for (let j = 1; j < rows.length; j++) {
        // O(N) row scanning delay simulation
        if (rows[j][0] === 'sess123' && rows[j][1] === r.studentId && rows[j][3] === 'q1') {
          await sheetBaseline.getRange(j + 1, 5).setValue(result.score);
          await sheetBaseline.getRange(j + 1, 7).setValue(result.feedback);
          await sheetBaseline.getRange(j + 1, 13).setValue('safeCtx');
          found = true; break;
        }
      }
    });
  }

  const endBaseline = performance.now();
  const durationBaseline = endBaseline - startBaseline;
  console.log(`⏱️  Baseline (N+1 Writes + O(N) Scans): ${durationBaseline.toFixed(2)} ms`);

  // --- OPTIMIZED (1 Lock, O(1) Lookups, 1 setValues per row) ---
  const sheetOptimized = new MockSheet(NUM_ROWS_IN_SHEET);
  const startOptimized = performance.now();

  await DB.withLock(async () => {
    const rows = sheetOptimized.getDataRange().getValues();
    const numCols = rows[0].length;

    // O(1) Index
    const rowIndexMap = new Map();
    for (let j = 1; j < rows.length; j++) {
      if (rows[j][0] === 'sess123' && rows[j][3] === 'q1') {
        rowIndexMap.set(rows[j][1], j);
      }
    }

    for (const r of mockBatch) {
      const result = resultMap[r.studentId];
      if (!result) continue;

      const j = rowIndexMap.get(r.studentId);
      if (j !== undefined) {
        rows[j][4] = result.score;
        rows[j][6] = result.feedback;
        rows[j][12] = 'safeCtx';
        await sheetOptimized.getRange(j + 1, 1, 1, numCols).setValues([rows[j]]);
      }
    }
  });

  const endOptimized = performance.now();
  const durationOptimized = endOptimized - startOptimized;
  console.log(`⏱️  Optimized (1 Lock, O(1) Map, Batched Writes): ${durationOptimized.toFixed(2)} ms`);

  const improvement = durationBaseline - durationOptimized;
  const percentage = (improvement / durationBaseline) * 100;

  console.log(`\n🚀 Improvement: ${improvement.toFixed(2)} ms (${percentage.toFixed(2)}% faster)\n`);
}

runBenchmark().catch(console.error);
