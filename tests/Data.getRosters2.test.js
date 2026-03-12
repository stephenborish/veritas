const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert');

global.Logger = { log: console.log };

const codeGsPath = path.join(__dirname, '../Data.gs');
let codeGs = fs.readFileSync(codeGsPath, 'utf8');

codeGs = codeGs.replace('const DB = {', 'global.DB = {');
eval(codeGs);

// Generate huge rows
const rows = [['Block', 'CourseID', 'StudentsJSON', 'UpdatedAt']];
for (let i = 1; i <= 10000; i++) {
  // 50 students per roster
  const stu = Array.from({length: 50}, (_, j) => ({id: i*100+j, name: 'Student ' + (i*100+j)}));
  rows.push([
    'block_' + i,
    'course_1',
    JSON.stringify(stu),
    '2023-01-01T00:00:00Z'
  ]);
}

DB.sh = function(sheetName) {
  return {
    getDataRange: function() {
      return { getValues: function() { return rows; } }
    }
  }
};

test('Benchmark getRosters baseline', () => {
  const start = process.hrtime.bigint();
  const r = DB.getRosters();
  const end = process.hrtime.bigint();
  console.log('Baseline:', Number(end - start) / 1000000, 'ms');
  assert.strictEqual(Object.keys(r).length, 10000);
});
