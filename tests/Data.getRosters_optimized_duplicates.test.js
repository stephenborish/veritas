const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert');

global.Logger = { log: console.log };

const codeGsPath = path.join(__dirname, '../Data.gs');
let codeGs = fs.readFileSync(codeGsPath, 'utf8');

codeGs = codeGs.replace('const DB = {', 'global.DB = {');
eval(codeGs);

// Generate mock rows with explicit duplicates
const rows = [['Block', 'CourseID', 'StudentsJSON', 'UpdatedAt']];
for (let i = 1; i <= 50; i++) {
  // 50 students per roster
  const stu = Array.from({length: 5}, (_, j) => ({id: i*100+j, name: 'Student ' + (i*100+j)}));
  rows.push([
    'block_' + i,
    'course_1',
    JSON.stringify(stu),
    '2023-01-01T00:00:00Z'
  ]);
}
// Add duplicates at the end, these should overwrite the initial values
for (let i = 1; i <= 5; i++) {
  const stu = Array.from({length: 2}, (_, j) => ({id: i*1000+j, name: 'Student Updated ' + (i*1000+j)}));
  rows.push([
    'block_' + i,
    'course_2',
    JSON.stringify(stu),
    '2023-01-02T00:00:00Z'
  ]);
}

DB.sh = function(sheetName) {
  return {
    getDataRange: function() {
      return { getValues: function() { return rows; } }
    }
  }
};

DB.getRostersBaseline = function() {
  const d=this.sh('Rosters').getDataRange().getValues(); const r={};
  for(let i=1;i<d.length;i++){const stu=JSON.parse(d[i][2]||'[]');r[d[i][0]]={block:d[i][0],courseId:d[i][1],students:stu,count:stu.length,updatedAt:d[i][3]};}
  return r;
};

test('Verify optimized getRosters handles duplicates exactly like baseline', () => {
  const rBaseline = DB.getRostersBaseline();
  const rOptimized = DB.getRosters();

  // They should be equivalent
  assert.strictEqual(Object.keys(rBaseline).length, 50);
  assert.strictEqual(Object.keys(rOptimized).length, 50);

  const b1_baseline = rBaseline['block_1'];
  const b1_optimized = rOptimized['block_1'];

  assert.strictEqual(b1_baseline.count, 2);
  assert.strictEqual(b1_optimized.count, 2);

  assert.strictEqual(b1_baseline.courseId, 'course_2');
  assert.strictEqual(b1_optimized.courseId, 'course_2');

  assert.strictEqual(b1_baseline.updatedAt, '2023-01-02T00:00:00Z');
  assert.strictEqual(b1_optimized.updatedAt, '2023-01-02T00:00:00Z');

  assert.deepStrictEqual(b1_baseline.students, b1_optimized.students);

  // Non-duplicates should be preserved
  assert.strictEqual(rBaseline['block_50'].count, 5);
  assert.strictEqual(rOptimized['block_50'].count, 5);
});
