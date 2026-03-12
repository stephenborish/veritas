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

DB.getRostersOptimized = function() {
  const d=this.sh('Rosters').getDataRange().getValues();
  const r={};
  for(let i=1;i<d.length;i++){
    const rawJSON = d[i][2] || '[]';
    const blockId = d[i][0];
    const courseId = d[i][1];
    const updatedAt = d[i][3];

    // Use property getters to lazily parse students and count
    Object.defineProperty(r, blockId, {
      get: function() {
        const stu = JSON.parse(rawJSON);
        return {
          block: blockId,
          courseId: courseId,
          students: stu,
          count: stu.length,
          updatedAt: updatedAt
        };
      },
      enumerable: true
    });
  }
  return r;
};

test('Benchmark getRosters lazy', () => {
  const start = process.hrtime.bigint();
  const r = DB.getRostersOptimized();
  const end = process.hrtime.bigint();
  console.log('Optimized creation:', Number(end - start) / 1000000, 'ms');

  const start2 = process.hrtime.bigint();
  const block100 = r['block_100'];
  const end2 = process.hrtime.bigint();
  console.log('Optimized access single:', Number(end2 - start2) / 1000000, 'ms');

  assert.strictEqual(Object.keys(r).length, 10000);
});
