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
for (let i = 1; i <= 1000; i++) {
  // 50 students per roster
  const stu = Array.from({length: 5}, (_, j) => ({id: i*100+j, name: 'Student ' + (i*100+j)}));
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

test('Verify optimized getRosters works', () => {
  const r2 = DB.getRosters();

  // They should be equivalent
  assert.strictEqual(Object.keys(r2).length, 1000);

  const b1 = r2['block_500'];
  assert.strictEqual(b1.count, 5);
  assert.strictEqual(b1.students.length, 5);
  assert.strictEqual(b1.block, 'block_500');
  assert.strictEqual(b1.courseId, 'course_1');
});
