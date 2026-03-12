const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert');

// A simple mock for SpreadsheetApp / Logger / etc.
global.Logger = { log: console.log };

const codeGsPath = path.join(__dirname, '../Data.gs');
let codeGs = fs.readFileSync(codeGsPath, 'utf8');

codeGs = codeGs.replace('const DB = {', 'global.DB = {');
eval(codeGs);

// Mock DB dependencies
DB.sh = function(sheetName) {
  return {
    getDataRange: function() {
      return {
        getValues: function() {
          const rows = [['Block', 'CourseID', 'StudentsJSON', 'UpdatedAt']];
          // generate 1000 rows
          for (let i = 1; i <= 1000; i++) {
            rows.push([
              'block_' + i,
              'course_1',
              JSON.stringify([{id: i, name: 'Student ' + i}]),
              '2023-01-01T00:00:00Z'
            ]);
          }
          return rows;
        }
      }
    }
  }
};

test('Benchmark getRosters', () => {
  console.time('getRosters');
  const r = DB.getRosters();
  console.timeEnd('getRosters');

  assert.strictEqual(Object.keys(r).length, 1000);
});
