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
  const d=this.sh('Rosters').getDataRange().getValues(); const r={};
  for(let i=1;i<d.length;i++){
    const block=d[i][0],courseId=d[i][1],rawJSON=d[i][2]||'[]',updatedAt=d[i][3];
    let cached=null;
    Object.defineProperty(r,block,{
      get:()=>{if(!cached){const stu=JSON.parse(rawJSON);cached={block,courseId,students:stu,count:stu.length,updatedAt};}return cached;},
      enumerable:true
    });
  }
  return r;
};

test('Benchmark getRosters baseline vs optimized (gas v8 syntax)', () => {
  const start = process.hrtime.bigint();
  const r1 = DB.getRosters();
  const end = process.hrtime.bigint();
  console.log('Baseline:', Number(end - start) / 1000000, 'ms');

  const start2 = process.hrtime.bigint();
  const r2 = DB.getRostersOptimized();
  const end2 = process.hrtime.bigint();
  console.log('Optimized creation:', Number(end2 - start2) / 1000000, 'ms');

  // They should be equivalent
  assert.strictEqual(Object.keys(r1).length, 10000);
  assert.strictEqual(Object.keys(r2).length, 10000);

  const b1 = r1['block_500'];
  const b2 = r2['block_500'];
  assert.deepStrictEqual(b1, b2);
});
