const assert = require('assert');
const fs = require('fs');
const { performance } = require('perf_hooks');

const gsContent = fs.readFileSync('./Data.gs', 'utf-8');

class Range {
  constructor(values, startRow = 1) {
    this.values = values;
    this.startRow = startRow;
    this.apiCalls = 0;
  }
  getValues() {
    this.apiCalls++;
    return this.values;
  }
  getValue() {
    this.apiCalls++;
    return this.values[0][0];
  }
  setValue(val) {
    this.apiCalls++;
    if (this.values && this.values.length > 0 && this.values[0] && this.values[0].length > 0) {
      this.values[0][0] = val;
    }
  }
  deleteRow() {
     this.apiCalls++;
  }
}

class TextFinder {
  constructor(sheet, text) {
    this.sheet = sheet;
    this.text = text;
  }
  matchEntireCell(match) { return this; }
  findAll() {
    this.sheet.apiCalls++; // API call for TextFinder
    const matches = [];
    for (let r=0; r<this.sheet.data.length; r++) {
       for (let c=0; c<this.sheet.data[r].length; c++) {
          if (String(this.sheet.data[r][c]) === this.text) {
             matches.push({
                getRow: () => r + 1,
                getColumn: () => c + 1
             });
          }
       }
    }
    return matches;
  }
}

class Sheet {
  constructor(name, data) {
    this.name = name;
    this.data = data;
    this.apiCalls = 0;
    this.rowsDeleted = 0;
  }
  getDataRange() {
    this.apiCalls++;
    return new Range(this.data);
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    this.apiCalls++;
    return new Range([[this.data[row-1][col-1]]], row);
  }
  deleteRow(row) {
    this.rowsDeleted++;
    this.apiCalls++;
  }
  createTextFinder(text) {
    return new TextFinder(this, text);
  }
}

global.SpreadsheetApp = {
    openById: () => ({ getSheetByName: (name) => global.SpreadsheetApp._sheets[name] }),
    _sheets: {}
};
global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => 'fake_id' }) };
global.Logger = { log: console.log };

let DB;
eval(gsContent.replace('const DB = {', 'DB = {'));

function runBenchmark(numViolations) {
  const sessionId = 'session-123';
  const timestamp = '1234567890';
  const dataJSON = JSON.stringify({ violations: [{ timestamp: timestamp }]});

  global.SpreadsheetApp._sheets['Archive'] = new Sheet('Archive', [
    ['SessionID', 'Code', 'SetName', 'Block', 'StartedAt', 'EndedAt', 'StudentCount', 'AvgPct', 'DataJSON'],
    [sessionId, 'code1', 'Set1', 'Block1', 'T1', 'T2', 10, 80, dataJSON]
  ]);

  const violData = [['SessionID', 'StudentID', 'StudentName', 'Type', 'Timestamp', 'Resolved']];
  for (let i = 0; i < numViolations; i++) {
    if (i === numViolations - 1) {
      violData.push([sessionId, 'student-1', 'John Doe', 'Tab Switch', timestamp, false]);
    } else {
      violData.push(['other-session', 'other-student', 'Other', 'Type', '0', false]);
    }
  }

  const violSheet = new Sheet('Violations', violData);
  global.SpreadsheetApp._sheets['Violations'] = violSheet;

  const start = performance.now();
  const res = DB.dismissViolation(sessionId, timestamp);
  const end = performance.now();

  return {
    success: res,
    time: end - start,
    apiCalls: global.SpreadsheetApp._sheets['Violations'].apiCalls + global.SpreadsheetApp._sheets['Archive'].apiCalls,
    rowsDeleted: global.SpreadsheetApp._sheets['Violations'].rowsDeleted
  };
}

console.log('Optimized Result (100,000 rows):', runBenchmark(100000));
