const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert');

// Helper to extract functions from Code.gs
function extractFunctions(code, names) {
  let extracted = '';
  for (const name of names) {
    const startIdx = code.indexOf(`function ${name}(`);
    if (startIdx === -1) throw new Error(`Function ${name} not found`);

    let braceCount = 0;
    let started = false;
    let endIdx = -1;
    for (let i = startIdx; i < code.length; i++) {
      if (code[i] === '{') {
        braceCount++;
        started = true;
      } else if (code[i] === '}') {
        braceCount--;
      }
      if (started && braceCount === 0) {
        endIdx = i;
        break;
      }
    }
    extracted += code.substring(startIdx, endIdx + 1) + '\n\n';
  }
  return extracted;
}

const codeGsPath = path.join(__dirname, '../Code.gs');
const codeGs = fs.readFileSync(codeGsPath, 'utf8');
const functionsToTest = [
  'buildStudentJoinUrl',
  'normalizeStudentCode',
  'normalizeStudentToken',
  'isCanonicalExecUrl'
];

const extractedCode = extractFunctions(codeGs, functionsToTest);

// Evaluate extracted functions in a context
const context = {};
const fn = new Function('exports', extractedCode + '\n' + functionsToTest.map(name => `exports.${name} = ${name};`).join('\n'));
fn(context);

const { buildStudentJoinUrl } = context;

test('buildStudentJoinUrl - standard URL and parameters', () => {
  const baseUrl = 'https://example.com/join';
  const sessionCode = 'UNIT123';
  const studentToken = 'tok_xyz789';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  assert.strictEqual(result, 'https://example.com/join?code=UNIT123&studentToken=tok_xyz789');
});

test('buildStudentJoinUrl - canonical Apps Script execution URL', () => {
  const baseUrl = 'https://script.google.com/macros/s/AKfycbz_123/exec';
  const sessionCode = 'UNIT123';
  const studentToken = 'tok_xyz789';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  assert.strictEqual(result, 'https://script.google.com/macros/s/AKfycbz_123/exec?page=student&code=UNIT123&studentToken=tok_xyz789');
});

test('buildStudentJoinUrl - legacy Google Workspace URL fix', () => {
  const baseUrl = 'https://script.google.com/a/my-school.edu/macros/s/AKfycbz_123/exec';
  const sessionCode = 'UNIT123';
  const studentToken = 'tok_xyz789';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  // Should fix the URL format to canonical macro format
  assert.strictEqual(result, 'https://script.google.com/a/macros/my-school.edu/s/AKfycbz_123/exec?page=student&code=UNIT123&studentToken=tok_xyz789');
});

test('buildStudentJoinUrl - URL with existing query parameters', () => {
  const baseUrl = 'https://example.com/join?ref=email';
  const sessionCode = 'UNIT123';
  const studentToken = 'tok_xyz789';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  assert.strictEqual(result, 'https://example.com/join?ref=email&code=UNIT123&studentToken=tok_xyz789');
});

test('buildStudentJoinUrl - missing session code and student token', () => {
  const baseUrl = 'https://example.com/join';
  const result = buildStudentJoinUrl(baseUrl, '', '');
  assert.strictEqual(result, 'https://example.com/join');
});

test('buildStudentJoinUrl - missing baseUrl', () => {
  const result = buildStudentJoinUrl('', 'UNIT123', 'tok_xyz789');
  assert.strictEqual(result, '');
});

test('buildStudentJoinUrl - parameter encoding', () => {
  const baseUrl = 'https://example.com/join';
  const sessionCode = 'unit 123!'; // normalizeStudentCode will strip spaces and !
  const studentToken = 'token&with=specialChars';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  // normalizeStudentCode('unit 123!') -> 'UNIT123'
  // normalizeStudentToken('token&with=specialChars') -> 'token&with=specialChars' (no spaces to remove)
  assert.strictEqual(result, 'https://example.com/join?code=UNIT123&studentToken=token%26with%3DspecialChars');
});

test('buildStudentJoinUrl - whitespace handling', () => {
  const baseUrl = '  https://example.com/join  ';
  const sessionCode = '  UNIT123  ';
  const studentToken = '  tok_xyz789  ';
  const result = buildStudentJoinUrl(baseUrl, sessionCode, studentToken);
  assert.strictEqual(result, 'https://example.com/join?code=UNIT123&studentToken=tok_xyz789');
});
