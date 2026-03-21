const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const test = require('node:test');
const assert = require('node:assert');

// ── GAS API mocks ──────────────────────────────────────────────────────────
const SECRET_STORE = { STUDENT_LINK_SECRET: 'test-secret-for-unit-tests' };

const mockUtilities = {
  computeHmacSha256Signature(data, key) {
    const keyBuf = typeof key === 'string' ? Buffer.from(key, 'utf8') : Buffer.from(key);
    const dataBuf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
    const hmac = crypto.createHmac('sha256', keyBuf);
    hmac.update(dataBuf);
    return Array.from(hmac.digest());
  },
  base64EncodeWebSafe(data) {
    const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
    return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  },
  base64DecodeWebSafe(data) {
    const normalized = String(data || '').replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized + '==='.slice((normalized.length + 3) % 4);
    return Array.from(Buffer.from(padded, 'base64'));
  },
  newBlob(bytes) {
    return { getDataAsString() { return Buffer.from(bytes).toString('utf8'); } };
  },
  Charset: { UTF_8: 'UTF-8' },
  getUuid() {
    return crypto.randomUUID();
  },
};

const mockPropertiesService = {
  getScriptProperties() {
    return {
      getProperty(key) { return SECRET_STORE[key] || null; },
      setProperty(key, val) { SECRET_STORE[key] = val; },
      deleteProperty(key) { delete SECRET_STORE[key]; },
    };
  },
};

const mockLockService = {
  getScriptLock() {
    return { waitLock() {}, releaseLock() {} };
  },
};

// ── Function extractor (same pattern as buildStudentJoinUrl.test.js) ────────
function extractFunctions(code, names) {
  let extracted = '';
  for (const name of names) {
    const startIdx = code.indexOf(`function ${name}(`);
    if (startIdx === -1) throw new Error(`Function ${name} not found in Code.gs`);
    let braceCount = 0, started = false, endIdx = -1;
    for (let i = startIdx; i < code.length; i++) {
      if (code[i] === '{') { braceCount++; started = true; }
      else if (code[i] === '}') { braceCount--; }
      if (started && braceCount === 0) { endIdx = i; break; }
    }
    extracted += code.substring(startIdx, endIdx + 1) + '\n\n';
  }
  return extracted;
}

const codeGs = fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8');

const fnNames = [
  'normalizeStudentCode',
  'normalizeStudentToken',
  'normalizeStudentEmail_',
  'normalizeStudentIdentityName_',
  'getStudentNameParts_',
  'base64UrlEncode_',
  'base64UrlDecodeToString_',
  'constantTimeEquals_',
  'getStudentLinkSecret_',
  'rotateStudentLinkSecret',
  'signStudentAccessPayload_',
  'createStudentAccessToken',
  'verifyStudentAccessToken',
];

const extracted = extractFunctions(codeGs, fnNames);

const ctx = {};
const fn = new Function(
  'exports', 'Utilities', 'PropertiesService', 'LockService', 'Logger', 'DB',
  extracted + '\n' + fnNames.map(n => `exports.${n} = ${n};`).join('\n')
);
fn(
  ctx,
  mockUtilities,
  mockPropertiesService,
  mockLockService,
  { log() {} },
  { logAuditEvent() {} }  // stub for rotateStudentLinkSecret calling DB.logAuditEvent
);

const {
  createStudentAccessToken,
  verifyStudentAccessToken,
  base64UrlEncode_,
  signStudentAccessPayload_,
} = ctx;

// ── Helper: craft a token with an arbitrary payload ───────────────────────
function craftToken(payload) {
  const seg = base64UrlEncode_(JSON.stringify(payload));
  return seg + '.' + signStudentAccessPayload_(seg);
}

const VALID_SESSION = { sessionId: 'sess_test01', code: 'PHOTON42', block: '2' };
const VALID_STUDENT = { email: 'alice@school.edu', firstName: 'Alice', lastName: 'Smith' };

// ── Tests ──────────────────────────────────────────────────────────────────

test('createStudentAccessToken includes expiresAt ~7 days in the future', () => {
  const token = createStudentAccessToken(VALID_SESSION, VALID_STUDENT);
  assert.ok(token, 'token should be non-empty');
  const payloadJson = JSON.parse(Buffer.from(
    token.split('.')[0].replace(/-/g, '+').replace(/_/g, '/') + '===',
    'base64'
  ).toString('utf8'));
  assert.ok(payloadJson.expiresAt, 'token payload must contain expiresAt');
  const diff = new Date(payloadJson.expiresAt) - new Date(payloadJson.issuedAt);
  const sevenDaysMs = 7 * 24 * 60 * 60 * 1000;
  assert.ok(Math.abs(diff - sevenDaysMs) < 5000, 'expiresAt should be ~7 days after issuedAt');
});

test('verifyStudentAccessToken accepts a freshly created token', () => {
  const token = createStudentAccessToken(VALID_SESSION, VALID_STUDENT);
  const payload = verifyStudentAccessToken(token);
  assert.ok(payload, 'should return a valid payload');
  assert.strictEqual(payload.sid, 'sess_test01');
  assert.strictEqual(payload.email, 'alice@school.edu');
});

test('verifyStudentAccessToken rejects a token with a past expiresAt', () => {
  const past = new Date(Date.now() - 1000).toISOString(); // 1 second ago
  const token = craftToken({
    v: 1, sid: 'sess_test01', code: 'PHOTON42', block: '2',
    email: 'alice@school.edu', name: 'Alice Smith',
    firstName: 'Alice', lastName: 'Smith', normalizedName: 'alice smith',
    issuedAt: new Date(Date.now() - 8 * 24 * 60 * 60 * 1000).toISOString(),
    expiresAt: past,
  });
  const result = verifyStudentAccessToken(token);
  assert.strictEqual(result, null, 'expired token must be rejected');
});

test('verifyStudentAccessToken accepts a token with a future expiresAt', () => {
  const future = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
  const token = craftToken({
    v: 1, sid: 'sess_test01', code: 'PHOTON42', block: '2',
    email: 'alice@school.edu', name: 'Alice Smith',
    firstName: 'Alice', lastName: 'Smith', normalizedName: 'alice smith',
    issuedAt: new Date().toISOString(),
    expiresAt: future,
  });
  const result = verifyStudentAccessToken(token);
  assert.ok(result, 'non-expired token must be accepted');
  assert.strictEqual(result.sid, 'sess_test01');
});

test('verifyStudentAccessToken accepts a legacy token with no expiresAt (backward-compat)', () => {
  // Tokens issued before this feature did not have expiresAt — they should still work.
  const token = craftToken({
    v: 1, sid: 'sess_legacy', code: 'PHOTON42', block: '1',
    email: 'bob@school.edu', name: 'Bob Jones',
    firstName: 'Bob', lastName: 'Jones', normalizedName: 'bob jones',
    issuedAt: new Date().toISOString(),
    // no expiresAt
  });
  const result = verifyStudentAccessToken(token);
  assert.ok(result, 'legacy token without expiresAt should still be accepted');
  assert.strictEqual(result.sid, 'sess_legacy');
});

test('verifyStudentAccessToken rejects a tampered token', () => {
  const token = createStudentAccessToken(VALID_SESSION, VALID_STUDENT);
  const tampered = token.slice(0, -4) + 'XXXX'; // corrupt last 4 chars of signature
  const result = verifyStudentAccessToken(tampered);
  assert.strictEqual(result, null, 'tampered token must be rejected');
});

test('verifyStudentAccessToken rejects an empty token', () => {
  assert.strictEqual(verifyStudentAccessToken(''), null);
  assert.strictEqual(verifyStudentAccessToken(null), null);
  assert.strictEqual(verifyStudentAccessToken('notavalidtoken'), null);
});

test('rotateStudentLinkSecret invalidates old tokens', () => {
  // Capture a valid token with the current secret.
  const token = createStudentAccessToken(VALID_SESSION, VALID_STUDENT);
  assert.ok(verifyStudentAccessToken(token), 'token valid before rotation');

  // Rotate the secret.
  ctx.rotateStudentLinkSecret();

  // The old token must now be invalid because the HMAC secret changed.
  const result = verifyStudentAccessToken(token);
  assert.strictEqual(result, null, 'old token must be invalid after secret rotation');
});
