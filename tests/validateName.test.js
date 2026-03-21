const test = require('node:test');
const assert = require('node:assert');

// ── Inline implementation matching Data.gs DB.validateName_() ─────────────
// We mirror the production logic here so the test acts as a living spec.
function validateName(name) {
  const s = String(name || '').trim();
  if (!s) return 'Name cannot be empty.';
  if (s.length > 200) return 'Name too long (max 200 characters).';
  return null;
}

// ── Tests ──────────────────────────────────────────────────────────────────

test('validateName rejects an empty string', () => {
  assert.strictEqual(validateName(''), 'Name cannot be empty.');
});

test('validateName rejects a whitespace-only string', () => {
  assert.strictEqual(validateName('   '), 'Name cannot be empty.');
  assert.strictEqual(validateName('\t\n'), 'Name cannot be empty.');
});

test('validateName rejects null and undefined', () => {
  assert.strictEqual(validateName(null), 'Name cannot be empty.');
  assert.strictEqual(validateName(undefined), 'Name cannot be empty.');
});

test('validateName rejects a name longer than 200 characters', () => {
  const long = 'A'.repeat(201);
  assert.strictEqual(validateName(long), 'Name too long (max 200 characters).');
});

test('validateName accepts exactly 200 characters', () => {
  const exact = 'A'.repeat(200);
  assert.strictEqual(validateName(exact), null);
});

test('validateName accepts a typical short name', () => {
  assert.strictEqual(validateName('Biology 101'), null);
  assert.strictEqual(validateName('AP Chemistry'), null);
  assert.strictEqual(validateName('Block 3 Roster'), null);
});

test('validateName trims leading/trailing whitespace before checking length', () => {
  // 198 'A's padded with 1 space on each side = 200 chars total → trim → 198 chars → valid.
  const padded = ' ' + 'A'.repeat(198) + ' ';
  assert.strictEqual(validateName(padded), null);
});

test('validateName rejects a name that is 201 characters after trimming', () => {
  const tooLong = '  ' + 'A'.repeat(201) + '  ';
  assert.strictEqual(validateName(tooLong), 'Name too long (max 200 characters).');
});
