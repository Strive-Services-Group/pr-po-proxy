const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const { readBuildInfo, versionHandler } = require('../src/functions/version');

test('version reads one full commit SHA and exposes nothing else', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'prpo-version-'));
  const file = path.join(directory, 'build-info.json');
  const commit = '737793e9f6ef2ce2166ad898ed784602cfae5c57';
  try {
    fs.writeFileSync(file, JSON.stringify({ commit, ignored: 'not returned' }));
    assert.deepEqual(readBuildInfo(file), { commit });
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('version rejects missing, short, or malformed build identities', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'prpo-version-'));
  try {
    for (const value of ['', '737793e', 'not-a-commit', 'g'.repeat(40)]) {
      const file = path.join(directory, `${value.length}.json`);
      fs.writeFileSync(file, JSON.stringify({ commit: value }));
      assert.throws(() => readBuildInfo(file), /full Git commit SHA/);
    }
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('version fails closed without leaking paths or exception detail', async () => {
  const errors = [];
  const missing = path.join(os.tmpdir(), 'prpo-version-file-that-does-not-exist.json');
  const response = await versionHandler({}, { error: (...parts) => errors.push(parts) }, missing);
  assert.equal(response.status, 503);
  assert.deepEqual(response.jsonBody, { error: 'build identity is unavailable' });
  assert.equal(response.headers['Cache-Control'], 'no-store');
  assert.equal(errors.length, 1);
});
