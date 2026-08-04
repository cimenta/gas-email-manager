'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { APP_VERSION } = require('../src/02-main.js');

// --- APP_VERSION -------------------------------------------------------

test('APP_VERSION: is a non-empty semver-ish string (three dot-separated numeric segments)', () => {
  assert.equal(typeof APP_VERSION, 'string');
  assert.notEqual(APP_VERSION, '');
  assert.match(APP_VERSION, /^\d+\.\d+\.\d+$/);
});

test('APP_VERSION: matches the line-1 "// Version:" comment in src/02-main.js (dual-representation guard)', () => {
  const raw = fs.readFileSync(path.join(__dirname, '..', 'src', '02-main.js'), 'utf8');
  const firstLine = raw.split('\n')[0];
  const tokens = firstLine.split(/\s+/);
  assert.equal(tokens[2], APP_VERSION);
});
