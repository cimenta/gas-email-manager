'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { composeFailureBody } = require('../src/02-main.js');

test('composeFailureBody includes action name, thread subject, and error text as plain text', () => {
  const body = composeFailureBody(
    'ics-calendar-import',
    'Team offsite invite',
    'Error: Missing DTSTART in VEVENT 2'
  );

  assert.equal(typeof body, 'string');
  assert.ok(body.includes('ics-calendar-import'));
  assert.ok(body.includes('Team offsite invite'));
  assert.ok(body.includes('Error: Missing DTSTART in VEVENT 2'));
  assert.equal(/[<>]/.test(body), false, 'body must not contain HTML angle brackets');
});
