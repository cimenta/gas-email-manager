'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { extractEmailAddress, isAllowedSender, isExcludedSender } = require('../src/05-action-ics-import.js');

// --- extractEmailAddress ---------------------------------------------------

test('extractEmailAddress: angle-bracket header returns the bare lowercased address', () => {
  assert.equal(extractEmailAddress('Jana Nováková <jana.novakova@example.com>'), 'jana.novakova@example.com');
});

test('extractEmailAddress: bare address (no display name) returns itself', () => {
  assert.equal(extractEmailAddress('jana.novakova@example.com'), 'jana.novakova@example.com');
});

test('extractEmailAddress: mixed-case input is lowercased', () => {
  assert.equal(extractEmailAddress('Foo <Foo.Bar@Example.COM>'), 'foo.bar@example.com');
});

test('extractEmailAddress: surrounding whitespace trimmed for bare input', () => {
  assert.equal(extractEmailAddress('  jana@example.com  '), 'jana@example.com');
});

test('extractEmailAddress: whitespace inside brackets trimmed', () => {
  assert.equal(extractEmailAddress('Name < jana@example.com >'), 'jana@example.com');
});

test('extractEmailAddress: null returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(extractEmailAddress(null), '');
  });
});

test('extractEmailAddress: undefined returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(extractEmailAddress(undefined), '');
  });
});

test('extractEmailAddress: empty string returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(extractEmailAddress(''), '');
  });
});

// --- isAllowedSender --------------------------------------------------------

test('isAllowedSender: empty list allows any sender (import-all default)', () => {
  assert.equal(isAllowedSender('anyone@x.com', []), true);
});

test('isAllowedSender: null list allows any sender (import-all default)', () => {
  assert.equal(isAllowedSender('anyone@x.com', null), true);
});

test('isAllowedSender: undefined list allows any sender (import-all default)', () => {
  assert.equal(isAllowedSender('anyone@x.com', undefined), true);
});

test('isAllowedSender: non-empty list with match, bare list entry', () => {
  assert.equal(isAllowedSender('Jana <jana@example.com>', ['jana@example.com']), true);
});

test('isAllowedSender: non-empty list where the list entry is a full "Name <email>" string', () => {
  assert.equal(isAllowedSender('jana@example.com', ['Jana Nováková <jana@example.com>']), true);
});

test('isAllowedSender: non-empty list, no match', () => {
  assert.equal(isAllowedSender('bob@other.com', ['jana@example.com']), false);
});

test('isAllowedSender: case-insensitive match', () => {
  assert.equal(isAllowedSender('JANA@EXAMPLE.COM', ['jana@example.com']), true);
});

// --- isExcludedSender (quick-260803-us3, D-03) -------------------------------

test('isExcludedSender: empty list excludes nobody', () => {
  assert.equal(isExcludedSender('anyone@x.com', []), false);
});

test('isExcludedSender: null list excludes nobody', () => {
  assert.equal(isExcludedSender('anyone@x.com', null), false);
});

test('isExcludedSender: undefined list excludes nobody', () => {
  assert.equal(isExcludedSender('anyone@x.com', undefined), false);
});

test('isExcludedSender: a bare list entry matches a "Name <addr>" header', () => {
  assert.equal(isExcludedSender('RegioJet <jizdenky@regiojet.cz>', ['jizdenky@regiojet.cz']), true);
});

test('isExcludedSender: a "Name <addr>" list entry matches a bare header', () => {
  assert.equal(isExcludedSender('jizdenky@regiojet.cz', ['RegioJet <jizdenky@regiojet.cz>']), true);
});

test('isExcludedSender: matching is case-insensitive', () => {
  assert.equal(isExcludedSender('JIZDENKY@REGIOJET.CZ', ['jizdenky@regiojet.cz']), true);
});

test('isExcludedSender: a non-listed sender returns false', () => {
  assert.equal(isExcludedSender('someone-else@example.com', ['jizdenky@regiojet.cz']), false);
});

test('isExcludedSender: isAllowedSender is unchanged when a sender appears only in the exclude list (the two gates are independent)', () => {
  assert.equal(isAllowedSender('jizdenky@regiojet.cz', []), true);
  assert.equal(isExcludedSender('jizdenky@regiojet.cz', ['jizdenky@regiojet.cz']), true);
});
