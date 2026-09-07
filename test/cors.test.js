'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { cors } = require('../src/functions/pr');

const request = origin => ({ headers: { get: name => name === 'origin' ? origin : '' } });

test('dashboard and retained configured origin receive exact CORS origin', () => {
  process.env.ALLOWED_ORIGIN = 'https://chandansah605.github.io';
  assert.equal(cors(request('https://strive-services-group.github.io'))['Access-Control-Allow-Origin'], 'https://strive-services-group.github.io');
  assert.equal(cors(request('https://chandansah605.github.io'))['Access-Control-Allow-Origin'], 'https://chandansah605.github.io');
  assert.equal(cors(request('https://example.invalid'))['Access-Control-Allow-Origin'], 'https://strive-services-group.github.io');
});
