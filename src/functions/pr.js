'use strict';

const { app } = require('@azure/functions');
const { jwtVerify, createRemoteJWKSet } = require('jose');
const { getDataset } = require('../shared/prpoDataset');

const TENANT = process.env.TENANT_ID;
const AUDIENCE = process.env.DASHBOARD_CLIENT_ID;
const JWKS = TENANT ? createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${TENANT}/discovery/v2.0/keys`)) : null;

async function requireUser(request) {
  const value = (request.headers.get && request.headers.get('authorization')) || '';
  const match = value.match(/^Bearer\s+(.+)$/i);
  if (!match) throw new Error('unauthorized: missing token');
  await jwtVerify(match[1], JWKS, {
    issuer: `https://login.microsoftonline.com/${TENANT}/v2.0`,
    audience: AUDIENCE
  });
}

function cors() {
  return {
    'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Cache-Control': 'no-store'
  };
}

async function serve(kind, request, context) {
  const headers = { 'Content-Type': 'application/json', ...cors() };
  if (request.method === 'OPTIONS') return { status: 204, headers };
  try {
    const authz = (request.headers.get && request.headers.get('authorization')) || '';
    if (authz) await requireUser(request);
  } catch (_) {
    return { status: 401, headers, jsonBody: { error: 'unauthorized' } };
  }
  try {
    const url = new URL(request.url);
    const dataset = await getDataset({ force: url.searchParams.get('refresh') === '1' });
    if (kind === 'dataset') return { status: 200, headers, jsonBody: dataset };
    return {
      status: 200,
      headers,
      jsonBody: {
        type: kind,
        revision: dataset.revision,
        generatedAt: dataset.generatedAt,
        sourceState: dataset.sourceState,
        cached: Boolean(dataset.cached),
        refreshError: dataset.refreshError || null,
        ...dataset[kind]
      }
    };
  } catch (error) {
    context.error(error);
    return { status: 503, headers, jsonBody: { error: String(error && error.message || error), sourceState: 'FAILED' } };
  }
}

app.http('dataset', { methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'dataset', handler: (req, ctx) => serve('dataset', req, ctx) });
app.http('pr', { methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'pr', handler: (req, ctx) => serve('pr', req, ctx) });
app.http('po', { methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'po', handler: (req, ctx) => serve('po', req, ctx) });

module.exports = { serve };
