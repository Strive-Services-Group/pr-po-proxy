const fs = require('fs');
const path = require('path');
const { app } = require('@azure/functions');

const BUILD_INFO_PATH = path.join(__dirname, '..', '..', 'build-info.json');
const COMMIT_PATTERN = /^[0-9a-f]{40}$/;

function readBuildInfo(filePath = BUILD_INFO_PATH) {
  const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  const commit = String(parsed && parsed.commit || '').trim().toLowerCase();
  if (!COMMIT_PATTERN.test(commit)) {
    throw new Error('build-info.json does not contain a full Git commit SHA');
  }
  return { commit };
}

async function versionHandler(_request, context, filePath = BUILD_INFO_PATH) {
  try {
    return {
      status: 200,
      headers: { 'Cache-Control': 'no-store' },
      jsonBody: readBuildInfo(filePath),
    };
  } catch (error) {
    context.error('version endpoint failed:', error);
    return {
      status: 503,
      headers: { 'Cache-Control': 'no-store' },
      jsonBody: { error: 'build identity is unavailable' },
    };
  }
}

app.http('version', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'version',
  handler: versionHandler,
});

module.exports = { readBuildInfo, versionHandler, BUILD_INFO_PATH };
