const { app } = require('@azure/functions');

/**
 * Yahoo Finance CORS proxy for the Stock Market ("Pulse") dashboard.
 * ------------------------------------------------------------------
 * GET /api/yahoo/{*path}  →  https://query1.finance.yahoo.com/{path}?{query}
 *
 * Only whitelisted Yahoo paths are forwarded. Handles Yahoo's cookie+crumb
 * handshake for the screener endpoint (cached ~30 min per warm instance,
 * auto-retried once on 401). Anonymous + permissive CORS by design: it only
 * relays public market data — no secrets, no Dataverse/F&O access.
 * Override allowed origins with env YAHOO_ALLOWED_ORIGINS (comma-separated).
 */

const YAHOO = 'https://query1.finance.yahoo.com';
const ALLOWED_PATHS = [
  '/v1/finance/screener/predefined/saved',
  '/v7/finance/spark',
  '/v1/finance/search',
  '/v8/finance/chart/',
];
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36';

let crumbCache = { cookie: null, crumb: null, at: 0 };
const CRUMB_TTL = 30 * 60 * 1000;

async function getCrumb() {
  if (crumbCache.crumb && Date.now() - crumbCache.at < CRUMB_TTL) return crumbCache;
  const r1 = await fetch('https://fc.yahoo.com', { headers: { 'User-Agent': UA }, redirect: 'manual' });
  const cookie = (r1.headers.get('set-cookie') || '').split(';')[0];
  const r2 = await fetch(YAHOO + '/v1/test/getcrumb', { headers: { 'User-Agent': UA, Cookie: cookie } });
  const crumb = (await r2.text()).trim();
  if (crumb && !crumb.includes('<')) crumbCache = { cookie, crumb, at: Date.now() };
  return crumbCache;
}

function corsHeaders(request) {
  const allowed = (process.env.YAHOO_ALLOWED_ORIGINS || '*').split(',').map(s => s.trim());
  const origin = (request.headers.get && request.headers.get('origin')) || '';
  const ok = allowed.includes('*') ? '*' : (allowed.includes(origin) ? origin : '');
  return {
    'Access-Control-Allow-Origin': ok || 'null',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Accept',
    'Vary': 'Origin',
  };
}

app.http('yahooProxy', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'yahoo/{*path}',
  handler: async (request, context) => {
    const headers = corsHeaders(request);
    if (request.method === 'OPTIONS') return { status: 204, headers };

    try {
      const path = '/' + (request.params.path || '');
      if (!ALLOWED_PATHS.some(p => path.startsWith(p))) {
        return { status: 403, headers, body: 'Path not allowed' };
      }

      const search = new URL(request.url).search;
      const target = new URL(YAHOO + path + search);
      target.searchParams.delete('code'); // strip any function key from the forwarded query
      const upstream = { 'User-Agent': UA, 'Accept': 'application/json' };

      if (path.startsWith('/v1/finance/screener')) {
        const { cookie, crumb } = await getCrumb();
        if (crumb) { target.searchParams.set('crumb', crumb); upstream['Cookie'] = cookie; }
      }

      let r = await fetch(target, { headers: upstream });
      if (r.status === 401 && path.startsWith('/v1/finance/screener')) {
        crumbCache = { cookie: null, crumb: null, at: 0 };
        const { cookie, crumb } = await getCrumb();
        if (crumb) { target.searchParams.set('crumb', crumb); upstream['Cookie'] = cookie; }
        r = await fetch(target, { headers: upstream });
      }

      return {
        status: r.status,
        headers: Object.assign({}, headers, {
          'Content-Type': 'application/json; charset=utf-8',
          'Cache-Control': 'public, max-age=30',
        }),
        body: await r.text(),
      };
    } catch (e) {
      context.error('yahooProxy failed:', e);
      return { status: 502, headers, jsonBody: { error: String(e) } };
    }
  },
});
