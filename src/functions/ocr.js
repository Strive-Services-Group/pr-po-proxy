const { app } = require('@azure/functions');

// ---- OCR proxy: reads Emirates IDs + work permits via Azure AI Document
//      Intelligence (prebuilt-read). Keeps the Azure key server-side.
//      POST the raw image/PDF bytes as the body -> { text, lines }.
//      App settings needed: AZURE_DI_ENDPOINT, AZURE_DI_KEY
//      (optional AZURE_DI_API_VERSION, default 2024-11-30).
//      Add the Visitor app's origin to ALLOWED_ORIGIN (comma-separated). ----

const ENDPOINT = () => (process.env.AZURE_DI_ENDPOINT || '').replace(/\/+$/, '');
const KEY      = () => process.env.AZURE_DI_KEY || '';
const API_VER  = () => process.env.AZURE_DI_API_VERSION || '2024-11-30';

function cors(request) {
  const allowed = (process.env.ALLOWED_ORIGIN || '*').split(',').map(s => s.trim()).filter(Boolean);
  const reqOrigin = request && request.headers && request.headers.get ? (request.headers.get('origin') || '') : '';
  let origin;
  if (allowed.includes('*')) origin = '*';
  else if (allowed.includes(reqOrigin)) origin = reqOrigin;
  else origin = allowed[0] || '*';
  return {
    'Access-Control-Allow-Origin': origin,
    'Vary': 'Origin',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Cache-Control': 'no-store'
  };
}

async function readOcr(request, context) {
  const headers = { 'Content-Type': 'application/json', ...cors(request) };
  if (request.method === 'OPTIONS') return { status: 204, headers };

  try {
    if (!ENDPOINT() || !KEY()) return { status: 500, headers, jsonBody: { error: 'OCR not configured (AZURE_DI_ENDPOINT / AZURE_DI_KEY).' } };

    const bytes = Buffer.from(await request.arrayBuffer());
    if (!bytes.length) return { status: 400, headers, jsonBody: { error: 'No file bytes received.' } };

    const analyzeUrl = ENDPOINT() + '/documentintelligence/documentModels/prebuilt-read:analyze?api-version=' + API_VER();
    const start = await fetch(analyzeUrl, {
      method: 'POST',
      headers: {
        'Ocp-Apim-Subscription-Key': KEY(),
        'Content-Type': request.headers.get('content-type') || 'application/octet-stream'
      },
      body: bytes
    });
    if (!(start.status === 202 || start.ok)) {
      const t = await start.text().catch(() => '');
      return { status: 502, headers, jsonBody: { error: 'Azure analyze failed (' + start.status + ')', detail: t.slice(0, 400) } };
    }
    const op = start.headers.get('operation-location');
    if (!op) return { status: 502, headers, jsonBody: { error: 'No operation-location returned by Azure.' } };

    let result = null;
    for (let i = 0; i < 25; i++) {
      await new Promise(r => setTimeout(r, 700));
      const poll = await fetch(op, { headers: { 'Ocp-Apim-Subscription-Key': KEY() } });
      const j = await poll.json();
      if (j.status === 'succeeded') { result = j; break; }
      if (j.status === 'failed') return { status: 502, headers, jsonBody: { error: 'Azure read failed.' } };
    }
    if (!result) return { status: 504, headers, jsonBody: { error: 'Azure read timed out.' } };

    const ar = result.analyzeResult || {};
    const text = ar.content || '';
    let lines = [];
    (ar.pages || []).forEach(pg => (pg.lines || []).forEach(l => { if (l.content) lines.push(l.content); }));
    if (!lines.length && text) lines = text.split('\n');
    return { status: 200, headers, jsonBody: { text, lines } };
  } catch (e) {
    context.error(e);
    return { status: 500, headers, jsonBody: { error: String(e && e.message || e) } };
  }
}

app.http('ocr', { methods: ['POST', 'OPTIONS'], authLevel: 'function', route: 'ocr', handler: readOcr });
