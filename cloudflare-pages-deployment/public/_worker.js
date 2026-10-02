/*!
 * Genie Connect — Cloudflare Pages "Advanced Mode" worker.
 * -----------------------------------------------------------------------
 * This file is NOT a separate thing you deploy. Cloudflare Pages notices
 * a file named exactly `_worker.js` sitting next to index.html and runs
 * it for every request to this project — static files included. There is
 * no server process, nothing to keep running, nothing to patch: you
 * connect this repo in the Cloudflare dashboard once, and every git push
 * rebuilds and redeploys this automatically.
 *
 * GENIE_API_KEY and IAM_API_TOKEN are never in this file — they're set
 * as encrypted environment variables in the Pages dashboard (Settings →
 * Environment variables → "Encrypt"), or via `wrangler pages secret put`
 * if you prefer the CLI. Either way, nothing secret is ever in git.
 *
 * Because the site and this API are the exact same Pages project, every
 * fetch('/api/...') call from index.html is same-origin — no CORS setup
 * needed, unlike the split GitHub-Pages-site + separate-Worker approach.
 *
 * Routes (identical to server.js / worker.js):
 *   POST /api/genie/conversations
 *   POST /api/genie/conversations/:conversationId/messages   (streamed)
 *   GET  /api/iam/users?email=...&dataCenter=...
 *   GET  /api/health
 *   anything else → served as a normal static file from this folder
 * -----------------------------------------------------------------------
 */

function json(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

function genieBase(genieId) {
  return `https://genie-api.workato.com/api/v1/genies/${genieId}/chat`;
}

async function handleApi(request, env, url) {
  const { pathname } = url;

  // ---------------------------------------------------------------
  // POST /api/genie/conversations
  // ---------------------------------------------------------------
  if (pathname === '/api/genie/conversations' && request.method === 'POST') {
    const body = await request.json().catch(() => ({}));
    const genieId = body.genieId || env.GENIE_ID;
    const idpUserId = body.idpUserId;

    if (!genieId) return json({ error: 'Missing genieId' }, 400);
    if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400);
    if (!env.GENIE_API_KEY) return json({ error: 'Missing GENIE_API_KEY env var' }, 500);

    const upstream = await fetch(`${genieBase(genieId)}/conversations`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.GENIE_API_KEY}`,
        'X-IDP-User-ID': idpUserId,
        'Content-Type': 'application/json',
      },
    });
    const text = await upstream.text();
    return new Response(text, {
      status: upstream.status,
      headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json' },
    });
  }

  // ---------------------------------------------------------------
  // POST /api/genie/conversations/:conversationId/messages  (streamed)
  // ---------------------------------------------------------------
  const msgMatch = pathname.match(/^\/api\/genie\/conversations\/([^/]+)\/messages$/);
  if (msgMatch && request.method === 'POST') {
    const conversationId = msgMatch[1];
    const body = await request.json().catch(() => ({}));
    const genieId = body.genieId || env.GENIE_ID;
    const idpUserId = body.idpUserId;
    const message = body.message;

    if (!genieId) return json({ error: 'Missing genieId' }, 400);
    if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400);
    if (!message) return json({ error: 'Missing message' }, 400);
    if (!env.GENIE_API_KEY) return json({ error: 'Missing GENIE_API_KEY env var' }, 500);

    const upstream = await fetch(`${genieBase(genieId)}/conversations/${conversationId}/messages`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.GENIE_API_KEY}`,
        'X-IDP-User-ID': idpUserId,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message, stream: true }),
    });

    if (!upstream.ok || !upstream.body) {
      const errText = await upstream.text();
      return new Response(errText, { status: upstream.status });
    }

    // Pass the upstream stream straight through to the browser.
    return new Response(upstream.body, {
      status: 200,
      headers: {
        'Content-Type': upstream.headers.get('content-type') || 'text/event-stream',
        'Cache-Control': 'no-cache',
        Connection: 'keep-alive',
      },
    });
  }

  // ---------------------------------------------------------------
  // GET /api/iam/users?email=...&dataCenter=...
  // ---------------------------------------------------------------
  if (pathname === '/api/iam/users' && request.method === 'GET') {
    const email = url.searchParams.get('email');
    const dataCenter = (url.searchParams.get('dataCenter') || env.DATA_CENTER || '')
      .replace(/^https?:\/\//, '')
      .replace(/\/$/, '');

    if (!email) return json({ error: 'Missing email query param' }, 400);
    if (!dataCenter) return json({ error: 'Missing dataCenter' }, 400);
    if (!env.IAM_API_TOKEN) return json({ error: 'Missing IAM_API_TOKEN env var' }, 500);

    const iamUrl = `https://${dataCenter}/api/iam/users?${encodeURIComponent('emails[]')}=${encodeURIComponent(email)}`;
    const upstream = await fetch(iamUrl, {
      method: 'GET',
      headers: { Authorization: `Bearer ${env.IAM_API_TOKEN}`, 'Content-Type': 'application/json' },
    });
    const text = await upstream.text();
    return new Response(text, {
      status: upstream.status,
      headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json' },
    });
  }

  // ---------------------------------------------------------------
  // GET /api/health
  // ---------------------------------------------------------------
  if (pathname === '/api/health' && request.method === 'GET') {
    return json({
      ok: true,
      genieApiKeyConfigured: Boolean(env.GENIE_API_KEY),
      iamTokenConfigured: Boolean(env.IAM_API_TOKEN),
      defaultGenieId: env.GENIE_ID || null,
      defaultDataCenter: env.DATA_CENTER || null,
    });
  }

  return json({ error: 'Not found' }, 404);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    try {
      if (url.pathname.startsWith('/api/')) {
        return await handleApi(request, env, url);
      }
      // Everything else (index.html, etc.) — serve as a normal static file.
      return env.ASSETS.fetch(request);
    } catch (err) {
      return json({ error: 'Upstream request failed', detail: err.message }, 502);
    }
  },
};
