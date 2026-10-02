/*!
 * Genie Connect — Cloudflare Worker
 * -----------------------------------------------------------------------
 * The ONLY place GENIE_API_KEY and IAM_API_TOKEN live. They are set as
 * Worker secrets (never committed to git, never sent to the browser) via:
 *
 *   wrangler secret put GENIE_API_KEY
 *   wrangler secret put IAM_API_TOKEN
 *
 * The static site (GitHub Pages) calls THIS worker, never Workato
 * directly. This worker:
 *   1. Adds the Access-Control-Allow-Origin header Workato doesn't send,
 *      so the browser's CORS check passes.
 *   2. Attaches the real bearer tokens server-side, so they never appear
 *      in any request the browser makes or any file shipped to it.
 *
 * Routes (mirrors the original Node/Express server.js 1:1):
 *   POST /api/genie/conversations
 *   POST /api/genie/conversations/:conversationId/messages   (streamed)
 *   GET  /api/iam/users?email=...&dataCenter=...
 *   GET  /api/health
 * -----------------------------------------------------------------------
 */

function corsHeaders(env) {
  return {
    'Access-Control-Allow-Origin': env.ALLOWED_ORIGIN || '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
  };
}

function json(data, status, env) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(env) },
  });
}

function genieBase(genieId) {
  return `https://genie-api.workato.com/api/v1/genies/${genieId}/chat`;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const { pathname } = url;
    const cors = corsHeaders(env);

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    try {
      // ---------------------------------------------------------------
      // POST /api/genie/conversations
      // ---------------------------------------------------------------
      if (pathname === '/api/genie/conversations' && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const genieId = body.genieId || env.GENIE_ID;
        const idpUserId = body.idpUserId;

        if (!genieId) return json({ error: 'Missing genieId' }, 400, env);
        if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400, env);
        if (!env.GENIE_API_KEY) return json({ error: 'Worker is missing GENIE_API_KEY' }, 500, env);

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
          headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json', ...cors },
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

        if (!genieId) return json({ error: 'Missing genieId' }, 400, env);
        if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400, env);
        if (!message) return json({ error: 'Missing message' }, 400, env);
        if (!env.GENIE_API_KEY) return json({ error: 'Worker is missing GENIE_API_KEY' }, 500, env);

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
          return new Response(errText, { status: upstream.status, headers: cors });
        }

        // Pass the upstream stream straight through to the browser.
        return new Response(upstream.body, {
          status: 200,
          headers: {
            'Content-Type': upstream.headers.get('content-type') || 'text/event-stream',
            'Cache-Control': 'no-cache',
            Connection: 'keep-alive',
            ...cors,
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

        if (!email) return json({ error: 'Missing email query param' }, 400, env);
        if (!dataCenter) return json({ error: 'Missing dataCenter' }, 400, env);
        if (!env.IAM_API_TOKEN) return json({ error: 'Worker is missing IAM_API_TOKEN' }, 500, env);

        const iamUrl = `https://${dataCenter}/api/iam/users?${encodeURIComponent('emails[]')}=${encodeURIComponent(email)}`;
        const upstream = await fetch(iamUrl, {
          method: 'GET',
          headers: { Authorization: `Bearer ${env.IAM_API_TOKEN}`, 'Content-Type': 'application/json' },
        });
        const text = await upstream.text();
        return new Response(text, {
          status: upstream.status,
          headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json', ...cors },
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
        }, 200, env);
      }

      return json({ error: 'Not found' }, 404, env);
    } catch (err) {
      return json({ error: 'Upstream request failed', detail: err.message }, 502, env);
    }
  },
};
