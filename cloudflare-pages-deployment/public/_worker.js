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
 * GENIE_API_KEY, IAM_API_TOKEN, and DATATABLE_API_TOKEN are never in this
 * file — they're set as encrypted environment variables in the Pages
 * dashboard (Settings → Environment variables → "Encrypt"), or via
 * `wrangler pages secret put` if you prefer the CLI. Either way, nothing
 * secret is ever in git.
 *
 * Because the site and this API are the exact same Pages project, every
 * fetch('/api/...') call from index.html/widget.html/widget-embed.js is
 * same-origin — no CORS setup needed.
 *
 * Routes:
 *   POST /api/genie/conversations
 *   POST /api/genie/conversations/:conversationId/messages   (streamed)
 *   GET  /api/iam/users?email=...&dataCenter=...
 *   GET  /api/history?idpUserId=...&dataCenter=...                     [chat history: list]
 *   POST /api/history                                                  [chat history: upsert]
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

function cleanDataCenter(raw, env) {
  return (raw || env.DATA_CENTER || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
}

// A Workato API Platform endpoint (built from a recipe), used for both
// listing and upserting history rows. Auth is an `api-token` header (NOT
// `Authorization: Bearer`, unlike the Genie/IAM calls above — this
// endpoint's own convention). NOTE — the exact field names this endpoint
// expects/returns still haven't been confirmed against its actual
// contract; these match the Data Table's own column names as a best
// guess. A 400 (or a response that just looks "empty"/wrong-shaped) most
// likely means the field names are wrong.
function historyApiUrl(env) {
  return env.HISTORY_API_URL || 'https://apim.workato.com/frankd308/headless-v1/genie-chat-history-api';
}

// Cloudflare env vars pasted from a dashboard text box very commonly carry
// an invisible trailing newline or leading/trailing space from the copy —
// which breaks header auth while LOOKING identical to the correct value
// everywhere you'd visually check it. Trim defensively; it's a no-op if
// the value was already clean.
function historyApiToken(env) {
  return (env.DATATABLE_API_TOKEN || '').trim();
}

// Masked preview for /api/health — never return the real token. Lets you
// confirm server-side what's actually configured (length + first/last 2
// chars) without exposing the secret, so you can tell "empty", "has stray
// whitespace" (length looks off), and "wrong value" (the preview text
// won't match what you expect) apart from each other.
function maskedTokenPreview(value) {
  if (!value) return null;
  if (value.length <= 6) return `${value.length} chars (too short to preview safely)`;
  return `${value.length} chars, "${value.slice(0, 2)}...${value.slice(-2)}"`;
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
    const dataCenter = cleanDataCenter(url.searchParams.get('dataCenter'), env);

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
  // GET /api/history?idpUserId=...&dataCenter=...
  // List this user's saved conversations, newest first.
  // ---------------------------------------------------------------
  if (pathname === '/api/history' && request.method === 'GET') {
    const idpUserId = url.searchParams.get('idpUserId');

    if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400);
    const dtToken = historyApiToken(env);
    if (!dtToken) return json({ error: 'Missing DATATABLE_API_TOKEN env var' }, 500);

    const listUrl = `${historyApiUrl(env)}?idp_user_id=${encodeURIComponent(idpUserId)}`;
    const upstream = await fetch(listUrl, {
      headers: { 'api-token': dtToken },
    });
    const text = await upstream.text();
    return new Response(text, {
      status: upstream.status,
      headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json' },
    });
  }

  // ---------------------------------------------------------------
  // POST /api/history
  // Upsert one conversation's history row. This endpoint is assumed to
  // handle find-or-create itself (by conversation_id) on the Workato
  // side, since only one URL was given for both operations — there's no
  // separate "update by record id" path to call here.
  // ---------------------------------------------------------------
  if (pathname === '/api/history' && request.method === 'POST') {
    const body = await request.json().catch(() => ({}));
    const { idpUserId, genieId, conversationId, title, messages } = body;

    if (!idpUserId) return json({ error: 'Missing idpUserId' }, 400);
    if (!conversationId) return json({ error: 'Missing conversationId' }, 400);
    const dtTokenPost = historyApiToken(env);
    if (!dtTokenPost) return json({ error: 'Missing DATATABLE_API_TOKEN env var' }, 500);

    const row = {
      idp_user_id: idpUserId,
      genie_id: genieId || '',
      conversation_id: conversationId,
      title: title || '',
      updated_at: new Date().toISOString(),
      messages: JSON.stringify(messages || []),
    };

    try {
      const upstream = await fetch(historyApiUrl(env), {
        method: 'POST',
        headers: { 'api-token': dtTokenPost, 'Content-Type': 'application/json' },
        body: JSON.stringify(row),
      });
      const text = await upstream.text();
      return new Response(text, {
        status: upstream.status,
        headers: { 'Content-Type': upstream.headers.get('content-type') || 'application/json' },
      });
    } catch (err) {
      return json({ error: 'History API request failed', detail: err.message }, 502);
    }
  }

  // ---------------------------------------------------------------
  // GET /api/health
  // ---------------------------------------------------------------
  if (pathname === '/api/health' && request.method === 'GET') {
    return json({
      ok: true,
      genieApiKeyConfigured: Boolean(env.GENIE_API_KEY),
      iamTokenConfigured: Boolean(env.IAM_API_TOKEN),
      dataTableConfigured: Boolean(historyApiToken(env)),
      dataTableTokenPreview: maskedTokenPreview(historyApiToken(env)), // confirms what the server actually sees — never the real value
      historyApiUrl: historyApiUrl(env), // confirms which URL is actually being called (vs. the HISTORY_API_URL override)
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
      // Cloudflare Pages already handles clean URLs itself (/widget serves
      // widget.html, and redirects /widget.html -> /widget) — do not add
      // another redirect here, it'll loop against that one.
      return env.ASSETS.fetch(request);
    } catch (err) {
      return json({ error: 'Upstream request failed', detail: err.message }, 502);
    }
  },
};