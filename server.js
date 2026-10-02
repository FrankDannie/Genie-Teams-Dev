import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();

const PORT = process.env.PORT || 3000;
const GENIE_API_KEY = process.env.GENIE_API_KEY;
const DEFAULT_GENIE_ID = process.env.GENIE_ID;
const IAM_API_TOKEN = process.env.IAM_API_TOKEN;
const DEFAULT_DATA_CENTER = process.env.DATA_CENTER || 'www.workato.com';
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || '*';

if (!GENIE_API_KEY) {
  console.warn('[warn] GENIE_API_KEY is not set. Genie chat calls will fail until it is configured in .env');
}
if (!IAM_API_TOKEN) {
  console.warn('[warn] IAM_API_TOKEN is not set. Email → user ID lookup will fail until it is configured in .env');
}

app.use(cors({ origin: ALLOWED_ORIGIN }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ---------------------------------------------------------------------------
// Genie Chat API proxy
// The genie API key never leaves this server. The browser only ever talks
// to these local /api/genie/* routes, identifying which end user is chatting
// via idpUserId.
// ---------------------------------------------------------------------------

function genieBase(genieId) {
  return `https://genie-api.workato.com/api/v1/genies/${genieId}/chat`;
}

app.post('/api/genie/conversations', async (req, res) => {
  try {
    const genieId = req.body.genieId || DEFAULT_GENIE_ID;
    const idpUserId = req.body.idpUserId;

    if (!genieId) return res.status(400).json({ error: 'Missing genieId' });
    if (!idpUserId) return res.status(400).json({ error: 'Missing idpUserId' });
    if (!GENIE_API_KEY) return res.status(500).json({ error: 'Server is missing GENIE_API_KEY' });

    const upstream = await fetch(`${genieBase(genieId)}/conversations`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${GENIE_API_KEY}`,
        'X-IDP-User-ID': idpUserId,
        'Content-Type': 'application/json',
      },
    });

    const text = await upstream.text();
    res.status(upstream.status);
    res.set('Content-Type', upstream.headers.get('content-type') || 'application/json');
    res.send(text);
  } catch (err) {
    console.error('create conversation error', err);
    res.status(502).json({ error: 'Upstream request failed', detail: err.message });
  }
});

app.post('/api/genie/conversations/:conversationId/messages', async (req, res) => {
  try {
    const { conversationId } = req.params;
    const genieId = req.body.genieId || DEFAULT_GENIE_ID;
    const idpUserId = req.body.idpUserId;
    const message = req.body.message;

    if (!genieId) return res.status(400).json({ error: 'Missing genieId' });
    if (!idpUserId) return res.status(400).json({ error: 'Missing idpUserId' });
    if (!message) return res.status(400).json({ error: 'Missing message' });
    if (!GENIE_API_KEY) return res.status(500).json({ error: 'Server is missing GENIE_API_KEY' });

    const upstream = await fetch(`${genieBase(genieId)}/conversations/${conversationId}/messages`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${GENIE_API_KEY}`,
        'X-IDP-User-ID': idpUserId,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message, stream: true }),
    });

    if (!upstream.ok || !upstream.body) {
      const errText = await upstream.text();
      return res.status(upstream.status).send(errText);
    }

    res.set('Content-Type', upstream.headers.get('content-type') || 'text/event-stream');
    res.set('Cache-Control', 'no-cache');
    res.set('Connection', 'keep-alive');

    const reader = upstream.body.getReader();
    req.on('close', () => reader.cancel().catch(() => {}));

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      res.write(Buffer.from(value));
    }
    res.end();
  } catch (err) {
    console.error('send message error', err);
    if (!res.headersSent) {
      res.status(502).json({ error: 'Upstream request failed', detail: err.message });
    } else {
      res.end();
    }
  }
});

// ---------------------------------------------------------------------------
// Workato Identity IAM proxy — resolve an end user ID from an email address.
// The IAM token never leaves this server.
// ---------------------------------------------------------------------------

app.get('/api/iam/users', async (req, res) => {
  try {
    const email = req.query.email;
    const dataCenter = (req.query.dataCenter || DEFAULT_DATA_CENTER || '').replace(/^https?:\/\//, '').replace(/\/$/, '');

    if (!email) return res.status(400).json({ error: 'Missing email query param' });
    if (!dataCenter) return res.status(400).json({ error: 'Missing dataCenter' });
    if (!IAM_API_TOKEN) return res.status(500).json({ error: 'Server is missing IAM_API_TOKEN' });

    const url = `https://${dataCenter}/api/iam/users?${encodeURIComponent('emails[]')}=${encodeURIComponent(email)}`;

    const upstream = await fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${IAM_API_TOKEN}`,
        'Content-Type': 'application/json',
      },
    });

    const text = await upstream.text();
    res.status(upstream.status);
    res.set('Content-Type', upstream.headers.get('content-type') || 'application/json');
    res.send(text);
  } catch (err) {
    console.error('iam lookup error', err);
    res.status(502).json({ error: 'Upstream request failed', detail: err.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    genieApiKeyConfigured: Boolean(GENIE_API_KEY),
    iamTokenConfigured: Boolean(IAM_API_TOKEN),
    defaultGenieId: DEFAULT_GENIE_ID || null,
    defaultDataCenter: DEFAULT_DATA_CENTER,
  });
});

app.listen(PORT, () => {
  console.log(`Genie Connect server listening on http://localhost:${PORT}`);
});
