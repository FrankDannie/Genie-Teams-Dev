# Genie Connect

A small console for testing a Workato Genie's headless chat API and resolving
end user IDs by email — with all secrets held server-side.

```
genie-connect-project/
├── server.js          Express proxy — holds GENIE_API_KEY and IAM_API_TOKEN
├── package.json
├── .env.example       Copy to .env and fill in
├── .gitignore
└── public/
    ├── index.html         Full console (Settings tab, empty Console tab, chat lives in the widget)
    └── widget-embed.js    Standalone chat widget — a single <script> tag, for embedding on OTHER pages
```

This Express app is the local/dev version. Sibling folders hold the
deployable versions this project grew into:
- `cloudflare-pages-deployment/` — **the recommended way to actually run
  this anywhere** — console + API as one Cloudflare Pages project, no
  server to host or manage. See "Deploying this with nothing to host or
  run" below.
- `github-pages-deployment/` — an alternative split: static site on
  GitHub Pages calling a separately-deployed Cloudflare Worker, with
  Microsoft single sign-on. More moving parts; only worth it for silent
  SSO.
- `teams-bot/` — a separate, actual conversational **Teams bot** (chats in
  a 1:1 Teams DM), built on Azure Functions + Bot Framework. Unrelated to
  the tab — see its own README.

## Why a backend at all

The browser cannot call `genie-api.workato.com` or the IAM API directly with
a bearer token — Workato's APIs aren't set up to accept cross-origin requests
from arbitrary front-ends, and shipping a secret to client-side JS would
expose it to anyone who opens dev tools. This server sits in between: the
browser calls `localhost`, and the server attaches the real credentials
before forwarding upstream.

## How auth works

One flow: a server-side `GENIE_API_KEY` authenticates the app, and each
request is scoped to a specific person via `X-IDP-User-ID`. You get that ID
either by typing it directly or by looking it up from an email address (the
console's "Look up ID by email" panel), which calls the IAM API using a
separate `IAM_API_TOKEN` — also held server-side.

## Setup

```bash
npm install
cp .env.example .env
```

Fill in `.env`:

```
GENIE_API_KEY=<your genie's API key, from the genie build page>
GENIE_ID=gin-AbMAK4r6-rXgonW-CD
IAM_API_TOKEN=<a workspace API client token with Identity IAM scope>
DATA_CENTER=www.workato.com
```

Then run:

```bash
npm start
```

Open **http://localhost:3000**.

## What each route does

| Route | Method | Purpose |
|---|---|---|
| `/api/genie/conversations` | POST | Creates a conversation for a genie + end user, returns `conversation_id` |
| `/api/genie/conversations/:id/messages` | POST | Sends a message and streams the reply back to the browser |
| `/api/iam/users` | GET | Looks up a Workato Identity user by email (`?email=`, `?dataCenter=`) |
| `/api/health` | GET | Reports whether `GENIE_API_KEY` / `IAM_API_TOKEN` are configured, without revealing them |

## Security notes

- `.env` is git-ignored. Never commit real values from `.env.example`.
- Set `ALLOWED_ORIGIN` in `.env` to your actual frontend origin before deploying anywhere beyond localhost — the default of `*` is for local testing only.
- The IAM token and genie API key are workspace/genie-level credentials. Anyone who has them can read conversations or enumerate users, so keep `.env` out of any shared drive or repo.
- **Rotate any credential that's ever been pasted into a chat, ticket, or document.** Treat it as seen by more people than intended, regardless of what the client is.

## Deploying this with nothing to host or run

**Recommended: `cloudflare-pages-deployment/`** — one Cloudflare Pages
project serves the console *and* the API together. You connect a GitHub
repo in Cloudflare's dashboard once; every `git push` after that rebuilds
and redeploys automatically. There's no server process to start, keep
alive, or patch — see that folder's own README for the exact
click-by-click setup (about 5 minutes, all in a browser, no CLI needed).

**Teams tab:** `teams-tab-manifest/` at this root packages the console as
a personal Teams tab pointing at your `.pages.dev` URL — no Microsoft
sign-in setup needed, identity comes from the **Look up ID by email**
panel already in Settings. Fill in the URL + a fresh GUID in
`manifest.json`, zip it with the two icons, sideload via **Teams → Apps
→ Upload a custom app**.

*(Two other folders exist for different needs, both more involved than
the above: `github-pages-deployment/` splits the site and API into two
separate pieces with Microsoft single sign-on — only worth it if you
specifically want silent sign-in. `teams-bot/` is an actual
conversational Teams bot via Azure Functions, not a tab at all. Neither
is needed unless you have that specific requirement.)*

## Embedding the widget on another page

`public/widget-embed.js` is the chat bubble + popup only — no console, no
settings, nothing else — meant to be dropped onto a page your own app
already serves (e.g. a page registered as an Enterprise Application in
Microsoft Entra ID). Add one script tag:

```html
<script
  src="https://YOUR-GENIE-CONNECT-HOST/widget-embed.js"
  data-base-url="https://YOUR-GENIE-CONNECT-HOST"
  data-genie-id="gin-AbMAK4r6-rXgonW-CD"
  data-interface-name=Smart Genie""
  data-idp-user-id=""
></script>
```

- `data-base-url` — required whenever the embedding page is on a different
  origin than this server (it will be). Point it at wherever you deploy
  this `server.js`.
- `data-idp-user-id` — if your host page already knows the signed-in user
  (e.g. resolved via your app's own Entra SSO), pass it here, or set it
  after the fact:
  ```js
  window.GenieWidget.setUser("usr_2f8b1c...");
  ```
  If left blank, the widget asks for it once inline, no separate settings
  page needed.
- Also available: `window.GenieWidget.open()`, `.close()`, `.setGenieId(id)`.

**Cross-origin:** since the embedding page and this server are different
origins, set `ALLOWED_ORIGIN` in `.env` to that page's origin (or a
comma-aware value your deployment enforces) so the browser's CORS check
passes.

**On Microsoft Entra's "My Apps" portal specifically:** that portal
(myapplications.microsoft.com) only launches to a registered app's URL —
it has no mechanism to host a third-party widget inside its own page
chrome. This script is for the page your app itself serves once a user
gets there, not for injecting into the My Apps portal directly.

