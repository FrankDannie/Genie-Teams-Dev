# Genie Connect — Cloudflare Pages (one deploy, no server to run)

This is the "I don't want to host or manage a backend" version. One
folder, one Cloudflare Pages project, one git push to deploy both the
site and the API together.

```
cloudflare-pages-deployment/
└── public/              This whole folder is what gets deployed
    ├── index.html        The console (same one as the local server.js version)
    └── _worker.js         The API — see the big comment at its top for how this works
```

## Why this isn't "hosting a backend"

You're not running anything, not keeping a process alive, not patching an
OS. `_worker.js` is a plain JavaScript file that Cloudflare's own
infrastructure runs on demand, for free, whenever a request comes in for
one of the `/api/...` paths — same as the static HTML is served on
demand. There's no "server" to go down, restart, or sit idle costing
money. This is genuinely closer to "upload code and it runs" than any
traditional host (Render, Azure App Service, etc.) can be, because there
is no process boundary between "my site" and "my backend" to manage —
it's one deployment.

## One-time setup (all in the Cloudflare dashboard — no CLI required)

1. Push this `cloudflare-pages-deployment/` folder to a GitHub repo
   (can be the same repo as everything else, this folder just needs to
   be the one you point Cloudflare at).
2. [dash.cloudflare.com](https://dash.cloudflare.com) → **Workers & Pages
   → Create → Pages → Connect to Git** → pick the repo.
3. Build settings:
   - **Build command:** leave blank (nothing to build — it's already
     plain HTML/JS).
   - **Build output directory:** `cloudflare-pages-deployment/public`
     (or `public` if you made this folder its own repo root).
4. **Settings → Environment variables** → add, for Production:
   - `GENIE_API_KEY` (click **Encrypt**)
   - `IAM_API_TOKEN` (click **Encrypt**)
   - `GENIE_ID` — e.g. `gin-AbMAK4r6-rXgonW-CD`
   - `DATA_CENTER` — e.g. `www.workato.com`
5. **Save and Deploy.** Cloudflare builds it and gives you a URL like
   `https://genie-connect.pages.dev`.

From here on: every time you `git push`, Cloudflare automatically
rebuilds and redeploys. You never log into anything, run a deploy
command, or SSH anywhere again.

## OAuth 2.0 (PKCE) mode — sign in with Workato Identity

Instead of one shared API key plus an email lookup, each person signs in
through Workato Identity and the widget calls the Headless API **as that
user**. Chat history then comes from the Headless API itself (no Data Table).

1. **Agent Studio → your genie → Chat interface → Custom chat interface**:
   create the client with **OAuth 2.0 (PKCE)**. Set its redirect URL to
   `https://genie-teams-dev.pages.dev/oauth-callback` (must match exactly,
   or your own domain + `/oauth-callback`). Copy the **Client ID** with the
   Copy button.
2. **Pages → Settings → Environment variables**: add `OAUTH_CLIENT_ID`.
   Optional: `OAUTH_REDIRECT_URI` (if not `<site>/oauth-callback`),
   `OAUTH_SCOPE` (default `openid profile email` — never add
   `offline_access`), `OAUTH_ISSUER` (default `https://id.workato.com`).
3. Push to deploy. Check `/api/health` shows `oauthConfigured: true`.
4. Use `/widget?auth=oauth&genie=gin-XXXX&name=My%20Genie`. Without
   `auth=oauth` the widget behaves exactly as before (API key + email).

A genie supports one attached client, so one deployment's `OAUTH_CLIENT_ID`
serves one genie; `genie=` must be that genie's ID.

How it works: the browser does the PKCE redirect (a popup when embedded in
SharePoint, a normal redirect in its own tab); `/api/oauth/token` does the
code/refresh exchange server-side (no client secret exists); `/api/headless/…`
forwards the user's own Bearer token to `genie-api.workato.com` for an
allow-listed set of paths. Access tokens last about an hour and refresh
silently; refresh tokens rotate and are kept in the browser's localStorage.

## Using it as a Microsoft Teams tab

Same idea as before — point `teams-tab-manifest/manifest.json` (at the
project root) at this URL instead of a Render/Azure one:
- `staticTabs[0].contentUrl` and `websiteUrl` →
  `https://genie-connect.pages.dev/`
- `validDomains` → `["genie-connect.pages.dev"]`
- (or your custom domain, if you add one under **Custom domains** in the
  Pages project — either works, just make the manifest match whichever
  URL you actually use)

Then zip `manifest.json` + `color.png` + `outline.png` and sideload via
**Teams → Apps → Upload a custom app**, same as documented at the
project root. No Microsoft sign-in setup needed — identity still comes
from the **Look up ID by email** panel already in Settings.

## Local testing (optional)

```bash
npm install -g wrangler   # one-time
cd cloudflare-pages-deployment
wrangler pages dev public --binding GENIE_API_KEY=xxx --binding IAM_API_TOKEN=xxx --binding GENIE_ID=gin-... --binding DATA_CENTER=www.workato.com
```
This runs the exact same `_worker.js` + static files locally before you
ever push anything. Skip this entirely if you'd rather just push and
test on the real `.pages.dev` URL.

## Security notes

- `GENIE_API_KEY` / `IAM_API_TOKEN` live only as encrypted Pages
  environment variables — never in this folder, never in git.
- Nothing here needs CORS configuration — the site and the API are the
  same origin by construction, so there's no cross-origin call to allow.
- **Rotate either credential if it was ever pasted into a chat, ticket,
  or document.**