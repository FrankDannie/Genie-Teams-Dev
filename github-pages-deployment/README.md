# Genie Connect — static GitHub Pages client

Two pieces:

```
github-pages-deployment/
├── site/                 What you push to GitHub Pages
│   ├── index.html        Sidebar app: Chat + Settings tabs
│   └── app.js            MSAL sign-in, settings, chat logic — no secrets
└── worker/                Deploy this separately to Cloudflare (free)
    ├── worker.js          Holds GENIE_API_KEY / IAM_API_TOKEN, adds CORS
    ├── wrangler.toml
    ├── .dev.vars.example  Copy to .dev.vars for local testing (git-ignored)
    └── .gitignore
```

**Why both exist:** tested directly — `genie-api.workato.com` and the IAM
API don't send `Access-Control-Allow-Origin`, so a browser will block calls
to them from a `github.io` page no matter what the JS says. The worker is
the only thing in this setup that calls Workato directly; everything else
only ever talks to the worker. It's not a server you run or patch — it's a
stateless function Cloudflare hosts, free tier is generous for this volume.

## Does this need a .env?

- **The static site (`site/`)**: no env, no build step. It's plain HTML +
  JS — open it or serve it as-is.
- **The worker (`worker/`)**: yes, but not a `.env` file in production —
  `GENIE_API_KEY` and `IAM_API_TOKEN` are set as Cloudflare secrets via
  `wrangler secret put` (encrypted on Cloudflare's side, not stored as a
  file at all). For *local* testing only, you use a `.dev.vars` file —
  copy `worker/.dev.vars.example` to `worker/.dev.vars` and fill in real
  values. It's git-ignored, and `wrangler dev` (below) reads it automatically.

## Test it locally before deploying anywhere

**1. Run the worker locally:**
```bash
cd github-pages-deployment/worker
npm install -g wrangler   # one-time, if you don't have it
wrangler login            # one-time
cp .dev.vars.example .dev.vars
# edit .dev.vars, paste your real GENIE_API_KEY and IAM_API_TOKEN
wrangler dev
```
This starts the worker at `http://localhost:8787` and prints that URL.
Leave this running in its own terminal.

**2. Serve the static site locally** (don't just double-click `index.html` —
`file://` pages can't do fetch/CORS or MSAL redirects properly):
```bash
cd github-pages-deployment/site
npx serve .
# or: python3 -m http.server 5500
```
Open the URL it prints (e.g. `http://localhost:5500` or `http://localhost:3000`).

**3. Point it at your local worker:** in the app's **Settings** tab, set
**Worker base URL** to `http://localhost:8787` (from step 1), and set
**Genie ID**. Skip Microsoft sign-in for this first test — just type a
known Workato end user ID directly into **End user ID** and check
**Remember on this device**.

**4. Send a message** in the **Chat** tab. If it streams back a reply,
the worker ↔ Workato ↔ browser path is confirmed end to end. Common
failure signs:
- *"Worker is missing GENIE_API_KEY"* → `.dev.vars` wasn't picked up —
  confirm it's in `github-pages-deployment/worker/` (same folder as
  `worker.js`) and restart `wrangler dev`.
- *Network error / failed to fetch* → worker URL in Settings doesn't match
  what `wrangler dev` printed, or it's not running.
- *401/403 from the worker* → the key/token value itself is wrong or
  revoked — double check what's in `.dev.vars`.

Only once this works locally is it worth deploying for real (step 1 below)
and wiring up Entra sign-in (step 3 below) — both are one-time setup, not
things you need for day-to-day testing.

## 1. Deploy the worker

```bash
cd github-pages-deployment/worker
npm install -g wrangler   # one-time, if you don't have it
wrangler login
wrangler secret put GENIE_API_KEY     # paste when prompted
wrangler secret put IAM_API_TOKEN     # paste when prompted
wrangler deploy
```

This prints your worker's URL, something like:
`https://genie-connect-proxy.yoursubdomain.workers.dev`

Edit `wrangler.toml` first if you want different defaults for `GENIE_ID` /
`DATA_CENTER`. Once you know your GitHub Pages URL (step 2), come back and
set `ALLOWED_ORIGIN` to that exact origin, then `wrangler deploy` again —
leaving it as `*` works but means any site could call your worker.

## 2. Push the static site to GitHub Pages

```bash
# from the github-pages-deployment/site folder
git init
git add .
git commit -m "Genie Connect static client"
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

Then in the repo: **Settings → Pages → Source: `main` branch, `/ (root)`**.
Your page will be at `https://YOUR-USERNAME.github.io/YOUR-REPO/`.

Open it, go to **Settings** tab, paste the worker URL from step 1 into
**Worker base URL**. That's saved in this browser's localStorage — nothing
to rebuild or redeploy when you change it.

## 3. Wire up Microsoft Entra sign-in (optional but answers your "pull the email automatically" ask)

1. In Entra ID → **App registrations → New registration**.
2. Platform: **Single-page application (SPA)** — no client secret is created or needed.
3. Redirect URI: your GitHub Pages URL from step 2, exactly
   (`https://YOUR-USERNAME.github.io/YOUR-REPO/`).
4. Copy the **Application (client) ID** and **Directory (tenant) ID** into
   the top of `site/app.js`:
   ```js
   const MSAL_CONFIG = {
     clientId: 'paste-here',
     tenantId: 'paste-here',
   };
   ```
5. Commit + push that change; GitHub Pages picks it up automatically.
6. Add this app's tile to **Entra ID → Enterprise applications → My Apps**
   pointing at the GitHub Pages URL, same as any other SSO tile.

How the identity flow works once this is set:
- First visit: user clicks **Sign in with Microsoft** in Settings (one popup).
- The widget takes their email from the Entra token, calls the worker's IAM
  lookup automatically, resolves their Workato user ID, and remembers it
  in this browser's localStorage — no typing.
- Every later visit: MSAL's own cache (also localStorage) recognizes the
  signed-in account, so sign-in is silent and the ID is already filled in.
- If MSAL isn't configured, or sign-in is skipped, the **End user ID**
  field with **Remember on this device** still works as a manual fallback.

## Using it as a Microsoft Teams tab

`teams-tab-manifest/` packages this same site as a **personal Teams tab** —
the chat/settings UI above loads inside a Teams window, no separate bot or
backend. (If you want an actual conversational bot that lives in a 1:1 Teams
chat instead, that's the unrelated `../teams-bot/` folder.)

```
teams-tab-manifest/
├── manifest.json   Fill in the placeholders, then zip with the two icons
├── color.png        192×192 placeholder icon (swap for your own)
└── outline.png      32×32 placeholder icon (swap for your own)
```

Two extra files already live in `site/` for this to work:
`teams-auth-start.html` and `teams-auth-end.html` — a small bounce pair
that runs the Microsoft sign-in flow in a Teams-managed popup window
instead of `loginPopup()`, which Teams normally blocks when called from
inside a tab's iframe. They reuse the exact same MSAL client/tenant ID as
`app.js`; outside of Teams, sign-in still uses the regular `loginPopup()`
path and these two files are never touched.

**Setup:**

1. Deploy `site/` to GitHub Pages and the worker, as above.
2. In your Entra ID app registration, add a **second** redirect URI:
   `https://YOUR-USERNAME.github.io/YOUR-REPO/teams-auth-end.html`
   (keep the existing one from step 3 above too).
3. Put the same `clientId` / `tenantId` into the `MSAL_CONFIG` object at
   the top of **both** `site/teams-auth-start.html` and
   `site/teams-auth-end.html` (copy from `site/app.js`).
4. In `teams-tab-manifest/manifest.json`:
   - `id` — generate a fresh GUID (any GUID generator — this is the Teams
     app's own ID, unrelated to the Entra app ID).
   - `staticTabs[0].contentUrl` and `websiteUrl` — your GitHub Pages URL.
   - `validDomains` — just the GitHub Pages hostname (e.g.
     `yourname.github.io`). `login.microsoftonline.com` does NOT go here —
     Teams opens that in its own popup window, not inside the tab iframe.
   - `developer` / `name` / `description` — fill in for your org; swap the
     icons for real artwork if you want.
5. Zip the three files together (`manifest.json`, `color.png`,
   `outline.png` — flat, no subfolder) into `genie-connect-teams-tab.zip`.
6. In Teams: **Apps → Manage your apps → Upload an app → Upload a custom
   app**, select the zip. (Needs custom/sideloaded apps enabled for your
   tenant — otherwise go through **Teams admin center → Teams apps →
   Manage apps** to publish it org-wide.)

**What still works the same:** the worker, CORS, and `ALLOWED_ORIGIN`
below are completely unaffected — Teams just displays your existing
GitHub Pages origin in an iframe, so every fetch still comes from (and
needs to be allowed from) that same origin.

**Known limitation:** sign-out inside Teams only clears the account from
this tab's view — `logoutPopup()` is skipped (Teams blocks that popup too),
so the underlying MSAL cache entry is cleared from a regular browser tab
instead, if you need a full logout.

## Security notes

- The worker is the only holder of `GENIE_API_KEY` / `IAM_API_TOKEN` —
  set via `wrangler secret put`, encrypted at rest, never in git, never
  shipped to the browser.
- Once you know your GitHub Pages origin, set `ALLOWED_ORIGIN` in
  `wrangler.toml` (or `wrangler secret put ALLOWED_ORIGIN` / dashboard var)
  to that exact origin instead of `*`, so only your page can call the worker.
- Nothing in `site/` needs to be secret — the worker URL and MSAL client ID
  are not credentials, they're safe to have in a public repo.
- **Rotate `GENIE_API_KEY` / `IAM_API_TOKEN` if either was ever pasted into
  a chat, ticket, or the old `.env`-based server's commit history** —
  treat anything that left your machine as seen by more people than intended.
