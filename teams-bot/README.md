# Genie Connect — Teams bot

Connects a Teams bot to your Workato Genie's headless chat API via Azure
Functions + the Bot Framework SDK, authenticated with a **User-Assigned
Managed Identity** — no bot client secret to store or rotate.

```
teams-bot/
├── package.json
├── host.json
├── local.settings.json.example   Copy to local.settings.json for local testing
├── src/
│   ├── genieClient.js            Calls Workato — same endpoints as worker.js
│   ├── state.js                  Per-user identity, persisted in Blob Storage
│   ├── bot.js                    Bot logic: resolve identity, call Genie, reply
│   └── functions/
│       └── messages.js           HTTP trigger = the bot's messaging endpoint
└── teams-manifest/
    ├── manifest.json             Fill in the two placeholder IDs, then zip
    ├── color.png                 192×192 placeholder icon (swap for your own)
    └── outline.png               32×32 placeholder icon (swap for your own)
```

Unlike the GitHub Pages site, there's no CORS/key-exposure problem here —
Teams clients never call your API directly. Every message is relayed
through Microsoft's Bot Framework Service to this Function, which holds
all the secrets and replies on your behalf.

## How a message flows

1. User messages the bot in Teams → Bot Framework Service POSTs an Activity
   to your Function's `/api/messages` endpoint.
2. `messages.js` verifies it's a genuine request (handled by the SDK's
   `CloudAdapter`) and hands it to `bot.js`.
3. On first contact, `bot.js` asks Teams for the user's email
   (`TeamsInfo.getMember`), looks it up against your IAM API to resolve a
   Workato user ID, and caches that mapping in Blob Storage so it's skipped
   on every later message.
4. It creates (or reuses) a Genie conversation for that user and sends the
   message, collecting the full reply server-side (Teams doesn't do
   token-by-token streaming the way the web widget does).
5. The reply is sent back to Teams via the Bot Framework Connector.

Say `/reset` in the chat to start a fresh Genie conversation without
losing the cached identity mapping.

## 1. Create the User-Assigned Managed Identity

```bash
az identity create --name genie-connect-bot-identity --resource-group YOUR_RG
```
Note the output's `clientId` and `tenantId` — you'll need both twice below.

## 2. Create the Azure Bot resource

Azure Portal → **Create a resource → Azure Bot**:
- **Type of App**: User-Assigned Managed Identity
- **Managed Identity**: the one you just created
- Once created, go to **Configuration** and set the **Messaging endpoint**
  to `https://YOUR-FUNCTION-APP.azurewebsites.net/api/messages` (you'll
  have this URL after step 3 — you can come back and set it).
- Under **Channels**, add **Microsoft Teams**.

## 3. Create and configure the Function App

```bash
az functionapp create \
  --name YOUR-FUNCTION-APP \
  --resource-group YOUR_RG \
  --storage-account YOUR_STORAGE_ACCOUNT \
  --consumption-plan-location YOUR_REGION \
  --runtime node --runtime-version 20 --functions-version 4

# Assign the SAME managed identity to the Function App
az functionapp identity assign \
  --name YOUR-FUNCTION-APP --resource-group YOUR_RG \
  --identities /subscriptions/.../resourceGroups/YOUR_RG/providers/Microsoft.ManagedIdentity/userAssignedIdentities/genie-connect-bot-identity

# App settings — secrets here are encrypted at rest by Azure, same trust
# model as Cloudflare's `wrangler secret put` from the GitHub Pages setup
az functionapp config appsettings set --name YOUR-FUNCTION-APP --resource-group YOUR_RG --settings \
  MicrosoftAppType=UserAssignedMsi \
  MicrosoftAppId="<clientId from step 1>" \
  MicrosoftAppTenantId="<tenantId from step 1>" \
  GENIE_API_KEY="..." \
  IAM_API_TOKEN="..." \
  GENIE_ID="gin-AbMAK4r6-rXgonW-CD" \
  DATA_CENTER="www.workato.com" \
  BOT_STATE_CONTAINER="botstate"
```
`BOT_STATE_STORAGE_CONNECTION` can be left unset — `state.js` falls back to
`AzureWebJobsStorage`, which Azure sets automatically for the same storage
account you passed to `az functionapp create`.

Deploy the code:
```bash
cd teams-bot
npm install
func azure functionapp publish YOUR-FUNCTION-APP
```

Go back to the Azure Bot resource's **Configuration** and set the
messaging endpoint to `https://YOUR-FUNCTION-APP.azurewebsites.net/api/messages`.

## 4. Package and install the Teams app

1. In `teams-manifest/manifest.json`:
   - `id` — generate a fresh GUID for this (any GUID generator; it just
     needs to be unique, it's the Teams app's own ID, separate from the bot).
   - `botId` — must exactly equal the managed identity's **clientId** from
     step 1 (same value as `MicrosoftAppId` above).
   - `developer` / `name` / `description` — fill in for your org; swap
     `color.png` / `outline.png` for real artwork if you want (current ones
     are plain placeholders).
2. Zip the three files together (`manifest.json`, `color.png`, `outline.png`
   — flat, no subfolder) into `genie-connect-teams-app.zip`.
3. In Teams: **Apps → Manage your apps → Upload an app → Upload a custom app**,
   select the zip. (Your tenant needs custom/sideloaded apps enabled —
   if that's locked down, go through **Teams admin center → Teams apps →
   Manage apps** to publish it org-wide instead.)

## Local testing

Managed Identity only works when actually running in Azure (it authenticates
via Azure's instance metadata service, which doesn't exist on your machine).
For local testing, use a throwaway **SingleTenant** app registration with a
client secret instead — `local.settings.json.example` is already set up for
that path. Copy it to `local.settings.json`, fill in a test app's
`MicrosoftAppId`/`MicrosoftAppPassword`/`MicrosoftAppTenantId`, your real
`GENIE_API_KEY`/`IAM_API_TOKEN`, then:

```bash
npm install -g azure-functions-core-tools@4
npm install
func start
```

Test it with the [Bot Framework Emulator](https://github.com/microsoft/BotFramework-Emulator)
pointed at `http://localhost:7071/api/messages`, using the same test app's
ID/password. This exercises the Genie calling logic end-to-end without
needing a real Teams install — `TeamsInfo.getMember` will fail outside a
real Teams context, so for this kind of local test, temporarily hardcode a
known Workato user ID in `bot.js`'s `resolveIdentity` to bypass it, then
revert before deploying.

## Security notes

- Production auth has **no stored secret at all** — Managed Identity is
  authenticated by Azure itself, nothing to leak or rotate on that front.
- `GENIE_API_KEY` / `IAM_API_TOKEN` live only in Function App settings
  (encrypted by Azure), never in code, never sent to Teams or the client.
- The local-testing app registration's client secret is a real credential —
  keep `local.settings.json` out of git (not included in this zip by
  default; add it to `.gitignore` if you turn this folder into its own repo)
  and delete that throwaway app registration once you're done testing.
- **Rotate `GENIE_API_KEY` / `IAM_API_TOKEN` if either was ever pasted into
  a chat, ticket, or document** — treat anything that left your machine as
  seen by more people than intended.
