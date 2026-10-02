/*!
 * HTTP trigger: the bot's messaging endpoint (what you put in the Azure Bot
 * resource and what Teams/Bot Framework Service POSTs each Activity to).
 *
 * IMPORTANT: CloudAdapter.process(req, res, logic) expects a classic
 * Express/Restify-style req (.body/.headers/.method) and res
 * (.status()/.header()/.send()/.end() as callable methods) — confirmed by
 * inspecting botbuilder@4.23.3's own zod schemas. The Azure Functions v4
 * programming model instead gives you a Fetch-style Request and expects a
 * plain HttpResponseInit back, which does NOT satisfy that shape. The two
 * small helpers below bridge between them. If you later upgrade botbuilder
 * and something changes here, this is the first place to check.
 */
const { app } = require('@azure/functions');
const { CloudAdapter, ConfigurationBotFrameworkAuthentication } = require('botbuilder');
const { GenieConnectBot } = require('../bot');

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('Sorry, something went wrong processing that.');
};

const bot = new GenieConnectBot();

async function toWebRequest(request) {
  const headers = {};
  for (const [key, value] of request.headers) headers[key] = value;
  let body;
  try {
    body = await request.json();
  } catch (e) {
    body = undefined;
  }
  return { body, headers, method: request.method };
}

function createWebResponse() {
  const state = { status: 200, headers: {}, body: undefined };
  return {
    socket: undefined,
    status(code) { state.status = code; return this; },
    header(name, value) { state.headers[name] = value; return this; },
    send(body) { state.body = body; return this; },
    end() { return this; },
    _state: state,
  };
}

app.http('messages', {
  methods: ['POST'],
  authLevel: 'anonymous', // Bot Framework does its own JWT validation inside adapter.process
  route: 'api/messages',
  handler: async (request, context) => {
    const webReq = await toWebRequest(request);
    const webRes = createWebResponse();

    await adapter.process(webReq, webRes, async (turnContext) => {
      await bot.run(turnContext);
    });

    return {
      status: webRes._state.status,
      headers: webRes._state.headers,
      jsonBody: webRes._state.body,
    };
  },
});
