/*!
 * The bot itself. One Genie conversation per Teams user (1:1 chat), identity
 * resolved once (via the Teams member's email → Workato IAM lookup) and
 * cached in blob-backed user state so later messages skip the lookup.
 */
const { TeamsActivityHandler, TeamsInfo } = require('botbuilder');
const { userState, identityAccessor } = require('./state');
const { createConversation, sendMessageAndCollect, lookupUserByEmail } = require('./genieClient');

const GENIE_API_KEY = process.env.GENIE_API_KEY;
const IAM_API_TOKEN = process.env.IAM_API_TOKEN;
const GENIE_ID = process.env.GENIE_ID;
const DATA_CENTER = process.env.DATA_CENTER;

class GenieConnectBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim();

      if (/^\/reset\b/i.test(text)) {
        await identityAccessor.set(context, { ...(await this.getIdentity(context)), genieConversationId: null });
        await userState.saveChanges(context);
        await context.sendActivity('Session reset — your next message starts a new conversation.');
        await next();
        return;
      }

      if (!text) { await next(); return; }

      await context.sendActivity({ type: 'typing' });

      try {
        const identity = await this.resolveIdentity(context);
        if (!identity.idpUserId) {
          await context.sendActivity(
            "I couldn't match your account to a Workato user ID. Ask whoever manages Genie Connect to confirm you're provisioned in Workspace settings, then message me again.",
          );
          await next();
          return;
        }

        if (!identity.genieConversationId) {
          identity.genieConversationId = await createConversation({
            genieApiKey: GENIE_API_KEY,
            genieId: GENIE_ID,
            idpUserId: identity.idpUserId,
          });
          await identityAccessor.set(context, identity);
          await userState.saveChanges(context);
        }

        const reply = await sendMessageAndCollect({
          genieApiKey: GENIE_API_KEY,
          genieId: GENIE_ID,
          idpUserId: identity.idpUserId,
          conversationId: identity.genieConversationId,
          message: text,
        });

        await context.sendActivity(reply || '(No content returned. Check that the genie is running.)');
      } catch (err) {
        await context.sendActivity(`Something went wrong talking to the genie: ${err.message}`);
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity("Hi! I'm connected to the Genie headless API. Send me a message to get started, or `/reset` to start a fresh conversation.");
        }
      }
      await next();
    });
  }

  async getIdentity(context) {
    return (await identityAccessor.get(context, () => ({}))) || {};
  }

  /** Resolves (and caches) the Workato idpUserId for the current Teams user. */
  async resolveIdentity(context) {
    const identity = await this.getIdentity(context);
    if (identity.idpUserId) return identity;

    let email;
    try {
      const member = await TeamsInfo.getMember(context, context.activity.from.id);
      email = member.email || member.userPrincipalName;
    } catch (err) {
      throw new Error(`could not read your Teams profile (${err.message})`);
    }
    if (!email) throw new Error('your Teams profile has no email/UPN available');

    const matches = await lookupUserByEmail({ iamApiToken: IAM_API_TOKEN, dataCenter: DATA_CENTER, email });
    if (matches.length === 0) return identity; // leave idpUserId unset — caller handles the message
    // If several Workato users share an email domain match, take the first;
    // tighten this if your IAM lookup can return ambiguous results.
    identity.idpUserId = matches[0].id;
    identity.email = email;
    await identityAccessor.set(context, identity);
    await userState.saveChanges(context);
    return identity;
  }
}

module.exports = { GenieConnectBot };
