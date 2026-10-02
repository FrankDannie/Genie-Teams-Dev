/*!
 * Per-Teams-user state, persisted in Blob Storage so it survives between
 * messages (Azure Functions itself is stateless between invocations).
 * Reuses the same storage account the Function App already needs for
 * AzureWebJobsStorage — just a different container.
 */
const { UserState } = require('botbuilder');
const { BlobsStorage } = require('botbuilder-azure-blobs');

const storage = new BlobsStorage(
  process.env.BOT_STATE_STORAGE_CONNECTION || process.env.AzureWebJobsStorage,
  process.env.BOT_STATE_CONTAINER || 'botstate',
);

const userState = new UserState(storage);

// One property bag per Teams user: { idpUserId, email, genieConversationId }
const identityAccessor = userState.createProperty('genieIdentity');

module.exports = { userState, identityAccessor };
