/*!
 * Thin client for the Workato Genie headless API and IAM users endpoint.
 * Mirrors the logic already proven out in server.js / worker.js — same
 * endpoints, same headers, same response shape.
 */

function genieBase(genieId) {
  return `https://genie-api.workato.com/api/v1/genies/${genieId}/chat`;
}

async function createConversation({ genieApiKey, genieId, idpUserId }) {
  const res = await fetch(`${genieBase(genieId)}/conversations`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${genieApiKey}`,
      'X-IDP-User-ID': idpUserId,
      'Content-Type': 'application/json',
    },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Failed to create conversation (${res.status})`);
  const conversationId = data.conversation_id || data.result?.conversation_id;
  if (!conversationId) throw new Error('Response did not include a conversation_id.');
  return conversationId;
}

/**
 * Sends a message and collects the FULL streamed reply into one string.
 * Teams doesn't do token-by-token streaming the way the web widget does,
 * so we buffer server-side and send one complete message back.
 */
async function sendMessageAndCollect({ genieApiKey, genieId, idpUserId, conversationId, message }) {
  const res = await fetch(`${genieBase(genieId)}/conversations/${conversationId}/messages`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${genieApiKey}`,
      'X-IDP-User-ID': idpUserId,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ message, stream: true }),
  });
  if (!res.ok || !res.body) {
    const errText = await res.text().catch(() => '');
    throw new Error(errText || `Message request failed (${res.status} ${res.statusText})`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let full = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop();
    for (const raw of lines) {
      const line = raw.trim();
      if (!line || !line.startsWith('data:')) continue; // only "data:" lines carry content
      const payload = line.slice(5).trim();
      if (!payload || payload === '[DONE]') continue;
      let piece = '';
      try {
        const parsed = JSON.parse(payload);
        piece =
          parsed.delta || parsed.content || parsed.message || parsed.text ||
          parsed.result?.delta || parsed.result?.content || parsed.result?.message || parsed.result?.text || '';
      } catch (e) { /* ignore non-JSON metadata lines */ }
      if (piece) full += piece;
    }
  }
  return full;
}

async function lookupUserByEmail({ iamApiToken, dataCenter, email }) {
  const host = dataCenter.replace(/^https?:\/\//, '').replace(/\/$/, '');
  const url = `https://${host}/api/iam/users?${encodeURIComponent('emails[]')}=${encodeURIComponent(email)}`;
  const res = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${iamApiToken}`, 'Content-Type': 'application/json' },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `IAM lookup failed (${res.status})`);
  return data.data || [];
}

module.exports = { createConversation, sendMessageAndCollect, lookupUserByEmail };
