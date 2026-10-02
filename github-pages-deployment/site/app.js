/*!
 * Genie Connect — static client logic.
 * No secrets live in this file. Every Workato call goes through your
 * Cloudflare Worker (see ../worker/worker.js), identified only by the
 * worker's own base URL, which is not sensitive.
 * -----------------------------------------------------------------------
 * FILL THESE IN from your Microsoft Entra ID App Registration
 * (Platform type: Single-page application — no client secret involved):
 */
const MSAL_CONFIG = {
  clientId: 'YOUR-ENTRA-APP-CLIENT-ID',
  tenantId: 'YOUR-ENTRA-TENANT-ID', // or 'common' / 'organizations'
};
/* The redirect URI you register in Entra must exactly match where this
 * page is served from, e.g. https://yourname.github.io/genie-connect/
 * This file computes it automatically from the current page location.
 *
 * If you're also wiring up the Teams tab (see ../teams-tab-manifest/),
 * register TWO redirect URIs in Entra — this page's own URL AND
 * .../teams-auth-end.html — and copy this same clientId/tenantId into
 * teams-auth-start.html and teams-auth-end.html. All three must match. */
// -----------------------------------------------------------------------

(function () {
  var $ = function (id) { return document.getElementById(id); };
  var LS = {
    workerBaseUrl: 'gc_worker_base_url',
    genieId: 'gc_genie_id',
    dataCenter: 'gc_data_center',
    idpUserId: 'gc_idp_user_id',
  };

  var els = {
    chatTitle: $('chatTitle'),
    chatIdentityLabel: $('chatIdentityLabel'),
    chatBody: $('chatBody'),
    emptyState: $('emptyState'),
    composerForm: $('composerForm'),
    messageInput: $('messageInput'),
    sendBtn: $('sendBtn'),
    streamState: $('streamState'),
    alertSlot: $('alertSlot'),
    resetBtn: $('resetBtn'),
    workerBaseUrl: $('workerBaseUrl'),
    genieId: $('genieId'),
    dataCenter: $('dataCenter'),
    idpUserId: $('idpUserId'),
    rememberId: $('rememberId'),
    railWho: $('railWho'),
    railStatus: $('railStatus'),
    identityWho: $('identityWho'),
    identitySourceBadge: $('identitySourceBadge'),
    signInBtn: $('signInBtn'),
    signOutBtn: $('signOutBtn'),
    lookupToggleBtn: $('lookupToggleBtn'),
    lookupPanel: $('lookupPanel'),
    lookupEmail: $('lookupEmail'),
    lookupSubmitBtn: $('lookupSubmitBtn'),
    lookupSubmitLabel: $('lookupSubmitLabel'),
    lookupResults: $('lookupResults'),
  };

  var state = { conversationId: null, sending: false, identitySource: 'none' };

  // ---------------------------------------------------------------------
  // Teams tab detection — microsoftTeams.app.initialize() only resolves
  // when this page is actually running inside a Teams iframe, so a
  // rejection here just means "normal browser tab", not an error.
  // ---------------------------------------------------------------------
  var inTeams = false;
  var teamsReady = (typeof microsoftTeams === 'undefined')
    ? Promise.resolve()
    : microsoftTeams.app.initialize().then(function () { inTeams = true; }).catch(function () {});

  // ---------------------------------------------------------------------
  // Tab navigation
  // ---------------------------------------------------------------------
  var navBtns = { chat: $('navChatBtn'), settings: $('navSettingsBtn') };
  var panels = { chat: $('chatPanel'), settings: $('settingsPanel') };
  function showTab(tab) {
    Object.keys(panels).forEach(function (k) {
      panels[k].classList.toggle('active', k === tab);
      navBtns[k].classList.toggle('active', k === tab);
    });
  }
  navBtns.chat.addEventListener('click', function () { showTab('chat'); });
  navBtns.settings.addEventListener('click', function () { showTab('settings'); });

  // ---------------------------------------------------------------------
  // Settings persistence (localStorage only — this browser, nowhere else)
  // ---------------------------------------------------------------------
  function loadSettings() {
    els.workerBaseUrl.value = localStorage.getItem(LS.workerBaseUrl) || '';
    els.genieId.value = localStorage.getItem(LS.genieId) || '';
    els.dataCenter.value = localStorage.getItem(LS.dataCenter) || '';
    var savedId = localStorage.getItem(LS.idpUserId);
    if (savedId) {
      els.idpUserId.value = savedId;
      els.rememberId.checked = true;
      state.identitySource = 'remembered';
    }
  }
  [['workerBaseUrl', LS.workerBaseUrl], ['genieId', LS.genieId], ['dataCenter', LS.dataCenter]].forEach(function (pair) {
    els[pair[0]].addEventListener('input', function () { localStorage.setItem(pair[1], els[pair[0]].value.trim()); });
  });
  els.idpUserId.addEventListener('input', function () {
    state.identitySource = 'manual';
    if (els.rememberId.checked) localStorage.setItem(LS.idpUserId, els.idpUserId.value.trim());
    updateIdentityDisplay();
  });
  els.rememberId.addEventListener('change', function () {
    if (els.rememberId.checked) localStorage.setItem(LS.idpUserId, els.idpUserId.value.trim());
    else localStorage.removeItem(LS.idpUserId);
  });
  loadSettings();

  function workerApi(path) {
    var base = (els.workerBaseUrl.value || '').replace(/\/$/, '');
    return base + path;
  }

  function updateIdentityDisplay() {
    var id = els.idpUserId.value.trim();
    var label = id ? id : 'not set';
    els.chatIdentityLabel.textContent = label;
    els.railWho.textContent = id ? id : 'Not signed in';
    els.railStatus.textContent = state.identitySource === 'sso' ? 'via Microsoft sign-in'
      : state.identitySource === 'remembered' ? 'remembered on this device'
      : state.identitySource === 'manual' ? 'entered manually'
      : '—';
    els.identityWho.innerHTML = id ? '<strong>' + escapeHtml(id) + '</strong>' : '<strong>Not signed in</strong>';
    els.identitySourceBadge.textContent = state.identitySource;
    els.identitySourceBadge.className = 'badge' + (state.identitySource === 'none' ? ' muted' : '');
  }
  updateIdentityDisplay();

  function escapeHtml(str) {
    var d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
  }

  // ---------------------------------------------------------------------
  // Microsoft Entra ID sign-in (MSAL.js — fully client-side)
  // ---------------------------------------------------------------------
  var msalInstance = null;
  var msalReady = (async function initMsal() {
    if (typeof msal === 'undefined') return; // script failed to load (offline, blocked, etc.)
    if (MSAL_CONFIG.clientId === 'YOUR-ENTRA-APP-CLIENT-ID') return; // not configured yet

    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: MSAL_CONFIG.clientId,
        authority: 'https://login.microsoftonline.com/' + MSAL_CONFIG.tenantId,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: {
        // localStorage (not sessionStorage) so the signed-in account survives
        // across visits — this is what lets the user skip signing in again.
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
      },
    });
    await msalInstance.initialize();
    await msalInstance.handleRedirectPromise().catch(function () {});

    var accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
      await onSignedIn(accounts[0]);
    }
  })();

  async function onSignedIn(account) {
    var email = account.username || (account.idTokenClaims && account.idTokenClaims.email);
    if (!email) return;
    els.railStatus.textContent = 'resolving user ID…';
    // Auto-resolve the Workato user ID from the SSO email, so the user never types it.
    try {
      await lookupUserByEmail(email, { auto: true });
    } catch (err) {
      showAlert('Signed in as ' + email + ', but could not resolve a Workato user ID automatically: ' + err.message);
    }
  }

  if (els.signInBtn) {
    els.signInBtn.addEventListener('click', async function () {
      await msalReady;
      await teamsReady;
      if (!msalInstance) {
        showAlert('Microsoft sign-in is not configured yet — set clientId/tenantId at the top of app.js.');
        return;
      }
      try {
        if (inTeams) {
          // A plain loginPopup() call is blocked (or silently fails) from
          // inside a Teams tab's iframe in most browsers. Teams provides
          // its own popup window for this instead — see
          // teams-auth-start.html / teams-auth-end.html.
          await microsoftTeams.authentication.authenticate({
            url: window.location.origin + window.location.pathname.replace(/[^/]*$/, '') + 'teams-auth-start.html',
            width: 600,
            height: 535,
          });
          // The popup ran the redirect flow and cached the result in the
          // same-origin localStorage MSAL cache — re-read it here.
          var accounts = msalInstance.getAllAccounts();
          if (accounts.length === 0) throw new Error('Sign-in completed but no account was found.');
          msalInstance.setActiveAccount(accounts[0]);
          state.identitySource = 'sso';
          await onSignedIn(accounts[0]);
        } else {
          var result = await msalInstance.loginPopup({ scopes: ['openid', 'profile', 'email', 'User.Read'] });
          msalInstance.setActiveAccount(result.account);
          state.identitySource = 'sso';
          await onSignedIn(result.account);
        }
      } catch (err) {
        showAlert('Sign-in failed: ' + err.message);
      }
    });
  }
  if (els.signOutBtn) {
    els.signOutBtn.addEventListener('click', async function () {
      await msalReady;
      els.idpUserId.value = '';
      els.rememberId.checked = false;
      localStorage.removeItem(LS.idpUserId);
      state.identitySource = 'none';
      updateIdentityDisplay();
      if (msalInstance && !inTeams) {
        // logoutPopup() opens its own popup from the iframe, which Teams
        // tabs generally block — in Teams, just drop the active account
        // and leave the shared localStorage cache for a full logout done
        // from the regular browser tab instead.
        var account = msalInstance.getActiveAccount();
        if (account) { try { await msalInstance.logoutPopup({ account: account }); } catch (e) {} }
      } else if (msalInstance) {
        msalInstance.setActiveAccount(null);
      }
    });
  }

  // ---------------------------------------------------------------------
  // Email → user ID lookup (via the worker's /api/iam/users)
  // ---------------------------------------------------------------------
  els.lookupToggleBtn.addEventListener('click', function () {
    var visible = els.lookupPanel.style.display !== 'none';
    els.lookupPanel.style.display = visible ? 'none' : 'block';
  });

  function isValidEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); }

  async function lookupUserByEmail(emailArg, opts) {
    opts = opts || {};
    var email = emailArg || els.lookupEmail.value.trim();
    var dataCenter = (els.dataCenter.value || '').trim();
    if (!dataCenter) { showAlert('Set a data center in Settings (e.g. www.workato.com) before looking up by email.'); return; }
    if (!email || !isValidEmail(email)) { if (!opts.auto) showAlert('Enter a valid email address.'); return; }
    if (!els.workerBaseUrl.value.trim()) { showAlert('Set the worker base URL in Settings first.'); return; }

    if (!opts.auto) {
      els.lookupSubmitBtn.disabled = true;
      els.lookupSubmitLabel.textContent = 'Looking up…';
    }
    try {
      var res = await fetch(workerApi('/api/iam/users?email=' + encodeURIComponent(email) + '&dataCenter=' + encodeURIComponent(dataCenter)));
      var data = await res.json().catch(function () { return {}; });
      if (!res.ok) throw new Error(data.error || ('Lookup failed (' + res.status + ')'));
      var matches = data.data || [];
      if (matches.length === 0) {
        els.lookupResults.innerHTML = '<p class="hint">No user found for <strong>' + escapeHtml(email) + '</strong>.</p>';
        return;
      }
      if (matches.length === 1) {
        applyResolvedUser(matches[0], opts.auto ? 'sso' : 'manual');
        return;
      }
      renderLookupResults(matches, opts.auto ? 'sso' : 'manual');
    } finally {
      if (!opts.auto) {
        els.lookupSubmitBtn.disabled = false;
        els.lookupSubmitLabel.textContent = 'Find user ID';
      }
    }
  }

  function applyResolvedUser(user, source) {
    els.idpUserId.value = user.id;
    state.identitySource = source;
    localStorage.setItem(LS.idpUserId, user.id); // SSO/lookup-resolved IDs are remembered automatically
    els.rememberId.checked = true;
    els.lookupPanel.style.display = 'none';
    updateIdentityDisplay();
    clearAlert();
  }

  function renderLookupResults(matches, source) {
    els.lookupResults.innerHTML = matches.map(function (u, i) {
      return '<div class="user-result" data-idx="' + i + '">' +
        '<div class="who"><strong>' + escapeHtml(u.name || u.email) + '</strong><span>' + escapeHtml(u.id) + '</span></div>' +
        '<button class="btn btn-secondary use-id-btn" type="button" data-id="' + escapeHtml(u.id) + '" style="font-size:.7rem; padding:.4rem .65rem;">Use this ID</button>' +
        '</div>';
    }).join('');
    els.lookupResults.querySelectorAll('.use-id-btn').forEach(function (btn, i) {
      btn.addEventListener('click', function () { applyResolvedUser(matches[i], source); });
    });
  }

  els.lookupSubmitBtn.addEventListener('click', function () { lookupUserByEmail(); });
  els.lookupEmail.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); lookupUserByEmail(); } });

  // ---------------------------------------------------------------------
  // Chat
  // ---------------------------------------------------------------------
  function showAlert(text) { els.alertSlot.innerHTML = '<div class="alert">' + escapeHtml(text) + '</div>'; }
  function clearAlert() { els.alertSlot.innerHTML = ''; }

  function renderMessage(role, text) {
    var empty = $('emptyState');
    if (empty) empty.remove();
    var wrap = document.createElement('div');
    wrap.className = 'msg ' + role;
    wrap.innerHTML = '<div class="msg-avatar">' + (role === 'user' ? 'YOU' : 'AI') + '</div><div><div class="msg-bubble"></div></div>';
    els.chatBody.appendChild(wrap);
    var bubble = wrap.querySelector('.msg-bubble');
    bubble.textContent = text;
    els.chatBody.scrollTop = els.chatBody.scrollHeight;
    return bubble;
  }
  function renderTyping() {
    var wrap = document.createElement('div');
    wrap.className = 'msg assistant';
    wrap.id = 'typingIndicator';
    wrap.innerHTML = '<div class="msg-avatar">AI</div><div class="msg-bubble"><span class="typing-dots"><span></span><span></span><span></span></span></div>';
    els.chatBody.appendChild(wrap);
    els.chatBody.scrollTop = els.chatBody.scrollHeight;
  }
  function removeTyping() { var el = $('typingIndicator'); if (el) el.remove(); }

  function validate() {
    if (!els.workerBaseUrl.value.trim()) { showAlert('Set the worker base URL in Settings first.'); return null; }
    var userId = els.idpUserId.value.trim();
    if (!userId) { showAlert('Set an end user ID in Settings (sign in, look up by email, or enter it directly).'); return null; }
    return { genieId: els.genieId.value.trim(), userId: userId };
  }

  async function createConversation(genieId, userId) {
    var res = await fetch(workerApi('/api/genie/conversations'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ genieId: genieId || undefined, idpUserId: userId }),
    });
    var data = await res.json().catch(function () { return {}; });
    if (!res.ok) throw new Error(data.error || ('Failed to create conversation (' + res.status + ')'));
    var conversationId = data.conversation_id || (data.result && data.result.conversation_id);
    if (!conversationId) throw new Error('Response did not include a conversation_id.');
    return conversationId;
  }

  async function sendMessageStream(genieId, userId, conversationId, message, onChunk) {
    var res = await fetch(workerApi('/api/genie/conversations/' + conversationId + '/messages'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ genieId: genieId || undefined, idpUserId: userId, message: message }),
    });
    if (!res.ok || !res.body) {
      var errText = await res.text().catch(function () { return ''; });
      throw new Error(errText || ('Message request failed (' + res.status + ' ' + res.statusText + ')'));
    }
    var reader = res.body.getReader();
    var decoder = new TextDecoder();
    var buffer = '', full = '';
    while (true) {
      var out = await reader.read();
      if (out.done) break;
      buffer += decoder.decode(out.value, { stream: true });
      var lines = buffer.split('\n');
      buffer = lines.pop();
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line || !line.startsWith('data:')) continue; // only "data:" lines carry content
        var payload = line.slice(5).trim();
        if (!payload || payload === '[DONE]') continue;
        var piece = '';
        try {
          var parsed = JSON.parse(payload);
          piece = parsed.delta || parsed.content || parsed.message || parsed.text ||
            (parsed.result && (parsed.result.delta || parsed.result.content || parsed.result.message || parsed.result.text)) || '';
        } catch (e) { /* ignore non-JSON metadata lines */ }
        if (piece) { full += piece; onChunk(full); }
      }
    }
    return full;
  }

  async function handleSend(message) {
    if (!message || state.sending) return;
    var creds = validate();
    if (!creds) return;

    clearAlert();
    state.sending = true;
    els.sendBtn.disabled = true;
    els.streamState.textContent = 'Connecting…';
    renderMessage('user', message);
    els.messageInput.value = '';
    autoGrow();

    try {
      if (!state.conversationId) {
        state.conversationId = await createConversation(creds.genieId, creds.userId);
      }
      renderTyping();
      els.streamState.textContent = 'Streaming response…';
      var bubble = null;
      await sendMessageStream(creds.genieId, creds.userId, state.conversationId, message, function (fullText) {
        removeTyping();
        if (!bubble) bubble = renderMessage('assistant', '');
        bubble.textContent = fullText;
        els.chatBody.scrollTop = els.chatBody.scrollHeight;
      });
      removeTyping();
      if (!bubble) renderMessage('assistant', '(No content returned. Check that the genie is running.)');
      els.streamState.textContent = 'Idle';
    } catch (err) {
      removeTyping();
      els.streamState.textContent = 'Error';
      showAlert(err.message + ' — check your worker URL, user ID, and that the genie is running.');
    } finally {
      state.sending = false;
      els.sendBtn.disabled = false;
    }
  }

  els.composerForm.addEventListener('submit', function (e) { e.preventDefault(); handleSend(els.messageInput.value.trim()); });
  els.messageInput.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(els.messageInput.value.trim()); }
  });
  function autoGrow() {
    els.messageInput.style.height = 'auto';
    els.messageInput.style.height = Math.min(els.messageInput.scrollHeight, 120) + 'px';
  }
  els.messageInput.addEventListener('input', autoGrow);

  els.resetBtn.addEventListener('click', function () {
    state.conversationId = null;
    els.streamState.textContent = 'Idle';
    clearAlert();
    els.chatBody.innerHTML =
      '<div class="empty-state" id="emptyState">' +
      '<div class="glyph"><svg viewBox="0 0 24 24" fill="none"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg></div>' +
      '<h3>No conversation yet</h3><p>Set your identity in Settings (or sign in with Microsoft), then send a message below.</p>' +
      '</div>';
  });
})();
