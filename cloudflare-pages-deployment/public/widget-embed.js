/*!
 * Genie Connect — embeddable chat widget
 * -----------------------------------------------------------------------
 * Drop this on any page with a single <script> tag:
 *
 *   <script
 *     src="https://YOUR-GENIE-CONNECT-HOST/widget-embed.js"
 *     data-base-url="https://YOUR-GENIE-CONNECT-HOST"
 *     data-genie-id="gin-AbMAK4r6-rXgonW-CD"
 *     data-interface-name="Lenovo QA Genie"
 *     data-idp-user-id=""
 *   ></script>
 *
 * - data-base-url   : where this project's server.js is running. Required
 *                      whenever the widget is embedded on a DIFFERENT origin
 *                      than the backend (which it will be, e.g. embedded
 *                      inside your app's own page). Leave empty only if the
 *                      widget script itself is served from the same origin
 *                      as the backend.
 * - data-genie-id    : defaults to the server's GENIE_ID if omitted.
 * - data-idp-user-id : the signed-in end user's Workato user ID. If you
 *                      already resolve this via your app's own Entra/SSO
 *                      session, pass it here so the visitor is never asked
 *                      to type anything. If omitted, the widget shows a
 *                      one-line prompt asking for it once, then remembers
 *                      it in-memory for the rest of the page session.
 *
 * You can also set/update the user id at any point after the SSO session
 * resolves, from your own page's JS:
 *
 *   window.GenieWidget.setUser("usr_2f8b1c...");
 *   window.GenieWidget.open();
 *   window.GenieWidget.close();
 *
 * Note on Microsoft Entra's "My Apps" portal itself: that portal only
 * launches/links to registered app URLs — it has no mechanism to host a
 * third-party widget inside its own chrome. This script is meant to be
 * included on a page your own app serves (the app registered in Entra),
 * not injected into myapplications.microsoft.com.
 * -----------------------------------------------------------------------
 */
(function () {
  var thisScript = document.currentScript;
  var cfg = {
    baseUrl: (thisScript && thisScript.getAttribute('data-base-url')) || '',
    genieId: (thisScript && thisScript.getAttribute('data-genie-id')) || '',
    dataCenter: (thisScript && thisScript.getAttribute('data-data-center')) || '',
    interfaceName: (thisScript && thisScript.getAttribute('data-interface-name')) || 'Genie Connect',
    idpUserId: (thisScript && thisScript.getAttribute('data-idp-user-id')) || '',
    // A host page that already knows the signed-in user's email can pass it
    // directly instead of relying on the ?email= query param below.
    userEmail: (thisScript && thisScript.getAttribute('data-user-email')) || '',
    // Optional best-effort silent sign-in (e.g. for a SharePoint-embedded page
    // where the visitor is already signed into Microsoft 365). Only attempted
    // when both of these are set — otherwise the widget never loads MSAL at
    // all, keeping the plain drop-in case lightweight.
    ssoClientId: (thisScript && thisScript.getAttribute('data-sso-client-id')) || '',
    ssoTenantId: (thisScript && thisScript.getAttribute('data-sso-tenant-id')) || '',
  };

  var ROOT_ID = 'genie-widget-root';
  if (document.getElementById(ROOT_ID)) return; // already injected

  // ---------------------------------------------------------------------
  // Styles — scoped under #genie-widget-root so nothing leaks into, or is
  // affected by, the host page's own CSS.
  // ---------------------------------------------------------------------
  var fontLink = document.createElement('link');
  fontLink.rel = 'stylesheet';
  fontLink.href = 'https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap';
  document.head.appendChild(fontLink);

  var style = document.createElement('style');
  style.textContent = [
    '#' + ROOT_ID + '{',
    '  --gw-teal:#0D9488; --gw-teal-mid:#0F766E; --gw-teal-soft:#CCFBF1;',
    '  --gw-card:#FFFFFF; --gw-surface-2:#F5F7FA; --gw-border:#DDE3EA; --gw-border-light:#E7EBF0;',
    '  --gw-text:#101828; --gw-text-secondary:#3A4150; --gw-text-muted:#8993A3;',
    '  --gw-success:#059669; --gw-danger:#DC2626; --gw-warning:#B45309;',
    '  --gw-sans:"Poppins",sans-serif; --gw-mono:"JetBrains Mono",monospace;',
    '  font-family:var(--gw-sans); color:var(--gw-text); box-sizing:border-box;',
    '}',
    '#' + ROOT_ID + ' *{box-sizing:border-box;}',
    '#' + ROOT_ID + ' .gw-launcher{position:fixed; right:1.5rem; bottom:1.5rem; z-index:2147483000; width:54px; height:54px; border-radius:50%; border:none; cursor:pointer; background:linear-gradient(135deg,#0D9488,#0F766E); color:#fff; display:flex; align-items:center; justify-content:center; box-shadow:0 8px 24px rgba(15,118,110,0.35); transition:transform .15s ease;}',
    '#' + ROOT_ID + ' .gw-launcher:hover{transform:translateY(-2px) scale(1.03);}',
    '#' + ROOT_ID + ' .gw-launcher svg{width:24px; height:24px;}',
    '#' + ROOT_ID + ' .gw-launcher .gw-close-icon{display:none;}',
    '#' + ROOT_ID + ' .gw-launcher.open .gw-chat-icon{display:none;}',
    '#' + ROOT_ID + ' .gw-launcher.open .gw-close-icon{display:block;}',
    '#' + ROOT_ID + ' .gw-ping{position:absolute; top:-2px; right:-2px; width:11px; height:11px; border-radius:50%; background:var(--gw-success); border:2px solid #fff; display:none;}',
    '#' + ROOT_ID + ' .gw-ping.show{display:block;}',
    '#' + ROOT_ID + ' .gw-panel{position:fixed; right:1.5rem; bottom:5.5rem; z-index:2147482999; width:380px; max-width:calc(100vw - 2.5rem); height:min(600px, calc(100vh - 8rem)); background:var(--gw-card); border:1px solid var(--gw-border-light); border-radius:0.75rem; box-shadow:0 20px 48px rgba(16,24,40,0.18); display:flex; flex-direction:column; overflow:hidden; transform:translateY(12px) scale(0.98); opacity:0; pointer-events:none; transition:transform .18s ease, opacity .18s ease;}',
    '#' + ROOT_ID + ' .gw-panel.open{transform:translateY(0) scale(1); opacity:1; pointer-events:auto;}',
    '#' + ROOT_ID + ' .gw-head{padding:1rem 1.1rem; border-bottom:1px solid var(--gw-border-light); display:flex; align-items:center; justify-content:space-between; gap:.75rem; background:linear-gradient(180deg, var(--gw-teal-soft), transparent); flex-shrink:0;}',
    '#' + ROOT_ID + ' .gw-head h4{margin:0; font-size:.88rem; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
    '#' + ROOT_ID + ' .gw-head-actions{display:flex; align-items:center; gap:.3rem; flex-shrink:0;}',
    '#' + ROOT_ID + ' .gw-icon-btn{background:transparent; border:1px solid transparent; color:var(--gw-text-muted); width:1.9rem; height:1.9rem; border-radius:.25rem; display:flex; align-items:center; justify-content:center; cursor:pointer; transition:all .15s ease;}',
    '#' + ROOT_ID + ' .gw-icon-btn:hover{color:var(--gw-teal-mid); background:rgba(13,148,136,.08);}',
    '#' + ROOT_ID + ' .gw-icon-btn svg{width:15px; height:15px;}',
    '#' + ROOT_ID + ' .gw-body{flex:1; overflow-y:auto; padding:1.1rem; display:flex; flex-direction:column; gap:1rem; background:var(--gw-surface-2);}',
    '#' + ROOT_ID + ' .gw-empty{margin:auto; text-align:center; max-width:280px; color:var(--gw-text-muted); display:flex; flex-direction:column; align-items:center; gap:.8rem;}',
    '#' + ROOT_ID + ' .gw-empty .gw-glyph{width:46px; height:46px; border-radius:12px; background:var(--gw-teal-soft); border:1px solid rgba(13,148,136,.25); display:flex; align-items:center; justify-content:center;}',
    '#' + ROOT_ID + ' .gw-empty .gw-glyph svg{width:21px; height:21px; color:var(--gw-teal-mid);}',
    '#' + ROOT_ID + ' .gw-empty h3{color:var(--gw-text-secondary); font-size:.86rem; font-weight:600; margin:0;}',
    '#' + ROOT_ID + ' .gw-empty p{font-size:.74rem; line-height:1.55; margin:0;}',
    '#' + ROOT_ID + ' .gw-id-field{display:flex; gap:.4rem; margin-top:.3rem; width:100%;}',
    '#' + ROOT_ID + ' .gw-id-field input{flex:1; min-width:0; background:#fff; border:1px solid var(--gw-border); color:var(--gw-text); font-family:var(--gw-mono); font-size:.76rem; padding:.5rem .6rem; border-radius:.375rem;}',
    '#' + ROOT_ID + ' .gw-id-field input:focus{outline:none; border-color:var(--gw-teal); box-shadow:0 0 0 3px rgba(13,148,136,.12);}',
    '#' + ROOT_ID + ' .gw-id-field button{flex-shrink:0; background:linear-gradient(135deg,#0D9488,#0F766E); color:#fff; border:none; border-radius:.375rem; font-size:.74rem; font-weight:600; padding:0 .8rem; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-msg{display:flex; gap:.6rem; max-width:88%;}',
    '#' + ROOT_ID + ' .gw-msg.user{align-self:flex-end; flex-direction:row-reverse;}',
    '#' + ROOT_ID + ' .gw-msg-avatar{flex-shrink:0; width:24px; height:24px; border-radius:7px; display:flex; align-items:center; justify-content:center; font-size:.6rem; font-weight:700; font-family:var(--gw-mono); margin-top:.1rem;}',
    '#' + ROOT_ID + ' .gw-msg.assistant .gw-msg-avatar{background:var(--gw-teal-soft); border:1px solid rgba(13,148,136,.3); color:var(--gw-teal-mid);}',
    '#' + ROOT_ID + ' .gw-msg.user .gw-msg-avatar{background:rgba(124,58,237,.1); border:1px solid rgba(124,58,237,.28); color:#6d28d9;}',
    '#' + ROOT_ID + ' .gw-msg-bubble{padding:.6rem .8rem; border-radius:.5rem; font-size:.82rem; line-height:1.55; color:var(--gw-text-secondary); white-space:pre-wrap; word-break:break-word;}',
    '#' + ROOT_ID + ' .gw-msg.assistant .gw-msg-bubble{background:#fff; border:1px solid var(--gw-border-light); border-top-left-radius:.2rem;}',
    '#' + ROOT_ID + ' .gw-msg.user .gw-msg-bubble{background:rgba(124,58,237,.08); border:1px solid rgba(124,58,237,.2); border-top-right-radius:.2rem; color:var(--gw-text);}',
    '#' + ROOT_ID + ' .gw-typing{display:inline-flex; gap:3px; align-items:center; padding:.2rem 0;}',
    '#' + ROOT_ID + ' .gw-typing span{width:5px; height:5px; border-radius:50%; background:var(--gw-teal-mid); animation:gw-bounce 1.2s infinite ease-in-out;}',
    '#' + ROOT_ID + ' .gw-typing span:nth-child(2){animation-delay:.15s;}',
    '#' + ROOT_ID + ' .gw-typing span:nth-child(3){animation-delay:.3s;}',
    '@keyframes gw-bounce{0%,80%,100%{transform:translateY(0); opacity:.4;} 40%{transform:translateY(-4px); opacity:1;}}',
    '#' + ROOT_ID + ' .gw-footer{padding:.85rem 1rem 1rem; border-top:1px solid var(--gw-border-light); flex-shrink:0; background:var(--gw-card);}',
    '#' + ROOT_ID + ' .gw-composer{display:flex; align-items:flex-end; gap:.5rem; background:var(--gw-surface-2); border:1px solid var(--gw-border); border-radius:.5rem; padding:.4rem .5rem; transition:border-color .15s ease;}',
    '#' + ROOT_ID + ' .gw-composer:focus-within{border-color:var(--gw-teal); box-shadow:0 0 0 3px rgba(13,148,136,.12);}',
    '#' + ROOT_ID + ' .gw-composer textarea{flex:1; border:none; background:transparent; resize:none; font-family:var(--gw-sans); font-size:.82rem; color:var(--gw-text); max-height:100px; padding:.35rem;}',
    '#' + ROOT_ID + ' .gw-composer textarea:focus{outline:none;}',
    '#' + ROOT_ID + ' .gw-send{width:2.1rem; height:2.1rem; border-radius:.375rem; flex-shrink:0; background:linear-gradient(135deg,#0D9488,#0F766E); border:none; color:#fff; display:flex; align-items:center; justify-content:center; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-send:disabled{opacity:.35; cursor:not-allowed;}',
    '#' + ROOT_ID + ' .gw-send svg{width:15px; height:15px;}',
    '#' + ROOT_ID + ' .gw-foot-row{display:flex; justify-content:space-between; align-items:center; margin-top:.5rem; font-size:.65rem; color:var(--gw-text-muted); font-family:var(--gw-mono);}',
    '#' + ROOT_ID + ' .gw-alert{display:flex; gap:.5rem; padding:.6rem .75rem; border-radius:.375rem; font-size:.74rem; line-height:1.5; margin-bottom:.75rem; border:1px solid rgba(220,38,38,.22); background:rgba(220,38,38,.06); color:#b91c1c;}',
  ].join('\n');
  document.head.appendChild(style);

  // ---------------------------------------------------------------------
  // Markup
  // ---------------------------------------------------------------------
  var root = document.createElement('div');
  root.id = ROOT_ID;
  root.innerHTML = [
    '<button class="gw-launcher" id="gwLauncher" type="button" aria-label="Open chat">',
    '  <span class="gw-ping" id="gwPing"></span>',
    '  <svg class="gw-chat-icon" viewBox="0 0 24 24" fill="none"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
    '  <svg class="gw-close-icon" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6 6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>',
    '</button>',
    '<div class="gw-panel" id="gwPanel">',
    '  <div class="gw-head">',
    '    <h4 id="gwTitle"></h4>',
    '    <div class="gw-head-actions">',
    '      <button class="gw-icon-btn" id="gwResetBtn" type="button" title="Reset session"><svg viewBox="0 0 24 24" fill="none"><path d="M3 12a9 9 0 1 0 3-6.7M3 4v5h5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg></button>',
    '      <button class="gw-icon-btn" id="gwCloseBtn" type="button" title="Close chat" aria-label="Close chat"><svg viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6 6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>',
    '    </div>',
    '  </div>',
    '  <div class="gw-body" id="gwBody"></div>',
    '  <div class="gw-footer">',
    '    <div id="gwAlertSlot"></div>',
    '    <form class="gw-composer" id="gwComposerForm">',
    '      <textarea id="gwInput" rows="1" placeholder="What can you help me with?" aria-label="Message"></textarea>',
    '      <button class="gw-send" id="gwSendBtn" type="submit" aria-label="Send message"><svg viewBox="0 0 24 24" fill="none"><path d="M22 2 11 13" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M22 2 15 22l-4-9-9-4 20-7Z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg></button>',
    '    </form>',
    '    <div class="gw-foot-row"><span id="gwStreamState">Idle</span></div>',
    '  </div>',
    '</div>',
  ].join('\n');
  document.body.appendChild(root);

  // ---------------------------------------------------------------------
  // Behavior
  // ---------------------------------------------------------------------
  var $ = function (id) { return document.getElementById(id); };
  var launcher = $('gwLauncher'), panel = $('gwPanel'), ping = $('gwPing');
  var body = $('gwBody'), title = $('gwTitle');
  var form = $('gwComposerForm'), input = $('gwInput'), sendBtn = $('gwSendBtn');
  var streamState = $('gwStreamState'), alertSlot = $('gwAlertSlot');

  var state = { conversationId: null, sending: false, idpUserId: cfg.idpUserId || '' };
  title.textContent = cfg.interfaceName;

  function api(path) { return (cfg.baseUrl || '') + path; }
  function isValidEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); }
  function getQueryEmail() {
    try {
      return new URLSearchParams(window.location.search).get('email') || '';
    } catch (e) { return ''; }
  }

  // Pick up server-configured defaults (genie ID / data center) when the
  // host page didn't set data-genie-id / data-data-center explicitly. This
  // is awaited below (as defaultsReady) before any auto-identify attempt —
  // resolveEmailToUserId needs cfg.dataCenter, and firing an auto-identify
  // attempt before this resolves silently fails the lookup (the "no data
  // center" warning is suppressed in silent mode), leaving the widget
  // sitting on the manual email prompt for no visible reason.
  var defaultsReady = (async function loadDefaults() {
    try {
      var res = await fetch(api('/api/health'));
      var data = await res.json();
      if (!cfg.genieId && data.defaultGenieId) cfg.genieId = data.defaultGenieId;
      if (!cfg.dataCenter && data.defaultDataCenter) cfg.dataCenter = data.defaultDataCenter;
    } catch (e) { /* backend unreachable — widget still works if data attrs were set explicitly */ }
  })();

  function setOpen(open) {
    panel.classList.toggle('open', open);
    launcher.classList.toggle('open', open);
    if (open) ping.classList.remove('show');
  }
  launcher.addEventListener('click', function () { setOpen(!panel.classList.contains('open')); });
  $('gwCloseBtn').addEventListener('click', function () { setOpen(false); });

  function escapeHtml(str) {
    var d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
  }

  function renderEmptyState(statusLine) {
    var needsId = !state.idpUserId;
    body.innerHTML =
      '<div class="gw-empty" id="gwEmpty">' +
      '  <div class="gw-glyph"><svg viewBox="0 0 24 24" fill="none"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg></div>' +
      (needsId
        ? '  <h3>Sign in to start</h3><p>Enter your email once to start chatting.</p>' +
          '  <div class="gw-id-field"><input id="gwIdInput" type="text" placeholder="you@company.com" spellcheck="false" /><button id="gwIdSubmit" type="button">Start</button></div>' +
          (statusLine ? '  <p style="margin-top:.5rem;font-size:.7rem;color:var(--gw-text-muted);">' + escapeHtml(statusLine) + '</p>' : '')
        : '  <h3>No conversation yet</h3><p>Send a message below to get started.</p>') +
      '</div>';
    if (needsId) {
      var go = function () {
        var v = $('gwIdInput').value.trim();
        if (!v) return;
        resolveEmailToUserId(v, { silent: false });
      };
      $('gwIdSubmit').addEventListener('click', go);
      $('gwIdInput').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); go(); } });
    }
  }
  renderEmptyState();

  // ---------------------------------------------------------------------
  // Email -> X-IDP-User-ID, via the same /api/iam/users lookup the console
  // uses. `silent` suppresses the on-screen error for a background attempt
  // (e.g. the auto-identify below) so it fails quietly into the prompt
  // instead of flashing an alert the visitor never asked to see.
  // ---------------------------------------------------------------------
  async function resolveEmailToUserId(email, opts) {
    opts = opts || {};
    if (!email || !isValidEmail(email)) {
      if (!opts.silent) showAlert('Enter a valid email address.');
      return false;
    }
    if (!cfg.dataCenter) {
      if (!opts.silent) showAlert('No Workato data center configured for this widget (data-data-center).');
      return false;
    }
    if (!opts.silent) {
      var btn = $('gwIdSubmit');
      if (btn) { btn.disabled = true; btn.textContent = '…'; }
    }
    try {
      var res = await fetch(api('/api/iam/users?email=' + encodeURIComponent(email) + '&dataCenter=' + encodeURIComponent(cfg.dataCenter)));
      var data = await res.json().catch(function () { return {}; });
      if (!res.ok) throw new Error(data.error || ('Lookup failed (' + res.status + ')'));
      var matches = data.data || [];
      if (matches.length === 0) {
        if (!opts.silent) renderEmptyState('No user found for ' + email + ' — confirm they\u2019ve been provisioned in Workspace settings.');
        return false;
      }
      var match = matches[0];
      if (match.status !== 'active') {
        if (!opts.silent) renderEmptyState(email + ' was found but isn\u2019t active yet \u2014 they likely still need to accept their workspace invite.');
        return false;
      }
      state.idpUserId = match.id;
      renderEmptyState();
      return true;
    } catch (err) {
      if (!opts.silent) renderEmptyState(err.message);
      return false;
    }
  }

  // ---------------------------------------------------------------------
  // Best-effort silent Microsoft sign-in (only runs when data-sso-client-id
  // and data-sso-tenant-id are both set — e.g. for the SharePoint full-page
  // embed). Uses the browser's existing Microsoft 365 session, no popup.
  // This is genuinely best-effort: browsers that block third-party cookies,
  // or a host page that further sandboxes this iframe, can make it fail
  // every time — that's expected, not a bug, and it falls back to the email
  // prompt above whenever it does.
  // ---------------------------------------------------------------------
  async function trySilentSSO() {
    if (!cfg.ssoClientId || !cfg.ssoTenantId || state.idpUserId) return;
    try {
      if (typeof msal === 'undefined') {
        await new Promise(function (resolve, reject) {
          var s = document.createElement('script');
          s.src = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/dist/browser/msal-browser.min.js';
          s.onload = resolve;
          s.onerror = reject;
          document.head.appendChild(s);
        });
      }
      var msalInstance = new msal.PublicClientApplication({
        auth: { clientId: cfg.ssoClientId, authority: 'https://login.microsoftonline.com/' + cfg.ssoTenantId },
        cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false },
      });
      await msalInstance.initialize();
      var result = await msalInstance.ssoSilent({ scopes: ['openid', 'profile', 'email'] });
      var email = result && result.account && (result.account.username || (result.account.idTokenClaims && result.account.idTokenClaims.email));
      if (email) await resolveEmailToUserId(email, { silent: true });
    } catch (err) {
      // Expected to fail often (no existing session, blocked third-party
      // cookies, nested-iframe sandboxing). The email prompt is already
      // showing — nothing more to do.
    }
  }

  // An ?email= query param (e.g. set by the Chrome extension's side panel
  // using chrome.identity, or an SPFx web part using pageContext.user.email)
  // or a data-user-email attribute is the most reliable signal of all, since
  // it needs no cookies or SSO session. Only fall back to best-effort silent
  // SSO when neither is present. Both wait for defaultsReady first, since
  // resolveEmailToUserId needs cfg.dataCenter — this ordering is the actual
  // fix for the "auto-identify sometimes silently does nothing" symptom.
  (async function identify() {
    await defaultsReady;
    var knownEmail = cfg.userEmail || getQueryEmail();
    if (knownEmail) {
      await resolveEmailToUserId(knownEmail, { silent: true, sourceLabel: knownEmail });
    } else {
      await trySilentSSO();
    }
  })();

  function renderMessage(role, text) {
    var empty = $('gwEmpty');
    if (empty) empty.remove();
    var wrap = document.createElement('div');
    wrap.className = 'gw-msg ' + role;
    wrap.innerHTML =
      '<div class="gw-msg-avatar">' + (role === 'user' ? 'YOU' : 'AI') + '</div>' +
      '<div><div class="gw-msg-bubble"></div></div>';
    body.appendChild(wrap);
    var bubble = wrap.querySelector('.gw-msg-bubble');
    bubble.textContent = text;
    body.scrollTop = body.scrollHeight;
    return bubble;
  }
  function renderTyping() {
    var wrap = document.createElement('div');
    wrap.className = 'gw-msg assistant';
    wrap.id = 'gwTypingIndicator';
    wrap.innerHTML = '<div class="gw-msg-avatar">AI</div><div class="gw-msg-bubble"><span class="gw-typing"><span></span><span></span><span></span></span></div>';
    body.appendChild(wrap);
    body.scrollTop = body.scrollHeight;
  }
  function removeTyping() { var el = $('gwTypingIndicator'); if (el) el.remove(); }
  function showAlert(text) {
    alertSlot.innerHTML = '<div class="gw-alert">' + escapeHtml(text) + '</div>';
  }
  function clearAlert() { alertSlot.innerHTML = ''; }

  function autoGrow() {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, 100) + 'px';
  }
  input.addEventListener('input', autoGrow);
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(input.value.trim()); }
  });
  form.addEventListener('submit', function (e) { e.preventDefault(); handleSend(input.value.trim()); });

  async function createConversation(genieId, userId) {
    var res = await fetch(api('/api/genie/conversations'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ genieId: genieId, idpUserId: userId }),
    });
    var data = await res.json().catch(function () { return {}; });
    if (!res.ok) throw new Error(data.error || ('Failed to create conversation (' + res.status + ')'));
    var conversationId = data.conversation_id || (data.result && data.result.conversation_id);
    if (!conversationId) throw new Error('Response did not include a conversation_id.');
    return conversationId;
  }

  async function sendMessageStream(genieId, userId, conversationId, message, onChunk) {
    var res = await fetch(api('/api/genie/conversations/' + conversationId + '/messages'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ genieId: genieId, idpUserId: userId, message: message }),
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
    if (!state.idpUserId) { renderEmptyState(); return; }
    var genieId = cfg.genieId;
    if (!genieId) { showAlert('No Genie ID configured for this widget (data-genie-id).'); return; }

    clearAlert();
    state.sending = true;
    sendBtn.disabled = true;
    streamState.textContent = 'Connecting…';
    renderMessage('user', message);
    input.value = '';
    autoGrow();

    try {
      if (!state.conversationId) {
        state.conversationId = await createConversation(genieId, state.idpUserId);
      }
      renderTyping();
      streamState.textContent = 'Streaming response…';
      var bubble = null;
      await sendMessageStream(genieId, state.idpUserId, state.conversationId, message, function (fullText) {
        removeTyping();
        if (!bubble) bubble = renderMessage('assistant', '');
        bubble.textContent = fullText;
        body.scrollTop = body.scrollHeight;
      });
      removeTyping();
      if (!bubble) renderMessage('assistant', '(No content returned. Check that the genie is running.)');
      streamState.textContent = 'Idle';
      if (!panel.classList.contains('open')) ping.classList.add('show');
    } catch (err) {
      removeTyping();
      streamState.textContent = 'Error';
      showAlert(err.message + ' — check your user ID and that the genie is running.');
    } finally {
      state.sending = false;
      sendBtn.disabled = false;
    }
  }

  $('gwResetBtn').addEventListener('click', function () {
    state.conversationId = null;
    streamState.textContent = 'Idle';
    clearAlert();
    renderEmptyState();
  });

  // ---------------------------------------------------------------------
  // Public API for the host page
  // ---------------------------------------------------------------------
  window.GenieWidget = {
    open: function () { setOpen(true); },
    close: function () { setOpen(false); },
    setUser: function (idpUserId) { state.idpUserId = idpUserId || ''; if (!body.querySelector('.gw-msg')) renderEmptyState(); },
    setEmail: function (email) { return resolveEmailToUserId(email, { silent: false }); },
    setGenieId: function (genieId) { cfg.genieId = genieId; },
  };
})();