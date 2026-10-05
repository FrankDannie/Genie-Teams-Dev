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
 * Chat history: each conversation (id + its messages) is saved to this
 * browser's localStorage, keyed by genie ID — there is no server-side
 * history store. That's deliberate: it needs no database, no per-user
 * accounts, and no extra backend work, at the cost of history being
 * per-browser rather than synced across devices. Reopening a past
 * conversation from the sidebar resumes the SAME upstream conversationId,
 * so the Genie still has all the prior context — it's not just a local
 * transcript, you can keep chatting in it.
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
  //
  // Deliberately plain/neutral palette (grays + one muted blue accent,
  // system font stack, no Google Fonts import) so this doesn't visually
  // clash sitting inside SharePoint or a Chrome side panel — both use
  // similar light, low-saturation, system-font UI by default.
  // ---------------------------------------------------------------------
  var style = document.createElement('style');
  style.textContent = [
    '#' + ROOT_ID + '{',
    '  --gw-primary:#2F6FED; --gw-primary-mid:#2558C4; --gw-primary-soft:#EAF1FE;',
    '  --gw-card:#FFFFFF; --gw-surface-2:#F6F7F8; --gw-border:#E1E3E6; --gw-border-light:#ECEDEF;',
    '  --gw-text:#1F2328; --gw-text-secondary:#3D4350; --gw-text-muted:#8A9099;',
    '  --gw-success:#2E8B57; --gw-danger:#C44141; --gw-warning:#A5690C;',
    '  --gw-sans:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;',
    '  --gw-mono:ui-monospace,SFMono-Regular,"Segoe UI Mono",Consolas,monospace;',
    '  font-family:var(--gw-sans); color:var(--gw-text); box-sizing:border-box;',
    '}',
    '#' + ROOT_ID + ' *{box-sizing:border-box;}',
    '#' + ROOT_ID + ' .gw-launcher{position:fixed; right:1.5rem; bottom:1.5rem; z-index:2147483000; width:54px; height:54px; border-radius:50%; border:1px solid var(--gw-border); cursor:pointer; background:var(--gw-primary); color:#fff; display:flex; align-items:center; justify-content:center; box-shadow:0 4px 14px rgba(0,0,0,0.18); transition:transform .15s ease;}',
    '#' + ROOT_ID + ' .gw-launcher:hover{transform:translateY(-2px);}',
    '#' + ROOT_ID + ' .gw-launcher svg{width:24px; height:24px;}',
    '#' + ROOT_ID + ' .gw-launcher .gw-close-icon{display:none;}',
    '#' + ROOT_ID + ' .gw-launcher.open .gw-chat-icon{display:none;}',
    '#' + ROOT_ID + ' .gw-launcher.open .gw-close-icon{display:block;}',
    '#' + ROOT_ID + ' .gw-ping{position:absolute; top:-2px; right:-2px; width:11px; height:11px; border-radius:50%; background:var(--gw-success); border:2px solid #fff; display:none;}',
    '#' + ROOT_ID + ' .gw-ping.show{display:block;}',
    '#' + ROOT_ID + ' .gw-panel{position:fixed; right:1.5rem; bottom:5.5rem; z-index:2147482999; width:380px; max-width:calc(100vw - 2.5rem); height:min(600px, calc(100vh - 8rem)); background:var(--gw-card); border:1px solid var(--gw-border); border-radius:0.5rem; box-shadow:0 12px 32px rgba(0,0,0,0.16); display:flex; flex-direction:column; overflow:hidden; transform:translateY(12px) scale(0.98); opacity:0; pointer-events:none; transition:transform .18s ease, opacity .18s ease;}',
    '#' + ROOT_ID + ' .gw-panel.open{transform:translateY(0) scale(1); opacity:1; pointer-events:auto;}',
    '#' + ROOT_ID + ' .gw-head{padding:.85rem 1rem; border-bottom:1px solid var(--gw-border); display:flex; align-items:center; justify-content:space-between; gap:.6rem; background:var(--gw-surface-2); flex-shrink:0;}',
    '#' + ROOT_ID + ' .gw-head-left{display:flex; align-items:center; gap:.4rem; min-width:0;}',
    '#' + ROOT_ID + ' .gw-head h4{margin:0; font-size:.86rem; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
    '#' + ROOT_ID + ' .gw-head-actions{display:flex; align-items:center; gap:.2rem; flex-shrink:0;}',
    '#' + ROOT_ID + ' .gw-icon-btn{background:transparent; border:1px solid transparent; color:var(--gw-text-muted); width:1.85rem; height:1.85rem; border-radius:.25rem; display:flex; align-items:center; justify-content:center; cursor:pointer; transition:all .15s ease;}',
    '#' + ROOT_ID + ' .gw-icon-btn:hover{color:var(--gw-text); background:var(--gw-border-light);}',
    '#' + ROOT_ID + ' .gw-icon-btn.active{color:var(--gw-primary); background:var(--gw-primary-soft);}',
    '#' + ROOT_ID + ' .gw-icon-btn svg{width:15px; height:15px;}',
    '#' + ROOT_ID + ' .gw-main{flex:1; display:flex; position:relative; overflow:hidden; min-height:0;}',
    '#' + ROOT_ID + ' .gw-sidebar{position:absolute; top:0; left:0; bottom:0; width:230px; max-width:82%; background:var(--gw-surface-2); border-right:1px solid var(--gw-border); display:flex; flex-direction:column; transform:translateX(-100%); transition:transform .18s ease; z-index:3;}',
    '#' + ROOT_ID + ' .gw-sidebar.open{transform:translateX(0);}',
    '#' + ROOT_ID + ' .gw-sidebar-head{display:flex; align-items:center; justify-content:space-between; padding:.7rem .75rem; border-bottom:1px solid var(--gw-border); flex-shrink:0;}',
    '#' + ROOT_ID + ' .gw-sidebar-head span{font-size:.72rem; font-weight:700; letter-spacing:.03em; text-transform:uppercase; color:var(--gw-text-muted);}',
    '#' + ROOT_ID + ' .gw-new-chat{display:flex; align-items:center; gap:.3rem; background:var(--gw-card); border:1px solid var(--gw-border); color:var(--gw-text-secondary); font-size:.68rem; font-weight:600; padding:.3rem .55rem; border-radius:.3rem; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-new-chat:hover{background:var(--gw-border-light);}',
    '#' + ROOT_ID + ' .gw-new-chat svg{width:12px; height:12px;}',
    '#' + ROOT_ID + ' .gw-sidebar-list{flex:1; overflow-y:auto; padding:.4rem;}',
    '#' + ROOT_ID + ' .gw-history-empty{padding:.6rem; font-size:.72rem; color:var(--gw-text-muted); line-height:1.5;}',
    '#' + ROOT_ID + ' .gw-history-item{padding:.5rem .55rem; border-radius:.3rem; cursor:pointer; margin-bottom:.15rem;}',
    '#' + ROOT_ID + ' .gw-history-item:hover{background:var(--gw-border-light);}',
    '#' + ROOT_ID + ' .gw-history-item.active{background:var(--gw-primary-soft);}',
    '#' + ROOT_ID + ' .gw-history-item .t{font-size:.76rem; font-weight:500; color:var(--gw-text-secondary); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
    '#' + ROOT_ID + ' .gw-history-item.active .t{color:var(--gw-primary-mid);}',
    '#' + ROOT_ID + ' .gw-history-item .d{font-size:.64rem; color:var(--gw-text-muted); margin-top:.1rem;}',
    '#' + ROOT_ID + ' .gw-body{flex:1; overflow-y:auto; padding:1rem; display:flex; flex-direction:column; gap:.9rem; background:var(--gw-card); min-width:0;}',
    '#' + ROOT_ID + ' .gw-empty{margin:auto; text-align:center; max-width:280px; color:var(--gw-text-muted); display:flex; flex-direction:column; align-items:center; gap:.7rem;}',
    '#' + ROOT_ID + ' .gw-empty .gw-glyph{width:42px; height:42px; border-radius:10px; background:var(--gw-surface-2); border:1px solid var(--gw-border); display:flex; align-items:center; justify-content:center;}',
    '#' + ROOT_ID + ' .gw-empty .gw-glyph svg{width:19px; height:19px; color:var(--gw-text-muted);}',
    '#' + ROOT_ID + ' .gw-empty h3{color:var(--gw-text-secondary); font-size:.84rem; font-weight:600; margin:0;}',
    '#' + ROOT_ID + ' .gw-empty p{font-size:.73rem; line-height:1.5; margin:0;}',
    '#' + ROOT_ID + ' .gw-id-field{display:flex; gap:.4rem; margin-top:.3rem; width:100%;}',
    '#' + ROOT_ID + ' .gw-id-field input{flex:1; min-width:0; background:#fff; border:1px solid var(--gw-border); color:var(--gw-text); font-size:.76rem; padding:.5rem .6rem; border-radius:.3rem;}',
    '#' + ROOT_ID + ' .gw-id-field input:focus{outline:none; border-color:var(--gw-primary); box-shadow:0 0 0 3px var(--gw-primary-soft);}',
    '#' + ROOT_ID + ' .gw-id-field button{flex-shrink:0; background:var(--gw-primary); color:#fff; border:none; border-radius:.3rem; font-size:.74rem; font-weight:600; padding:0 .8rem; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-msg{display:flex; gap:.55rem; max-width:88%;}',
    '#' + ROOT_ID + ' .gw-msg.user{align-self:flex-end; flex-direction:row-reverse;}',
    '#' + ROOT_ID + ' .gw-msg-avatar{flex-shrink:0; width:22px; height:22px; border-radius:6px; display:flex; align-items:center; justify-content:center; font-size:.58rem; font-weight:700; margin-top:.1rem;}',
    '#' + ROOT_ID + ' .gw-msg.assistant .gw-msg-avatar{background:var(--gw-surface-2); border:1px solid var(--gw-border); color:var(--gw-text-secondary);}',
    '#' + ROOT_ID + ' .gw-msg.user .gw-msg-avatar{background:var(--gw-primary-soft); border:1px solid var(--gw-primary-soft); color:var(--gw-primary-mid);}',
    '#' + ROOT_ID + ' .gw-msg-bubble{padding:.55rem .75rem; border-radius:.45rem; font-size:.82rem; line-height:1.5; color:var(--gw-text-secondary); white-space:pre-wrap; word-break:break-word;}',
    '#' + ROOT_ID + ' .gw-msg.assistant .gw-msg-bubble{background:var(--gw-surface-2); border:1px solid var(--gw-border-light); border-top-left-radius:.2rem;}',
    '#' + ROOT_ID + ' .gw-msg.user .gw-msg-bubble{background:var(--gw-primary-soft); border:1px solid var(--gw-primary-soft); border-top-right-radius:.2rem; color:var(--gw-text);}',
    '#' + ROOT_ID + ' .gw-typing{display:inline-flex; gap:3px; align-items:center; padding:.2rem 0;}',
    '#' + ROOT_ID + ' .gw-typing span{width:5px; height:5px; border-radius:50%; background:var(--gw-text-muted); animation:gw-bounce 1.2s infinite ease-in-out;}',
    '#' + ROOT_ID + ' .gw-typing span:nth-child(2){animation-delay:.15s;}',
    '#' + ROOT_ID + ' .gw-typing span:nth-child(3){animation-delay:.3s;}',
    '@keyframes gw-bounce{0%,80%,100%{transform:translateY(0); opacity:.4;} 40%{transform:translateY(-4px); opacity:1;}}',
    '#' + ROOT_ID + ' .gw-footer{padding:.75rem .9rem .9rem; border-top:1px solid var(--gw-border); flex-shrink:0; background:var(--gw-card);}',
    '#' + ROOT_ID + ' .gw-composer{display:flex; align-items:flex-end; gap:.5rem; background:var(--gw-surface-2); border:1px solid var(--gw-border); border-radius:.4rem; padding:.35rem .45rem; transition:border-color .15s ease;}',
    '#' + ROOT_ID + ' .gw-composer:focus-within{border-color:var(--gw-primary); box-shadow:0 0 0 3px var(--gw-primary-soft);}',
    '#' + ROOT_ID + ' .gw-composer textarea{flex:1; border:none; background:transparent; resize:none; font-family:var(--gw-sans); font-size:.82rem; color:var(--gw-text); max-height:100px; padding:.3rem;}',
    '#' + ROOT_ID + ' .gw-composer textarea:focus{outline:none;}',
    '#' + ROOT_ID + ' .gw-send{width:2rem; height:2rem; border-radius:.3rem; flex-shrink:0; background:var(--gw-primary); border:none; color:#fff; display:flex; align-items:center; justify-content:center; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-send:disabled{opacity:.35; cursor:not-allowed;}',
    '#' + ROOT_ID + ' .gw-send svg{width:14px; height:14px;}',
    '#' + ROOT_ID + ' .gw-foot-row{display:flex; justify-content:space-between; align-items:center; margin-top:.45rem; font-size:.64rem; color:var(--gw-text-muted);}',
    '#' + ROOT_ID + ' .gw-alert{display:flex; gap:.5rem; padding:.55rem .7rem; border-radius:.3rem; font-size:.73rem; line-height:1.45; margin-bottom:.7rem; border:1px solid var(--gw-danger); background:#FBEDED; color:var(--gw-danger);}',
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
    '    <div class="gw-head-left">',
    '      <button class="gw-icon-btn" id="gwHistoryBtn" type="button" title="Chat history" aria-label="Chat history"><svg viewBox="0 0 24 24" fill="none"><path d="M4 6h16M4 12h16M4 18h16" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg></button>',
    '      <h4 id="gwTitle"></h4>',
    '    </div>',
    '    <div class="gw-head-actions">',
    '      <button class="gw-icon-btn" id="gwNewChatBtn" type="button" title="New chat"><svg viewBox="0 0 24 24" fill="none"><path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg></button>',
    '      <button class="gw-icon-btn" id="gwCloseBtn" type="button" title="Close chat" aria-label="Close chat"><svg viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6 6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>',
    '    </div>',
    '  </div>',
    '  <div class="gw-main">',
    '    <div class="gw-sidebar" id="gwSidebar">',
    '      <div class="gw-sidebar-head">',
    '        <span>History</span>',
    '        <button class="gw-new-chat" id="gwSidebarNewChatBtn" type="button"><svg viewBox="0 0 24 24" fill="none"><path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>New</button>',
    '      </div>',
    '      <div class="gw-sidebar-list" id="gwSidebarList"></div>',
    '    </div>',
    '    <div class="gw-body" id="gwBody"></div>',
    '  </div>',
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
  var sidebar = $('gwSidebar'), sidebarList = $('gwSidebarList');
  var body = $('gwBody'), title = $('gwTitle');
  var form = $('gwComposerForm'), input = $('gwInput'), sendBtn = $('gwSendBtn');
  var streamState = $('gwStreamState'), alertSlot = $('gwAlertSlot');

  var state = { conversationId: null, sending: false, idpUserId: cfg.idpUserId || '', messages: [], historyTitle: '' };
  title.textContent = cfg.interfaceName;

  function api(path) { return (cfg.baseUrl || '') + path; }
  function isValidEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); }
  function getQueryEmail() {
    try {
      return new URLSearchParams(window.location.search).get('email') || '';
    } catch (e) { return ''; }
  }

  // ---------------------------------------------------------------------
  // Chat history — localStorage only, keyed per genie. See the file-level
  // comment at the top for why this is client-side rather than a backend
  // history store.
  // ---------------------------------------------------------------------
  var HISTORY_KEY = 'genieWidgetHistory:' + (cfg.genieId || 'default');
  var history = [];
  try { history = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); } catch (e) { history = []; }

  function saveHistory() {
    try { localStorage.setItem(HISTORY_KEY, JSON.stringify(history.slice(0, 50))); } catch (e) { /* storage unavailable/full — history just won't persist */ }
  }

  function upsertHistory() {
    if (!state.conversationId || state.messages.length === 0) return;
    var idx = -1;
    for (var i = 0; i < history.length; i++) { if (history[i].id === state.conversationId) { idx = i; break; } }
    var entry = {
      id: state.conversationId,
      title: state.historyTitle || (idx >= 0 ? history[idx].title : 'New chat'),
      updatedAt: new Date().toISOString(),
      messages: state.messages,
    };
    if (idx >= 0) history.splice(idx, 1);
    history.unshift(entry);
    history = history.slice(0, 50);
    saveHistory();
    renderHistoryList();
  }

  function renderHistoryList() {
    if (history.length === 0) {
      sidebarList.innerHTML = '<p class="gw-history-empty">No past chats yet — conversations you have show up here.</p>';
      return;
    }
    sidebarList.innerHTML = history.map(function (c) {
      var active = c.id === state.conversationId ? ' active' : '';
      var when = new Date(c.updatedAt).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      return '<div class="gw-history-item' + active + '" data-id="' + escapeHtml(c.id) + '">' +
        '<div class="t">' + escapeHtml(c.title) + '</div>' +
        '<div class="d">' + escapeHtml(when) + '</div></div>';
    }).join('');
    var items = sidebarList.querySelectorAll('.gw-history-item');
    for (var i = 0; i < items.length; i++) {
      items[i].addEventListener('click', function () {
        var id = this.getAttribute('data-id');
        for (var j = 0; j < history.length; j++) {
          if (history[j].id === id) { openHistoryItem(history[j]); break; }
        }
      });
    }
  }

  function openHistoryItem(entry) {
    state.conversationId = entry.id;
    state.messages = entry.messages.slice();
    state.historyTitle = entry.title;
    body.innerHTML = '';
    for (var i = 0; i < state.messages.length; i++) {
      paintMessage(state.messages[i].role, state.messages[i].text);
    }
    setSidebarOpen(false);
    renderHistoryList();
  }

  function startNewChat() {
    state.conversationId = null;
    state.messages = [];
    state.historyTitle = '';
    streamState.textContent = 'Idle';
    clearAlert();
    renderEmptyState();
    setSidebarOpen(false);
    renderHistoryList();
  }

  function setSidebarOpen(open) {
    sidebar.classList.toggle('open', open);
    $('gwHistoryBtn').classList.toggle('active', open);
  }
  $('gwHistoryBtn').addEventListener('click', function () { setSidebarOpen(!sidebar.classList.contains('open')); });
  $('gwNewChatBtn').addEventListener('click', startNewChat);
  $('gwSidebarNewChatBtn').addEventListener('click', startNewChat);
  renderHistoryList();

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
  // resolveEmailToUserId needs cfg.dataCenter.
  (async function identify() {
    await defaultsReady;
    var knownEmail = cfg.userEmail || getQueryEmail();
    if (knownEmail) {
      await resolveEmailToUserId(knownEmail, { silent: true, sourceLabel: knownEmail });
    } else {
      await trySilentSSO();
    }
  })();

  // paintMessage does DOM only, no state — used both for live messages
  // (via renderMessage below, which also records to state.messages for
  // history) and for replaying a past conversation loaded from the
  // sidebar (where state.messages is already set from storage).
  function paintMessage(role, text) {
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
  function renderMessage(role, text) {
    var bubble = paintMessage(role, text);
    state.messages.push({ role: role, text: text });
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

    if (!state.historyTitle) state.historyTitle = message.length > 40 ? message.slice(0, 40) + '…' : message;

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
      var assistantMsgIndex = -1;
      await sendMessageStream(genieId, state.idpUserId, state.conversationId, message, function (fullText) {
        removeTyping();
        if (!bubble) {
          bubble = renderMessage('assistant', '');
          assistantMsgIndex = state.messages.length - 1;
        }
        bubble.textContent = fullText;
        if (assistantMsgIndex >= 0) state.messages[assistantMsgIndex].text = fullText;
        body.scrollTop = body.scrollHeight;
      });
      removeTyping();
      if (!bubble) {
        renderMessage('assistant', '(No content returned. Check that the genie is running.)');
      }
      streamState.textContent = 'Idle';
      if (!panel.classList.contains('open')) ping.classList.add('show');
      upsertHistory();
    } catch (err) {
      removeTyping();
      streamState.textContent = 'Error';
      showAlert(err.message + ' — check your user ID and that the genie is running.');
    } finally {
      state.sending = false;
      sendBtn.disabled = false;
    }
  }

  // ---------------------------------------------------------------------
  // Public API for the host page
  // ---------------------------------------------------------------------
  window.GenieWidget = {
    open: function () { setOpen(true); },
    close: function () { setOpen(false); },
    setUser: function (idpUserId) { state.idpUserId = idpUserId || ''; if (!body.querySelector('.gw-msg')) renderEmptyState(); },
    setEmail: function (email) { return resolveEmailToUserId(email, { silent: false }); },
    setGenieId: function (genieId) { cfg.genieId = genieId; },
    newChat: function () { startNewChat(); },
  };
})();