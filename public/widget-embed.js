/*!
 * Genie Connect — embeddable chat widget
 * -----------------------------------------------------------------------
 * Drop this on any page with a single <script> tag:
 *
 *   <script
 *     src="https://YOUR-GENIE-CONNECT-HOST/widget-embed.js"
 *     data-base-url="https://YOUR-GENIE-CONNECT-HOST"
 *     data-genie-id="gin-AbMAK4r6-rXgonW-CD"
 *     data-interface-name="Smart Genie"
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
 * OAuth 2.0 (PKCE) mode — each person signs in with Workato Identity and the
 * widget talks to the Headless API as them (no shared API key, no email
 * lookup). Turn it on with data-auth-mode="oauth", or on widget.html with
 * the URL  /widget?auth=oauth&genie=gin-XXXX&name=My%20Genie  (widget.html
 * opts in to reading those params via data-url-params="true"). Needs
 * OAUTH_CLIENT_ID on the server and oauth-callback.html registered as the
 * client's redirect URL. History then comes from the Headless API itself.
 *
 * Console layout — add layout=console (widget.html URL) or data-layout="console"
 * for the full-page look: dark top bar, single-agent sidebar, hero with
 * capability chips, quick-action chips, rounded composer with attachments
 * (OAuth mode). Wording lives in CONSOLE_COPY near the top of this file.
 *
 * Chat history: saved through the backend to a Workato Data Table
 * (GET/POST /api/history), keyed by idpUserId — NOT localStorage anymore.
 * This means history now follows the person across browsers and devices,
 * at the cost of needing that one Data Table configured server-side
 * (DATATABLE_ID / DATATABLE_API_TOKEN). Reopening a past conversation
 * resumes the SAME upstream conversationId, so the Genie still has all
 * the prior context — it's not just a local transcript.
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
    userEmail: (thisScript && thisScript.getAttribute('data-user-email')) || '',
    ssoClientId: (thisScript && thisScript.getAttribute('data-sso-client-id')) || '',
    ssoTenantId: (thisScript && thisScript.getAttribute('data-sso-tenant-id')) || '',
    authMode: (thisScript && thisScript.getAttribute('data-auth-mode')) || '',
    layout: (thisScript && thisScript.getAttribute('data-layout')) || '',
    oauth: null, // filled from /api/health when authMode is "oauth"
  };
  // widget.html (and only pages that opt in) may be configured from the URL.
  if (thisScript && thisScript.getAttribute('data-url-params') === 'true') {
    try {
      var qp = new URLSearchParams(window.location.search);
      if (/^gin-[\w-]+$/.test(qp.get('genie') || '')) cfg.genieId = qp.get('genie');
      if (qp.get('auth') === 'oauth') cfg.authMode = 'oauth';
      if (qp.get('layout') === 'console') cfg.layout = 'console';
      if (qp.get('name')) cfg.interfaceName = qp.get('name').slice(0, 60);
    } catch (e) { /* ignore */ }
  }
  var isOAuth = cfg.authMode === 'oauth';
  var isConsole = cfg.layout === 'console';

  var ROOT_ID = 'genie-widget-root';
  if (document.getElementById(ROOT_ID)) return; // already injected

  // =====================================================================
  // "Console" layout — ?layout=console (widget.html) or data-layout="console".
  // A full-page look: dark top bar, one-agent sidebar, agent header, hero
  // with capability chips, quick-action chips and a rounded composer.
  // It re-uses every element id the chat logic below already expects, so
  // sign-in, streaming, history and approvals work exactly as before.
  // Edit CONSOLE_COPY to change the wording; nothing else needs touching.
  // =====================================================================
  var CONSOLE_COPY = {
    brand: 'Agentic Solutions',
    tagline: 'DEMO CONSOLE - WOW 2026',
    logoUrl: '/assets/wow-logo.png',            // optional: drop the real logo here; text logo shows if the file is missing
    agentName: cfg.interfaceName,
    agentTag: 'IT',
    agentSub: 'Idea Lifestyle Furniture \u00b7 IT suite',
    description: 'Self-service IT helpdesk agent for access, credentials, group management, and ticket resolution \u2014 with human escalation when needed.',
    company: 'IDEA LIFESTYLE FURNITURE',
    capabilities: ['Access Requests', 'Approvals', 'Password Reset', 'Account Unlock', 'Group Management', 'Ticket Resolution'],
    quick: ['Request application access', 'Reset my password', 'Unlock my account', 'What do I have access to?'],
    placeholder: 'Ask IT Support Agent anything\u2026',
    disclaimer: 'Demo environment \u2014 please don\u2019t enter personal or confidential information.',
    hint: 'Enter to send \u00b7 Shift+Enter for a new line \u00b7 attachments up to 20 MB',
  };
  var ICON = {
    headset: '<path d="M3 11h3a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-5Zm0 0a9 9 0 1 1 18 0m0 0v5a2 2 0 0 1-2 2h-1a2 2 0 0 1-2-2v-3a2 2 0 0 1 2-2h3Z"/><path d="M21 16v2a4 4 0 0 1-4 4h-5"/>',
    bell: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
    gear: '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    reset: '<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>',
    logout: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="m16 17 5-5-5-5"/><path d="M21 12H9"/>',
    clip: '<path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/>',
    send: '<path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/>',
    bolt: '<path d="M13 2 3 14h9l-1 8 10-12h-9l1-8z"/>',
    shield: '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    close: '<path d="M6 6l12 12M18 6 6 18"/>',
  };
  function svgIcon(name, size) {
    return '<svg viewBox="0 0 24 24" width="' + size + '" height="' + size + '" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICON[name] + '</svg>';
  }

  function consoleHtml() {
    var C = CONSOLE_COPY;
    var name = escapeHtml(C.agentName);
    var quick = C.quick.map(function (q) {
      return '<button type="button" class="gc-quick-btn" data-q="' + escapeHtml(q) + '">' + svgIcon('bolt', 11) + '<span>' + escapeHtml(q) + '</span></button>';
    }).join('');
    return [
      '<button class="gw-launcher" id="gwLauncher" type="button" aria-hidden="true" tabindex="-1"><span class="gw-ping" id="gwPing"></span></button>',
      '<header class="gc-top">',
      '  <div class="gc-brand">',
      '    <div class="gc-logo"><img class="gc-logo-img" src="' + escapeHtml(C.logoUrl) + '" alt="World of Workato">',
      '      <div class="gc-logo-fallback" style="display:none"><span class="gc-logo-wow">WOW</span><span class="gc-logo-sub">World of<br>Workato</span></div></div>',
      '    <div class="gc-brand-text"><div class="gc-brand-name">' + escapeHtml(C.brand) + '</div><div class="gc-brand-tag">' + escapeHtml(C.tagline) + '</div></div>',
      '  </div>',
      '  <div class="gc-top-right">',
      '    <button type="button" class="gc-ghost" aria-label="Notifications" title="Notifications">' + svgIcon('bell', 16) + '</button>',
      '    <button type="button" class="gc-settings">' + svgIcon('gear', 14) + '<span>Settings</span></button>',
      '    <span class="gc-sep"></span>',
      '    <div class="gc-user"><div class="gc-user-name" id="gwUserName"></div><div class="gc-user-email" id="gwUserEmail"></div></div>',
      '    <button type="button" class="gc-ghost" id="gwSignOut" aria-label="Sign out" title="Sign out" style="display:none">' + svgIcon('logout', 16) + '</button>',
      '  </div>',
      '</header>',
      '<div class="gc-body">',
      '  <aside class="gc-side">',
      '    <div class="gc-side-label">AGENTS</div>',
      '    <div class="gc-agent active"><span class="gc-tile gc-tile-sm">' + svgIcon('headset', 14) + '</span>',
      '      <span class="gc-agent-text"><span class="gc-agent-name">' + name + '</span><span class="gc-agent-tag">' + escapeHtml(C.agentTag) + '</span></span></div>',
      '  </aside>',
      '  <section class="gc-main open" id="gwPanel">',
      '    <div class="gc-agent-head">',
      '      <div class="gc-agent-id"><span class="gc-tile">' + svgIcon('headset', 17) + '</span>',
      '        <div><h4 id="gwTitle"></h4><div class="gc-agent-sub">' + escapeHtml(C.agentSub) + '</div></div></div>',
      '      <div class="gc-agent-actions">',
      '        <span class="gc-pill" id="gwPill" data-state="idle"><i class="gc-dot"></i><span id="gwStreamState">Ready</span></span>',
      '        <button type="button" class="gc-btn" id="gwHistoryBtn" title="Chat history">' + svgIcon('clock', 14) + '<span>History</span></button>',
      '        <button type="button" class="gc-btn" id="gwNewChatBtn" title="Start a new conversation">' + svgIcon('reset', 14) + '<span>Reset demo</span></button>',
      '        <button type="button" id="gwCloseBtn" hidden aria-hidden="true"></button>',
      '      </div>',
      '    </div>',
      '    <div class="gw-main gc-content">',
      '      <div class="gw-sidebar" id="gwSidebar">',
      '        <div class="gw-sidebar-head"><span>History</span>',
      '          <button class="gw-new-chat" id="gwSidebarNewChatBtn" type="button">' + svgIcon('plus', 12) + 'New</button></div>',
      '        <div class="gw-sidebar-list" id="gwSidebarList"></div>',
      '      </div>',
      '      <div class="gw-body" id="gwBody"></div>',
      '    </div>',
      '    <div class="gc-footer"><div class="gc-footer-inner">',
      '      <div id="gwAlertSlot"></div>',
      '      <div class="gc-quick">' + quick + '</div>',
      '      <div class="gc-attach" id="gwAttachChip"></div>',
      '      <form class="gc-composer" id="gwComposerForm">',
      '        <button type="button" class="gc-attach-btn" id="gwAttachBtn" aria-label="Attach a file" title="Attach a file">' + svgIcon('clip', 16) + '</button>',
      '        <input type="file" id="gwFile" hidden>',
      '        <textarea id="gwInput" rows="1" placeholder="' + escapeHtml(C.placeholder) + '" aria-label="Message"></textarea>',
      '        <button class="gc-send" id="gwSendBtn" type="submit" aria-label="Send message">' + svgIcon('send', 15) + '</button>',
      '      </form>',
      '      <div class="gc-disclaimer">' + svgIcon('shield', 12) + '<span>' + escapeHtml(C.disclaimer) + '</span></div>',
      '      <div class="gc-hint">' + escapeHtml(C.hint) + '</div>',
      '    </div></div>',
      '  </section>',
      '</div>',
    ].join('\n');
  }

  function consoleCss() {
    var R = '#' + ROOT_ID;
    return [
      R + '.gc-root{--gw-primary:#2D6A62; --gw-primary-mid:#245750; --gw-primary-soft:#E3EEEB; --gw-card:#FFFFFF; --gw-surface-2:#F4F5F4; --gw-border:#DDE3E1; --gw-border-light:#E7ECEA; --gw-text:#15282C; --gw-text-secondary:#33474B; --gw-text-muted:#8A9794; --gw-success:#1F8F5F; --gw-danger:#C0392B; position:fixed; inset:0; display:flex; flex-direction:column; background:#F4F5F4; font-family:"Instrument Sans",Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; -webkit-font-smoothing:antialiased;}',
      R + ' .gw-launcher{display:none !important;}',
      R + ' button{font-family:inherit;}',
      // top bar
      R + ' .gc-top{height:47px; flex-shrink:0; background:#203438; color:#fff; display:flex; align-items:center; justify-content:space-between; padding:0 17px; gap:12px;}',
      R + ' .gc-brand{display:flex; align-items:center; gap:14px; min-width:0;}',
      R + ' .gc-logo{display:flex; align-items:center; flex-shrink:0;}',
      R + ' .gc-logo-img{height:25px; width:auto; display:block;}',
      R + ' .gc-logo-fallback{align-items:flex-end; gap:3px;}',
      R + ' .gc-logo-wow{font-weight:900; font-size:25px; line-height:.9; letter-spacing:-.03em; background:linear-gradient(90deg,#7B5CE5,#C27FE6); -webkit-background-clip:text; background-clip:text; color:transparent;}',
      R + ' .gc-logo-sub{font-size:5.5px; line-height:1.2; letter-spacing:.09em; color:#B9A9E8; text-transform:uppercase;}',
      R + ' .gc-brand-text{display:flex; flex-direction:column; gap:3px; min-width:0;}',
      R + ' .gc-brand-name{font-size:12.5px; font-weight:600; line-height:1.1; white-space:nowrap;}',
      R + ' .gc-brand-tag{font-size:8.5px; letter-spacing:.2em; color:#9DB0B3; line-height:1; white-space:nowrap;}',
      R + ' .gc-top-right{display:flex; align-items:center; gap:10px; flex-shrink:0;}',
      R + ' .gc-ghost{width:28px; height:28px; display:inline-flex; align-items:center; justify-content:center; background:transparent; border:none; border-radius:7px; color:#C5D2D4; cursor:pointer;}',
      R + ' .gc-ghost:hover{background:rgba(255,255,255,.08); color:#fff;}',
      R + ' .gc-settings{height:32px; padding:0 12px; display:inline-flex; align-items:center; gap:6px; background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.22); border-radius:9px; color:#fff; font-size:11.5px; font-weight:600; cursor:pointer;}',
      R + ' .gc-sep{width:1px; height:32px; background:rgba(255,255,255,.14); margin:0 4px;}',
      R + ' .gc-user{text-align:right; line-height:1.25; max-width:220px;}',
      R + ' .gc-user-name{font-size:11.5px; font-weight:600; color:#fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
      R + ' .gc-user-email{font-size:9.5px; color:#A9BBBE; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
      // body + sidebar
      R + ' .gc-body{flex:1; display:flex; min-height:0;}',
      R + ' .gc-side{width:182px; flex-shrink:0; background:#F2F4F3; border-right:1px solid #E1E6E4; padding:20px 9px; overflow-y:auto;}',
      R + ' .gc-side-label{font-size:8.5px; font-weight:700; letter-spacing:.2em; color:#8A9792; padding:0 10px; margin-bottom:9px;}',
      R + ' .gc-agent{display:flex; align-items:center; gap:10px; padding:0 10px; height:45px; border-radius:10px; color:#15282C;}',
      R + ' .gc-agent.active{background:#E3EEEB;}',
      R + ' .gc-agent-text{display:flex; flex-direction:column; gap:1px; min-width:0;}',
      R + ' .gc-agent-name{font-size:11.5px; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
      R + ' .gc-agent-tag{font-size:9.5px; color:#8A9792;}',
      R + ' .gc-tile{width:32px; height:32px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; border-radius:8px; background:#EAF1EF; border:1px solid #D5E1DE; color:#2D6A62;}',
      R + ' .gc-tile-sm{width:26px; height:26px; border-radius:7px; background:#D3E4E0; border-color:#C2D8D3;}',
      // main column
      R + ' .gc-main{flex:1; min-width:0; display:flex; flex-direction:column; background:#F4F5F4;}',
      R + ' .gc-agent-head{height:58px; flex-shrink:0; display:flex; align-items:center; justify-content:space-between; gap:12px; padding:0 20px; border-bottom:1px solid #E1E6E4; background:#F7F8F7;}',
      R + ' .gc-agent-id{display:flex; align-items:center; gap:11px; min-width:0;}',
      R + ' .gc-agent-id h4{margin:0; font-size:13px; font-weight:700; color:#15282C; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
      R + ' .gc-agent-sub{font-size:10.5px; color:#7F8D8A; margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}',
      R + ' .gc-agent-actions{display:flex; align-items:center; gap:10px; flex-shrink:0;}',
      R + ' .gc-pill{height:21px; display:inline-flex; align-items:center; gap:6px; padding:0 10px; border-radius:999px; font-size:10.5px; font-weight:600; background:#E6F6EE; border:1px solid #BFE4D1; color:#1F8F5F;}',
      R + ' .gc-dot{width:5px; height:5px; border-radius:50%; background:currentColor; display:inline-block;}',
      R + ' .gc-pill[data-state="idle"]{background:#EEF0EF; border-color:#DDE3E1; color:#7F8D8A;}',
      R + ' .gc-pill[data-state="busy"]{background:#FFF6E5; border-color:#F1DDB0; color:#A5690C;}',
      R + ' .gc-pill[data-state="error"]{background:#FBEDED; border-color:#EBC4C4; color:#C0392B;}',
      R + ' .gc-btn{height:32px; display:inline-flex; align-items:center; gap:7px; padding:0 13px; background:#fff; border:1px solid #DDE3E1; border-radius:9px; color:#15282C; font-size:11.5px; font-weight:600; cursor:pointer;}',
      R + ' .gc-btn:hover{background:#F4F7F6;}',
      R + ' .gc-btn.active{background:#E3EEEB; border-color:#C2D8D3;}',
      R + ' .gc-content{flex:1; min-height:0;}',
      R + ' .gw-body{background:transparent; padding:20px max(20px, calc((100% - 700px) / 2));}',
      R + ' .gw-sidebar{background:#F7F8F7; z-index:5;}',
      R + ' .gw-msg{max-width:92%;}',
      R + ' .gw-msg.user .gw-msg-bubble{background:#E3EEEB; border-color:#CFDDD9; color:#15282C;}',
      R + ' .gw-msg.assistant .gw-msg-bubble{background:#fff; border-color:#E3E8E6;}',
      R + ' .gw-msg-bubble{font-size:12.5px;}',
      // hero
      R + ' .gc-hero{max-width:560px; align-items:center; gap:0; color:#7F8D8A;}',
      R + ' .gc-hero-icon{width:52px; height:52px; border-radius:12px; background:#E7EEEC; border:1px solid #D5E1DE; display:flex; align-items:center; justify-content:center; color:#2D6A62; margin-bottom:20px;}',
      R + ' .gc-hero-title{margin:0; font-size:18px; font-weight:700; color:#15282C; letter-spacing:-.01em;}',
      R + ' .gc-hero-desc{margin:11px 0 0; font-size:11.5px; line-height:1.65; color:#7B8986; max-width:420px;}',
      R + ' .gc-hero-company{margin-top:16px; font-size:9px; letter-spacing:.22em; color:#9AA5A2; font-weight:600;}',
      R + ' .gc-chips{display:flex; flex-wrap:wrap; justify-content:center; gap:7px; margin-top:16px; max-width:380px;}',
      R + ' .gc-chip{height:23px; display:inline-flex; align-items:center; padding:0 11px; border-radius:999px; background:#E4EEEB; border:1px solid #CFDDD9; color:#2D6A62; font-size:10.5px; font-weight:500;}',
      // footer
      R + ' .gc-footer{flex-shrink:0; padding:12px 0 12px; background:linear-gradient(180deg,rgba(244,245,244,0),#fff 55%); border-top:1px solid #E7ECEA; display:flex; justify-content:center;}',
      R + ' .gc-footer-inner{width:min(583px, calc(100% - 40px));}',
      R + ' .gc-quick{display:flex; flex-wrap:wrap; justify-content:center; gap:7px; margin-bottom:9px;}',
      R + ' .gc-quick-btn{height:21px; display:inline-flex; align-items:center; gap:6px; padding:0 11px; background:#fff; border:1px solid #DDE3E1; border-radius:999px; color:#33474B; font-size:10.5px; font-weight:500; cursor:pointer;}',
      R + ' .gc-quick-btn:hover{background:#F4F7F6; border-color:#C9D6D2;}',
      R + ' .gc-quick-btn svg{color:#6B7B78;}',
      R + ' .gc-composer{display:flex; align-items:center; gap:8px; min-height:49px; padding:8px 8px 8px 12px; background:#fff; border:1px solid #DDE3E1; border-radius:14px; box-shadow:0 1px 2px rgba(20,40,44,.04);}',
      R + ' .gc-composer:focus-within{border-color:#2D6A62; box-shadow:0 0 0 3px rgba(45,106,98,.12);}',
      R + ' .gc-composer textarea{flex:1; min-width:0; border:none; background:transparent; resize:none; outline:none; font-family:inherit; font-size:12px; line-height:1.5; color:#15282C; max-height:110px; padding:4px 2px;}',
      R + ' .gc-composer textarea::placeholder{color:#9AA5A2;}',
      R + ' .gc-attach-btn{width:28px; height:28px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; background:transparent; border:none; border-radius:7px; color:#8A9794; cursor:pointer;}',
      R + ' .gc-attach-btn:hover{background:#F0F3F2; color:#33474B;}',
      R + '.gc-legacy .gc-attach-btn{display:none;}',
      R + ' .gc-send{width:33px; height:33px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; border:none; border-radius:9px; background:#E3EEEB; color:#9AB0AB; cursor:pointer; transition:background .15s ease,color .15s ease;}',
      R + ' .gc-send.has-text{background:#2D6A62; color:#fff;}',
      R + ' .gc-send:disabled{opacity:.5; cursor:not-allowed;}',
      R + ' .gc-attach{display:none; margin-bottom:7px;}',
      R + ' .gc-attach.show{display:flex; align-items:center; gap:8px; width:fit-content; max-width:100%; padding:4px 6px 4px 10px; background:#fff; border:1px solid #DDE3E1; border-radius:999px; font-size:10.5px; color:#33474B;}',
      R + ' .gc-attach-name{overflow:hidden; text-overflow:ellipsis; white-space:nowrap;}',
      R + ' .gc-attach button{border:none; background:#EEF2F1; color:#5B6B68; width:18px; height:18px; border-radius:50%; cursor:pointer; line-height:1; font-size:12px;}',
      R + ' .gc-disclaimer{display:flex; align-items:center; gap:6px; margin-top:9px; font-size:10.5px; color:#7F8D8A;}',
      R + ' .gc-hint{margin-top:4px; font-size:9.5px; color:#A7B1AE;}',
      R + ' .gw-alert{margin-bottom:8px;}',
      '@media (max-width:760px){' + R + ' .gc-side{display:none;} ' + R + ' .gc-brand-tag,' + R + ' .gc-user,' + R + ' .gc-settings span,' + R + ' .gc-btn span{display:none;} ' + R + ' .gc-agent-head{padding:0 12px;}}',
    ].join('\n');
  }

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
    '#' + ROOT_ID + ' .gw-note{font-size:.7rem; color:var(--gw-text-muted); text-align:center; line-height:1.4;}',
    '#' + ROOT_ID + ' .gw-link{color:var(--gw-primary); text-decoration:none; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-card-note{border:1px solid var(--gw-border); background:var(--gw-surface-2); border-radius:.4rem; padding:.6rem .7rem; font-size:.76rem; color:var(--gw-text-secondary);}',
    '#' + ROOT_ID + ' .gw-card-note pre{margin:.45rem 0 0; padding:.45rem; max-height:140px; overflow:auto; background:#fff; border:1px solid var(--gw-border-light); border-radius:.3rem; font-family:var(--gw-mono); font-size:.68rem; white-space:pre-wrap; word-break:break-word;}',
    '#' + ROOT_ID + ' .gw-card-actions{display:flex; gap:.4rem; margin-top:.55rem;}',
    '#' + ROOT_ID + ' .gw-card-actions button{border:1px solid var(--gw-primary); background:var(--gw-primary); color:#fff; border-radius:.3rem; font-size:.72rem; font-weight:600; padding:.35rem .75rem; cursor:pointer;}',
    '#' + ROOT_ID + ' .gw-card-actions button.secondary{background:#fff; color:var(--gw-text-secondary); border-color:var(--gw-border);}',
    '#' + ROOT_ID + ' .gw-card-actions button:disabled{opacity:.5; cursor:not-allowed;}',
  ].join('\n');
  if (isConsole) style.textContent += '\n' + consoleCss();
  document.head.appendChild(style);

  var root = document.createElement('div');
  root.id = ROOT_ID;
  root.innerHTML = isConsole ? consoleHtml() : [
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
    '    <div class="gw-foot-row"><span id="gwStreamState">Idle</span><span id="gwUserInfo"></span></div>',
    '  </div>',
    '</div>',
  ].join('\n');
  if (isConsole) { root.classList.add('gc-root'); if (!isOAuth) root.classList.add('gc-legacy'); }
  document.body.appendChild(root);

  var $ = function (id) { return document.getElementById(id); };
  var launcher = $('gwLauncher'), panel = $('gwPanel'), ping = $('gwPing');
  var sidebar = $('gwSidebar'), sidebarList = $('gwSidebarList');
  var body = $('gwBody'), title = $('gwTitle');
  var form = $('gwComposerForm'), input = $('gwInput'), sendBtn = $('gwSendBtn');
  var streamStateEl = $('gwStreamState'), alertSlot = $('gwAlertSlot');
  var streamState = streamStateEl;
  if (isConsole) {
    streamState = {};
    Object.defineProperty(streamState, 'textContent', {
      get: function () { return streamStateEl.textContent; },
      set: function (v) {
        var st = v === 'Idle' ? 'ready' : (v === 'Error' ? 'error' : 'busy');
        streamStateEl.textContent = v === 'Idle' ? 'Ready' : v;
        var pill = document.getElementById('gwPill');
        if (pill) pill.setAttribute('data-state', st);
      },
    });
  }

  var state = { conversationId: null, sending: false, idpUserId: cfg.idpUserId || '', messages: [], historyTitle: '' };
  var oauth = { status: 'loading', session: null, claims: null, pending: null, waiting: null, popup: null, waitTimer: null, refreshing: null, note: '', turn: null, turnHandler: null };
  title.textContent = cfg.interfaceName;

  function api(path) { return (cfg.baseUrl || '') + path; }
  function isValidEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); }
  function getQueryEmail() {
    try {
      return new URLSearchParams(window.location.search).get('email') || '';
    } catch (e) { return ''; }
  }

  var history = [];

  async function loadHistoryFromServer() {
    if (!state.idpUserId) return;
    try {
      var res = await fetch(api('/api/history?idpUserId=' + encodeURIComponent(state.idpUserId) + '&dataCenter=' + encodeURIComponent(cfg.dataCenter || '')));
      var data = await res.json().catch(function () { return {}; });
      var rows = data.data || data.records || [];
      history = rows.map(function (r) {
        var fields = r.data || r;
        var msgs = [];
        try { msgs = JSON.parse(fields.messages || '[]'); } catch (e) { msgs = []; }
        return {
          id: fields.conversation_id,
          title: fields.title || 'New chat',
          updatedAt: fields.updated_at || new Date().toISOString(),
          messages: msgs,
        };
      }).sort(function (a, b) { return new Date(b.updatedAt) - new Date(a.updatedAt); });
    } catch (e) {
      history = [];
    }
    renderHistoryList();
  }

  async function upsertHistory() {
    if (!state.conversationId || state.messages.length === 0 || !state.idpUserId) return;
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
    renderHistoryList();

    try {
      await fetch(api('/api/history'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          idpUserId: state.idpUserId,
          genieId: cfg.genieId,
          conversationId: state.conversationId,
          title: entry.title,
          messages: state.messages,
          dataCenter: cfg.dataCenter || '',
        }),
      });
    } catch (e) { /* save failed silently — this session still works */ }
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
    if (isOAuth && entry.messages === null) { loadConversationOAuth(entry); return; }
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

  var defaultsReady = (async function loadDefaults() {
    try {
      var res = await fetch(api('/api/health'));
      var data = await res.json();
      if (!cfg.genieId && data.defaultGenieId) cfg.genieId = data.defaultGenieId;
      if (!cfg.dataCenter && data.defaultDataCenter) cfg.dataCenter = data.defaultDataCenter;
      if (data.oauthClientId) cfg.oauth = { clientId: data.oauthClientId, authorizeUrl: data.oauthAuthorizeUrl, redirectUri: data.oauthRedirectUri, scope: data.oauthScope };
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
    if (isConsole) { syncPill(); if (consoleReady()) { renderHero(); return; } }
    var needsId = !state.idpUserId;
    body.innerHTML =
      '<div class="gw-empty" id="gwEmpty">' +
      '  <div class="gw-glyph"><svg viewBox="0 0 24 24" fill="none"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg></div>' +
      (isOAuth ? oauthEmptyInner(statusLine) : needsId
        ? '  <h3>Sign in to start</h3><p>Enter your email once to start chatting.</p>' +
          '  <div class="gw-id-field"><input id="gwIdInput" type="text" placeholder="you@company.com" spellcheck="false" /><button id="gwIdSubmit" type="button">Start</button></div>' +
          (statusLine ? '  <p style="margin-top:.5rem;font-size:.7rem;color:var(--gw-text-muted);">' + escapeHtml(statusLine) + '</p>' : '')
        : '  <h3>No conversation yet</h3><p>Send a message below to get started.</p>') +
      '</div>';
    if (isOAuth) { wireOAuthEmpty(); return; }
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
      loadHistoryFromServer();
      return true;
    } catch (err) {
      if (!opts.silent) renderEmptyState(err.message);
      return false;
    }
  }

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
    } catch (err) { /* expected to fail often — email prompt already showing */ }
  }

  (async function identify() {
    await defaultsReady;
    if (isOAuth) { await initOAuth(); return; }
    if (state.idpUserId) { loadHistoryFromServer(); return; }
    var knownEmail = cfg.userEmail || getQueryEmail();
    if (knownEmail) {
      await resolveEmailToUserId(knownEmail, { silent: true, sourceLabel: knownEmail });
    } else {
      await trySilentSSO();
    }
  })();

  // =====================================================================
  // OAuth 2.0 (PKCE) mode — enabled with ?auth=oauth (widget page) or
  // data-auth-mode="oauth". Signs the person in through Workato Identity
  // and talks to the Headless API as THAT user (no shared API key, no
  // email lookup). Everything below is inert unless isOAuth is true.
  // =====================================================================
  var memStore = {};
  function storeGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return memStore[k] || null; } }
  function storeSet(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { memStore[k] = v; } }
  function storeDel(k) { try { window.localStorage.removeItem(k); } catch (e) { /* ignore */ } delete memStore[k]; }
  function sessionKey() { return 'genie_oauth_v1:' + (cfg.genieId || ''); }
  function enc(v) { return encodeURIComponent(v); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function inIframe() { try { return window.top !== window.self; } catch (e) { return true; } }

  function b64url(buf) {
    var bytes = new Uint8Array(buf), s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomUrlSafe(len) {
    var a = new Uint8Array(len);
    window.crypto.getRandomValues(a);
    return b64url(a).slice(0, len);
  }
  // code_challenge = BASE64URL(SHA256(ASCII(code_verifier)))
  async function makePending() {
    var verifier = randomUrlSafe(64);
    var digest = await window.crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    oauth.pending = { verifier: verifier, challenge: b64url(digest), state: randomUrlSafe(32) };
  }
  function decodeJwt(token) {
    try {
      var p = String(token).split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      while (p.length % 4) p += '=';
      var bin = atob(p), bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return JSON.parse(new TextDecoder().decode(bytes));
    } catch (e) { return null; }
  }

  function saveSession(tok) {
    var prev = oauth.session || {};
    oauth.session = {
      access_token: tok.access_token,
      refresh_token: tok.refresh_token || prev.refresh_token || '', // refresh tokens rotate — always keep the newest
      id_token: tok.id_token || prev.id_token || '',
      expires_at: Date.now() + (Number(tok.expires_in) || 3600) * 1000,
    };
    oauth.claims = decodeJwt(oauth.session.id_token); // display only — never trusted for access decisions
    storeSet(sessionKey(), JSON.stringify(oauth.session));
  }
  function loadSession() {
    try {
      var raw = storeGet(sessionKey());
      if (!raw) return false;
      oauth.session = JSON.parse(raw);
      oauth.claims = decodeJwt(oauth.session.id_token || '');
      return !!(oauth.session && oauth.session.access_token);
    } catch (e) { return false; }
  }
  function clearSession() { oauth.session = null; oauth.claims = null; storeDel(sessionKey()); }

  // Back-channel token calls go through this site's own worker (same-origin).
  async function tokenCall(payload) {
    var res = await fetch(api('/api/oauth/token'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    var data = await res.json().catch(function () { return {}; });
    if (!res.ok || !data.access_token) {
      var err = new Error(data.error_description || data.error || ('Token request failed (' + res.status + ')'));
      err.status = res.status;
      err.oauthError = data.error;
      throw err;
    }
    return data;
  }

  // Refresh tokens are single-use, so only ever one refresh in flight.
  function refreshSession() {
    if (oauth.refreshing) return oauth.refreshing;
    var rt = oauth.session && oauth.session.refresh_token;
    if (!rt) return Promise.resolve(false);
    oauth.refreshing = tokenCall({ grant_type: 'refresh_token', refresh_token: rt })
      .then(function (t) { saveSession(t); return true; })
      .catch(function (err) {
        // Only a definitive rejection ends the session; a network blip keeps it.
        if (err.oauthError === 'invalid_grant' || err.status === 400 || err.status === 401) clearSession();
        return false;
      })
      .then(function (ok) { oauth.refreshing = null; return ok; });
    return oauth.refreshing;
  }
  async function getAccessToken() {
    if (!oauth.session) return null;
    if (oauth.session.expires_at - 60000 > Date.now()) return oauth.session.access_token;
    return (await refreshSession()) && oauth.session ? oauth.session.access_token : null;
  }

  function requireLogin(note) {
    oauth.status = 'signedout';
    oauth.note = note || '';
    updateUserInfo();
    if (!body.querySelector('.gw-msg')) renderEmptyState();
    else showAlert(note || 'Please sign in again.');
  }

  function authorizeUrl(p) {
    var c = cfg.oauth;
    return c.authorizeUrl +
      '?response_type=code' +
      '&client_id=' + enc(c.clientId) +
      '&redirect_uri=' + enc(c.redirectUri) +
      '&scope=' + enc(c.scope || 'openid profile email') +
      '&state=' + enc(p.state) +
      '&code_challenge=' + enc(p.challenge) +
      '&code_challenge_method=S256';
  }

  // Must run synchronously inside the click so the browser allows the popup,
  // which is why the PKCE values are prepared ahead of time (makePending).
  function startSignIn() {
    if (!cfg.oauth || !cfg.oauth.clientId) return;
    if (!oauth.pending) { makePending().then(startSignIn); return; }
    var p = oauth.pending;
    oauth.pending = null;
    var url = authorizeUrl(p);

    if (!inIframe()) {
      // Own tab: a normal full-page redirect; oauth-callback.html finishes it.
      try {
        window.sessionStorage.setItem('genie_oauth_pending', JSON.stringify({
          state: p.state, verifier: p.verifier, redirectUri: cfg.oauth.redirectUri,
          genieId: cfg.genieId, returnUrl: window.location.href,
        }));
      } catch (e) { showAlert('Your browser blocked the storage needed to sign in.'); makePending(); return; }
      window.location.assign(url);
      return;
    }

    // Embedded (e.g. SharePoint): Workato Identity can't load inside a frame,
    // so sign in through a popup and receive the result via postMessage.
    var popup = window.open(url, 'genie_oauth', 'width=520,height=720');
    makePending(); // fresh values ready for the next attempt
    if (!popup) {
      oauth.status = 'signedout';
      oauth.note = 'Your browser blocked the sign-in window. Allow pop-ups for this site, or open the chat in a new tab.';
      renderEmptyState();
      return;
    }
    oauth.popup = popup;
    oauth.waiting = p;
    oauth.status = 'waiting';
    oauth.note = '';
    renderEmptyState();
    clearTimeout(oauth.waitTimer);
    oauth.waitTimer = setTimeout(function () {
      if (oauth.waiting !== p) return;
      oauth.waiting = null;
      oauth.status = 'signedout';
      oauth.note = 'Sign-in timed out. Please try again.';
      renderEmptyState();
    }, 5 * 60 * 1000);
  }

  if (isOAuth) {
    window.addEventListener('message', async function (ev) {
      if (ev.origin !== window.location.origin) return;           // only our own callback page
      var d = ev.data;
      if (!d || d.source !== 'genie-oauth') return;
      var p = oauth.waiting;
      if (!p || d.state !== p.state) return;                       // CSRF check
      oauth.waiting = null;
      clearTimeout(oauth.waitTimer);
      if (d.error || !d.code) {
        oauth.status = 'signedout';
        oauth.note = 'Sign-in failed: ' + (d.error_description || d.error || 'no authorization code returned');
        renderEmptyState();
        return;
      }
      try {
        var t = await tokenCall({ grant_type: 'authorization_code', code: d.code, code_verifier: p.verifier, redirect_uri: cfg.oauth.redirectUri });
        saveSession(t);
        oauth.status = 'signedin';
        oauth.note = '';
        renderEmptyState();
        updateUserInfo();
        loadHistoryOAuth();
      } catch (err) {
        oauth.status = 'signedout';
        oauth.note = 'Sign-in failed: ' + err.message;
        renderEmptyState();
      }
    });
  }

  function signOut() {
    clearSession();
    history = [];
    oauth.status = 'signedout';
    oauth.note = '';
    updateUserInfo();
    startNewChat();
  }

  function updateUserInfo() {
    if (isConsole) {
      var on = !!(isOAuth && oauth.session);
      var cc = (on && oauth.claims) || {};
      var em = cc.email || cc.preferred_username || '';
      $('gwUserName').textContent = on ? (cc.name || (em ? em.split('@')[0] : 'Signed in')) : '';
      $('gwUserEmail').textContent = on ? em : '';
      $('gwSignOut').style.display = on ? '' : 'none';
      syncPill();
      return;
    }
    var el = $('gwUserInfo');
    if (!el) return;
    if (!isOAuth || !oauth.session) { el.innerHTML = ''; return; }
    var c = oauth.claims || {};
    var who = c.email || c.name || c.preferred_username || 'Signed in';
    el.innerHTML = '<span>' + escapeHtml(who) + '</span> \u00b7 <a href="#" class="gw-link" id="gwSignOut">Sign out</a>';
    $('gwSignOut').addEventListener('click', function (e) { e.preventDefault(); signOut(); });
  }

  function oauthEmptyInner(statusLine) {
    var openTab = inIframe()
      ? '<p style="margin-top:.4rem;font-size:.7rem;">Trouble signing in? <a href="#" class="gw-link" id="gwOpenTab">Open chat in a new tab</a></p>'
      : '';
    var note = (oauth.note || statusLine)
      ? '<p style="margin-top:.4rem;font-size:.72rem;color:var(--gw-danger);">' + escapeHtml(oauth.note || statusLine) + '</p>'
      : '';
    switch (oauth.status) {
      case 'loading':
        return '  <h3>Checking sign-in\u2026</h3>';
      case 'unconfigured':
        return '  <h3>Sign-in isn\u2019t available</h3><p>The Genie Connect server didn\u2019t return OAuth settings. Check that OAUTH_CLIENT_ID is set.</p>';
      case 'waiting':
        return '  <h3>Finish signing in</h3><p>Complete the sign-in in the window that just opened.</p>' + openTab;
      case 'signedin':
        return '  <h3>No conversation yet</h3><p>Send a message below to get started.</p>';
      default:
        return '  <h3>Sign in to start</h3><p>Sign in with your Workato Identity account to chat.</p>' +
          '  <div class="gw-id-field"><button id="gwSignIn" type="button" style="flex:1;padding:.6rem 1rem;">Sign in with Workato</button></div>' +
          note + openTab;
    }
  }
  function wireOAuthEmpty() {
    var b = $('gwSignIn'); if (b) b.addEventListener('click', startSignIn);
    var t = $('gwOpenTab');
    if (t) t.addEventListener('click', function (e) { e.preventDefault(); window.open(window.location.href, '_blank'); });
  }

  async function initOAuth() {
    if (!cfg.oauth || !cfg.oauth.clientId || !window.crypto || !window.crypto.subtle) {
      oauth.status = 'unconfigured';
      renderEmptyState();
      return;
    }
    await makePending();
    if (loadSession()) {
      var tok = await getAccessToken();
      if (tok) {
        oauth.status = 'signedin';
        renderEmptyState();
        updateUserInfo();
        loadHistoryOAuth();
        return;
      }
    }
    oauth.status = 'signedout';
    renderEmptyState();
  }

  // ---- Headless API calls (as the signed-in user) ----------------------
  async function hlFetch(path, init) {
    init = init || {};
    for (var attempt = 0; attempt < 2; attempt++) {
      var tok = await getAccessToken();
      if (!tok) { requireLogin('Your session expired. Sign in again to continue.'); throw new Error('Not signed in'); }
      var headers = Object.assign({}, init.headers || {}, { Authorization: 'Bearer ' + tok });
      var res = await fetch(api('/api/headless/' + enc(cfg.genieId) + '/' + path), Object.assign({}, init, { headers: headers }));
      if (res.status === 401 && attempt === 0 && oauth.session) { oauth.session.expires_at = 0; continue; } // force one refresh, then retry
      return res;
    }
    throw new Error('Request failed');
  }
  async function httpError(res, prefix) {
    var t = await res.text().catch(function () { return ''; });
    var m = '';
    try {
      var j = JSON.parse(t);
      m = j.error_description || j.error || j.message || '';
      if (typeof m !== 'string') m = JSON.stringify(m);
    } catch (e) { m = t.slice(0, 200); }
    return new Error((prefix || 'Request failed') + ' (' + res.status + ')' + (m ? ': ' + m : ''));
  }

  async function hlCreateConversation() {
    var res = await hlFetch('conversations', { method: 'POST', headers: { 'Content-Type': 'application/json' } });
    if (!res.ok) throw await httpError(res, 'Failed to create conversation');
    var data = await res.json().catch(function () { return {}; });
    var id = data.conversation_id || (data.result && data.result.conversation_id);
    if (!id) throw new Error('Response did not include a conversation_id.');
    return id;
  }

  // Reads one SSE response. Returns where the genie run stands so the
  // caller can decide whether to reconnect.
  async function consumeSSE(res, onEvent, prev) {
    var turn = {
      runId: (prev && prev.runId) || '', lastEventId: (prev && prev.lastEventId) || '',
      finished: false, interrupted: false, awaiting: false, retryAfter: 0,
    };
    function handleBlock(block) {
      var dataLines = [], evName = '', id = '';
      block.split(/\r?\n/).forEach(function (l) {
        if (l.indexOf('data:') === 0) dataLines.push(l.slice(5).replace(/^ /, ''));
        else if (l.indexOf('event:') === 0) evName = l.slice(6).trim();
        else if (l.indexOf('id:') === 0) id = l.slice(3).trim();
      });
      if (!dataLines.length) return;
      var raw = dataLines.join('\n');
      if (raw === '[DONE]') return;
      var ev;
      try { ev = JSON.parse(raw); } catch (e) { return; }
      if (!ev.type && evName) ev.type = evName;
      var eid = id || ev.event_id;
      if (eid) turn.lastEventId = eid;
      if (ev.genie_run_id) turn.runId = ev.genie_run_id;
      if (ev.type === 'processing.finished') turn.finished = true;
      if (ev.type === 'system.stream_interrupted') { turn.interrupted = true; turn.retryAfter = Number(ev.retry_after_ms) || 0; }
      if (ev.type === 'skill.confirmation_required' || ev.type === 'runtime_connection.auth_required') turn.awaiting = true;
      onEvent(ev);
    }
    var reader = res.body.getReader(), dec = new TextDecoder(), buf = '';
    while (true) {
      var out = await reader.read();
      if (out.done) break;
      buf += dec.decode(out.value, { stream: true });
      var parts = buf.split(/\r?\n\r?\n/);
      buf = parts.pop();
      for (var i = 0; i < parts.length; i++) handleBlock(parts[i]);
    }
    if (buf.trim()) handleBlock(buf);
    return turn;
  }

  // Reconnect to a run (after a dropped stream, or after an approval) until it finishes.
  async function followRun(turn) {
    var tries = 0;
    while (!turn.finished && !turn.awaiting && turn.runId && tries < 8) {
      tries++;
      if (turn.interrupted || tries > 1) await sleep(turn.retryAfter || 1500);
      var headers = { Accept: 'text/event-stream' };
      if (turn.lastEventId) headers['Last-Event-ID'] = turn.lastEventId;
      var res = await hlFetch('conversations/' + enc(state.conversationId) + '/genie-runs/' + enc(turn.runId), { method: 'GET', headers: headers });
      if (!res.ok || !res.body) throw await httpError(res, 'Could not reconnect to the response');
      turn = await consumeSSE(res, oauth.turnHandler, turn);
    }
    return turn;
  }

  function renderNote(text) {
    var d = document.createElement('div');
    d.className = 'gw-note';
    d.textContent = text;
    body.appendChild(d);
    body.scrollTop = body.scrollHeight;
  }

  function makeEventHandler() {
    return function (ev) {
      switch (ev.type) {
        case 'agent.message': {
          var text = typeof ev.message === 'string' ? ev.message : (ev.message && (ev.message.content || ev.message.text)) || '';
          if (!text) return;
          removeTyping();
          renderMessage('assistant', text);
          break;
        }
        case 'skill.running':
          streamState.textContent = 'Running ' + (ev.skill_name || 'a skill') + '\u2026';
          break;
        case 'skill.completed':
        case 'skill.stopped':
          streamState.textContent = 'Finished ' + (ev.skill_name || 'a skill');
          break;
        case 'skill.failed':
          renderNote('A step failed: ' + (ev.skill_name || 'skill') + (ev.error ? ' \u2014 ' + (typeof ev.error === 'string' ? ev.error : JSON.stringify(ev.error)) : ''));
          break;
        case 'skill.confirmation_required':
          removeTyping();
          renderApprovalCard(ev);
          break;
        case 'runtime_connection.auth_required':
          removeTyping();
          renderConnectCard(ev);
          break;
        default: break; // processing.*, system.ping, … need no UI
      }
    };
  }

  function cardShell(html) {
    var wrap = document.createElement('div');
    wrap.className = 'gw-card-note';
    wrap.innerHTML = html;
    body.appendChild(wrap);
    body.scrollTop = body.scrollHeight;
    return wrap;
  }

  function renderApprovalCard(ev) {
    var params = '';
    try { params = ev.skill_parameters ? JSON.stringify(ev.skill_parameters, null, 2).slice(0, 1500) : ''; } catch (e) {}
    var card = cardShell(
      '<div>Approval needed to run <b>' + escapeHtml(ev.skill_name || 'a skill') + '</b></div>' +
      (params ? '<pre>' + escapeHtml(params) + '</pre>' : '') +
      '<div class="gw-card-actions"><button type="button" data-r="approved">Approve</button><button type="button" class="secondary" data-r="rejected">Reject</button></div>'
    );
    var btns = card.querySelectorAll('button');
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener('click', async function () {
        var resolution = this.getAttribute('data-r');
        for (var j = 0; j < btns.length; j++) btns[j].disabled = true;
        try {
          var res = await hlFetch('conversations/' + enc(state.conversationId) + '/skill_approval/' + enc(ev.call_id), {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ resolution: resolution }),
          });
          if (!res.ok) throw await httpError(res, 'Could not record your decision');
          card.querySelector('.gw-card-actions').outerHTML = '<div class="gw-note">' + (resolution === 'approved' ? 'Approved.' : 'Rejected.') + '</div>';
          resumeTurn();
        } catch (err) {
          showAlert(err.message);
          for (var k = 0; k < btns.length; k++) btns[k].disabled = false;
        }
      });
    }
  }

  function renderConnectCard(ev) {
    var connector = (ev.auth_link && ev.auth_link.connector_name) || 'an app';
    var card = cardShell(
      '<div>The genie needs you to connect <b>' + escapeHtml(connector) + '</b> to continue.</div>' +
      '<div class="gw-card-actions"><button type="button" data-a="connect">Connect</button><button type="button" class="secondary" data-a="skip">Skip</button></div>'
    );
    var actions = card.querySelector('.gw-card-actions');
    actions.querySelector('[data-a="connect"]').addEventListener('click', async function () {
      var btn = this; btn.disabled = true;
      try {
        var url = ev.auth_link && ev.auth_link.url;
        var status = url ? 'auth_required' : '';
        if (!url) {
          var res = await hlFetch('runtime_connection/' + enc(ev.runtime_connection_attempt_id) + '/link', { method: 'POST', headers: { 'Content-Type': 'application/json' } });
          if (!res.ok) throw await httpError(res, 'Could not get the connection link');
          var d = await res.json().catch(function () { return {}; });
          status = d.status; url = d.auth_link && d.auth_link.url;
        }
        if (status === 'authorized' || !url) { actions.outerHTML = '<div class="gw-note">Already connected.</div>'; resumeTurn(); return; }
        window.open(url, '_blank');
        actions.innerHTML = '<button type="button" data-a="continue">I\u2019ve connected \u2014 continue</button>';
        actions.querySelector('[data-a="continue"]').addEventListener('click', function () { actions.outerHTML = ''; resumeTurn(); });
      } catch (err) { showAlert(err.message); btn.disabled = false; }
    });
    actions.querySelector('[data-a="skip"]').addEventListener('click', async function () {
      try {
        await hlFetch('runtime_connection/' + enc(ev.runtime_connection_attempt_id) + '/reject', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
        actions.outerHTML = '<div class="gw-note">Skipped.</div>';
        resumeTurn();
      } catch (err) { showAlert(err.message); }
    });
  }

  async function resumeTurn() {
    if (!oauth.turn || !oauth.turn.runId) return;
    state.sending = true; sendBtn.disabled = true;
    oauth.turn.awaiting = false;
    renderTyping();
    streamState.textContent = 'Waiting for the genie\u2026';
    try { oauth.turn = await followRun(oauth.turn); }
    catch (err) { showAlert(err.message); }
    finally {
      removeTyping();
      state.sending = false; sendBtn.disabled = false;
      finishTurn();
    }
  }

  function finishTurn() {
    var t = oauth.turn;
    streamState.textContent = t && t.awaiting ? 'Waiting for your response\u2026' : 'Idle';
    if (!panel.classList.contains('open')) ping.classList.add('show');
    loadHistoryOAuth();
  }

  async function handleSendOAuth(message) {
    if (oauth.status !== 'signedin') { renderEmptyState(); return; }
    if (!cfg.genieId) { showAlert('No Genie ID configured (use ?genie=gin-\u2026 on the widget URL).'); return; }
    if (!state.historyTitle) state.historyTitle = message.length > 40 ? message.slice(0, 40) + '\u2026' : message;

    clearAlert();
    state.sending = true;
    sendBtn.disabled = true;
    streamState.textContent = 'Connecting\u2026';
    renderMessage('user', message);
    input.value = '';
    autoGrow();

    try {
      if (!state.conversationId) state.conversationId = await hlCreateConversation();
      var fileId = '';
      if (state.pendingFile) {
        streamState.textContent = 'Uploading\u2026';
        fileId = await hlUploadFile(state.pendingFile);
        renderNote('Attached: ' + state.pendingFile.name);
        state.pendingFile = null;
        if (fileInput) fileInput.value = '';
        renderAttachChip();
      }
      renderTyping();
      streamState.textContent = 'Streaming response\u2026';
      oauth.turnHandler = makeEventHandler();
      var res = await hlFetch('conversations/' + enc(state.conversationId) + '/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Accept: 'text/event-stream' },
        body: JSON.stringify(fileId ? { message: message, stream: true, file_id: fileId } : { message: message, stream: true }),
      });
      if (!res.ok || !res.body) throw await httpError(res, 'Message request failed');
      oauth.turn = await consumeSSE(res, oauth.turnHandler, null);
      oauth.turn = await followRun(oauth.turn); // no-op unless the stream dropped early
      removeTyping();
      if (!body.querySelector('.gw-msg.assistant') && !oauth.turn.awaiting) {
        renderNote('No reply came back. Check that the genie is running.');
      }
      finishTurn();
    } catch (err) {
      removeTyping();
      streamState.textContent = 'Error';
      if (err.message !== 'Not signed in') showAlert(err.message);
    } finally {
      state.sending = false;
      sendBtn.disabled = false;
    }
  }

  // ---- history comes straight from the Headless API --------------------
  async function loadHistoryOAuth() {
    if (oauth.status !== 'signedin') return;
    try {
      var res = await hlFetch('conversations?limit=50', { method: 'GET', headers: { Accept: 'application/json' } });
      var data = await res.json().catch(function () { return {}; });
      history = (data.list || []).map(function (r) {
        return { id: r.conversation_id, title: r.topic || 'New chat', updatedAt: r.last_updated_at || r.created_at || new Date().toISOString(), messages: null };
      }).sort(function (a, b) { return new Date(b.updatedAt) - new Date(a.updatedAt); });
    } catch (e) { history = []; }
    renderHistoryList();
  }

  async function loadConversationOAuth(entry) {
    state.conversationId = entry.id;
    state.messages = [];
    state.historyTitle = entry.title;
    body.innerHTML = '';
    setSidebarOpen(false);
    renderHistoryList();
    renderTyping();
    try {
      var res = await hlFetch('conversations/' + enc(entry.id) + '/messages?limit=100', { method: 'GET', headers: { Accept: 'application/json' } });
      if (!res.ok) throw await httpError(res, 'Could not load this conversation');
      var data = await res.json().catch(function () { return {}; });
      var msgs = (data.messages || []).slice().reverse().map(function (m) { // API returns newest first
        return { role: m.source === 'user' ? 'user' : 'assistant', text: m.content || '' };
      });
      removeTyping();
      entry.messages = msgs;
      state.messages = msgs.slice();
      for (var i = 0; i < msgs.length; i++) paintMessage(msgs[i].role, msgs[i].text);
    } catch (err) {
      removeTyping();
      if (err.message !== 'Not signed in') showAlert(err.message);
    }
  }

  // ---- console layout wiring ---------------------------------------------
  var syncSend = function () { if (isConsole) sendBtn.classList.toggle('has-text', !!input.value.trim()); };
  var attachBtn = $('gwAttachBtn'), fileInput = $('gwFile'), attachChip = $('gwAttachChip');
  var MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

  function consoleReady() { return isOAuth ? oauth.status === 'signedin' : !!state.idpUserId; }

  // The status pill under the agent name: Ready / Sign in required / busy states.
  function syncPill() {
    if (!isConsole || state.sending) return;
    var txt, st;
    if (consoleReady()) { txt = 'Ready'; st = 'ready'; }
    else if (isOAuth && oauth.status === 'waiting') { txt = 'Signing in\u2026'; st = 'busy'; }
    else if (isOAuth && oauth.status === 'loading') { txt = 'Checking sign-in\u2026'; st = 'busy'; }
    else { txt = 'Sign in required'; st = 'idle'; }
    streamStateEl.textContent = txt;
    var pill = $('gwPill');
    if (pill) pill.setAttribute('data-state', st);
  }

  function renderHero() {
    var C = CONSOLE_COPY;
    body.innerHTML =
      '<div class="gw-empty gc-hero" id="gwEmpty">' +
      '  <div class="gc-hero-icon">' + svgIcon('headset', 28) + '</div>' +
      '  <h2 class="gc-hero-title">' + escapeHtml(cfg.interfaceName) + '</h2>' +
      '  <p class="gc-hero-desc">' + escapeHtml(C.description) + '</p>' +
      '  <div class="gc-hero-company">' + escapeHtml(C.company) + '</div>' +
      '  <div class="gc-chips">' + C.capabilities.map(function (c) { return '<span class="gc-chip">' + escapeHtml(c) + '</span>'; }).join('') + '</div>' +
      '</div>';
  }

  function renderAttachChip() {
    if (!attachChip) return;
    if (!state.pendingFile) { attachChip.className = 'gc-attach'; attachChip.innerHTML = ''; return; }
    attachChip.className = 'gc-attach show';
    attachChip.innerHTML = '<span class="gc-attach-name">' + escapeHtml(state.pendingFile.name) + '</span>' +
      '<button type="button" id="gwAttachRemove" aria-label="Remove attachment">\u00d7</button>';
    $('gwAttachRemove').addEventListener('click', function () {
      state.pendingFile = null;
      if (fileInput) fileInput.value = '';
      renderAttachChip();
    });
  }

  // Headless API: upload first, then send the returned file_id with the message.
  async function hlUploadFile(file) {
    var fd = new FormData();
    fd.append('file', file, file.name);
    var res = await hlFetch('conversations/' + enc(state.conversationId) + '/upload', { method: 'POST', body: fd });
    if (!res.ok) throw await httpError(res, 'Could not upload the file');
    var d = await res.json().catch(function () { return {}; });
    if (!d.file_id) throw new Error('The upload response did not include a file_id.');
    return d.file_id;
  }

  if (isConsole) {
    var logoImg = root.querySelector('.gc-logo-img');
    if (logoImg) logoImg.addEventListener('error', function () {
      logoImg.style.display = 'none';
      var fb = root.querySelector('.gc-logo-fallback');
      if (fb) fb.style.display = 'flex';
    });

    var quickBtns = root.querySelectorAll('.gc-quick-btn');
    for (var qi = 0; qi < quickBtns.length; qi++) {
      quickBtns[qi].addEventListener('click', function () {
        if (state.sending) return;
        if (!consoleReady()) { renderEmptyState(); return; }
        handleSend(this.getAttribute('data-q'));
      });
    }
    input.addEventListener('input', syncSend);
    $('gwSignOut').addEventListener('click', function () { if (isOAuth) signOut(); });

    if (attachBtn && fileInput) {
      attachBtn.addEventListener('click', function () {
        if (!consoleReady()) { renderEmptyState(); return; }
        fileInput.click();
      });
      fileInput.addEventListener('change', function () {
        var f = fileInput.files && fileInput.files[0];
        if (!f) return;
        if (f.size > MAX_UPLOAD_BYTES) {
          fileInput.value = '';
          showAlert('That file is larger than 20 MB.');
          return;
        }
        clearAlert();
        state.pendingFile = f;
        renderAttachChip();
      });
    }
    syncSend();
    syncPill();
  }

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
    syncSend();
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
        if (!line || !line.startsWith('data:')) continue;
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
    if (isOAuth) return handleSendOAuth(message);
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

  window.GenieWidget = {
    open: function () { setOpen(true); },
    close: function () { setOpen(false); },
    setUser: function (idpUserId) { state.idpUserId = idpUserId || ''; if (!body.querySelector('.gw-msg')) renderEmptyState(); if (idpUserId) loadHistoryFromServer(); },
    setEmail: function (email) { return resolveEmailToUserId(email, { silent: false }); },
    setGenieId: function (genieId) { cfg.genieId = genieId; },
    newChat: function () { startNewChat(); },
    signOut: function () { if (isOAuth) signOut(); },
  };
})();