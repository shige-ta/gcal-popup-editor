/*
  GCal Popup Editor - content script
  - Detects Google Calendar event quick popup dialogs
  - Injects an editor UI (title + description + Save/Cancel)
  - On Save, automates the official UI: opens editor, fills fields, clicks Save
  - No Google Calendar API used

  Notes:
  - Google Calendar DOM changes often. Selectors are resilient and multi-lingual where possible.
  - Uses Shadow DOM for styling isolation.
*/

(() => {
  const LOG_PREFIX = '[GCalPopupEditor]';
  const DEBUG = false;

  const log = (...args) => DEBUG && console.debug(LOG_PREFIX, ...args);
  const warn = (...args) => console.warn(LOG_PREFIX, ...args);

  // i18n-ish labels used to find elements reliably in JP/EN
  const LABELS = {
    // Buttons that open the full editor
    editEvent: [
      /^edit\s*event$/i,
      /^edit$/i,
      /^open\s*detailed\s*view$/i,
      /^詳細を表示/,
      /^予定を編集/,
      /^編集$/
    ],
    // Title field labels/placeholders
    title: [
      /^title$/i,
      /^event\s*title$/i,
      /^add\s*title$/i,
      /^タイトル$/, /^タイトルを追加$/, /^件名$/, /^件名を追加$/
    ],
    // Description/Notes field labels/placeholders
    description: [
      /^description$/i,
      /^notes?$/i,
      /^add\s*description$/i,
      /^説明$/, /^説明を追加$/, /^メモ$/, /^メモを追加$/
    ],
    // Save/Confirm buttons
    save: [
      /^save$/i,
      /^save\s*&\s*close$/i,
      /^update$/i,
      /^done$/i,
      /^保存$/, /^保存して閉じる$/, /^更新$/, /^完了$/, /^送信$/
    ],
    discard: [/^discard/i, /^破棄$/]
  };

  function matchesAny(text, regexList) {
    if (!text) return false;
    for (const r of regexList) {
      if (r.test(text.trim())) return true;
    }
    return false;
  }

  function waitFor(conditionFn, { timeout = 15000, interval = 100 } = {}) {
    const start = Date.now();
    return new Promise((resolve, reject) => {
      const tick = () => {
        try {
          const res = conditionFn();
          if (res) return resolve(res);
        } catch (e) {
          // ignore and keep polling
        }
        if (Date.now() - start >= timeout) return reject(new Error('Timeout'));
        setTimeout(tick, interval);
      };
      tick();
    });
  }

  function dispatchInputEvents(el) {
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function setTextInputValue(input, value) {
    const descriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(input), 'value');
    if (descriptor && descriptor.set) {
      descriptor.set.call(input, value);
    } else {
      input.value = value;
    }
    dispatchInputEvents(input);
  }

  function setContentEditableText(box, text) {
    box.focus();
    try {
      // Clear existing content
      document.execCommand('selectAll', false, undefined);
      document.execCommand('insertText', false, text);
    } catch (e) {
      // Fallback: direct textContent assignment
      box.textContent = text;
      dispatchInputEvents(box);
    }
  }

  function queryClosestDialog(node) {
    return node.closest('div[role="dialog"], div[role="region"]');
  }

  function findQuickPopupDialogs(root = document.body) {
    // Heuristic: small dialogs with role=dialog and not full-screen editors
    const dialogs = Array.from(root.querySelectorAll('div[role="dialog"]'));
    return dialogs.filter(d => d.offsetParent !== null && d.getBoundingClientRect().width < 700);
  }

  function isVisible(el) { return !!(el && el.offsetParent !== null); }

  function findEditButton(container) {
    // Prefer searching within the same dialog region, then fallback to whole document.
    const scopes = [];
    if (container) scopes.push(container);
    const nearest = queryClosestDialog(container || document.body);
    if (nearest && nearest !== container) scopes.push(nearest);
    scopes.push(document);

    const selectors = [
      'div[role="button"][aria-label]','button[aria-label]','[role="button"][aria-label]','[aria-label]'
    ];

    for (const scope of scopes) {
      // aria-label match
      let cands = selectors.flatMap(sel => Array.from(scope.querySelectorAll(sel))).filter(isVisible);
      let btn = cands.find(el => matchesAny(el.getAttribute('aria-label'), LABELS.editEvent));
      if (btn) return btn;

      // tooltip/title/textContent match
      cands = Array.from(scope.querySelectorAll('div[role="button"],button,[role="menuitem"]')).filter(isVisible);
      btn = cands.find(el => matchesAny(el.getAttribute('data-tooltip') || el.getAttribute('title') || el.textContent, LABELS.editEvent));
      if (btn) return btn;
    }
    return null;
  }

  function extractTitleFromPopup(container) {
    // Try heading role or h2
    const heading = container.querySelector('[role="heading"], h2, h1');
    return heading ? heading.textContent.trim() : '';
  }

  function extractDescriptionFromPopup(container) {
    // Try common description blocks inside the quick popup
    // Look for nodes labelled Description/説明, or a block after an icon
    const labelled = Array.from(container.querySelectorAll('[aria-label], [data-tooltip]'))
      .find(el => matchesAny(el.getAttribute('aria-label') || el.getAttribute('data-tooltip'), LABELS.description));
    if (labelled) {
      // The labelled element may itself be the editor/view; prefer its text
      const text = labelled.textContent.trim();
      if (text) return text;
    }

    // Fallback: search for long text blocks likely being description
    const paragraphs = Array.from(container.querySelectorAll('div, p'))
      .map(el => el.textContent?.trim() || '')
      .filter(t => t && t.length >= 5)
      .sort((a, b) => b.length - a.length);
    return paragraphs[0] || '';
  }

  function createEditorUI(initial) {
    const host = document.createElement('div');
    host.className = 'gpe-host';
    const shadow = host.attachShadow({ mode: 'open' });

    const style = document.createElement('style');
    style.textContent = `
      :host, .gpe { font-family: Roboto, Arial, sans-serif; }
      :host { all: initial; }
      .gpe { box-sizing: border-box; margin-top: 8px; }
      :root, .gpe { --gpe-bg:#fff; --gpe-fg:#1f1f1f; --gpe-border:#dadce0; --gpe-muted:#5f6368; --gpe-primary:#1a73e8; --gpe-shadow: 0 1px 2px rgba(0,0,0,.08), 0 4px 12px rgba(0,0,0,.06); }
      @media (prefers-color-scheme: dark) {
        :root, .gpe { --gpe-bg:#202124; --gpe-fg:#e8eaed; --gpe-border:#3c4043; --gpe-muted:#9aa0a6; --gpe-shadow: 0 1px 2px rgba(0,0,0,.6), 0 4px 12px rgba(0,0,0,.4); }
      }
      .togglebar { display:flex; justify-content:flex-end; padding: 0 0 6px 0; }
      .toggle-btn { display:inline-flex; align-items:center; gap:6px; padding:4px 8px; font-size:12px; border-radius:6px; cursor:pointer; user-select:none; border:1px solid var(--gpe-border); background: var(--gpe-bg); color: var(--gpe-primary); }
      .toggle-btn .ic { width:14px; height:14px; }
      .card { background: var(--gpe-bg); color: var(--gpe-fg); border: 1px solid var(--gpe-border); border-radius: 10px; box-shadow: var(--gpe-shadow); overflow: hidden; }
      .toolbar { display:flex; align-items:center; justify-content:space-between; padding: 8px 10px; border-bottom: 1px solid var(--gpe-border); gap:8px; }
      .toolbar .left { display:flex; align-items:center; gap:8px; min-width:0; }
      .toolbar .title { font-size: 13px; font-weight: 600; white-space: nowrap; }
      .badge { font-size:10px; color: var(--gpe-muted); border:1px solid var(--gpe-border); padding:2px 6px; border-radius: 999px; }
      .toolbar .right { display:flex; gap:6px; }
      .btn { display:inline-flex; align-items:center; gap:6px; padding: 6px 10px; font-size:12px; border-radius: 6px; cursor:pointer; user-select:none; border:1px solid var(--gpe-border); background: var(--gpe-bg); color: var(--gpe-primary); }
      .btn[disabled] { opacity:.55; cursor:not-allowed; }
      .btn.primary { background: var(--gpe-primary); border-color: var(--gpe-primary); color:#fff; }
      .btn.ghost { background: transparent; color: var(--gpe-primary); }
      .btn .ic { width:14px; height:14px; display:inline-block; }
      .content { padding: 10px; display:flex; flex-direction:column; gap:10px; }
      .field { position:relative; }
      .field input { width:100%; font-size:13px; color:var(--gpe-fg); background: var(--gpe-bg); border:1px solid var(--gpe-border); border-radius:8px; padding: 16px 12px 10px 12px; outline:none; box-sizing:border-box; transition:border-color .15s ease; }
      .field input:focus { border-color: var(--gpe-primary); }
      .field label { position:absolute; left:12px; top:10px; font-size:12px; color:var(--gpe-muted); background:var(--gpe-bg); padding:0 4px; transform-origin:left top; transition: transform .12s ease, color .12s ease, top .12s ease; pointer-events:none; }
      .field.filled label, .field:focus-within label { top:-7px; transform: scale(.88); color: var(--gpe-primary); }
      .status { display:flex; align-items:center; gap:8px; padding: 8px 10px; border-top: 1px solid var(--gpe-border); font-size:11px; color:var(--gpe-muted); min-height: 18px; }
      .spinner { width:14px; height:14px; border:2px solid var(--gpe-border); border-top-color: var(--gpe-primary); border-radius:50%; animation: gpe_spin .9s linear infinite; }
      .check { width:14px; height:14px; color:#188038; }
      .hidden { display:none !important; }
      @keyframes gpe_spin { to { transform: rotate(360deg); } }
    `;

    const wrap = document.createElement('div');
    wrap.className = 'gpe';
    wrap.innerHTML = `
      <div class="togglebar">
        <button class="toggle-btn" data-action="toggle" title="クイック編集の表示/非表示">
          <svg class="ic" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 6a9.77 9.77 0 0 1 9 6 9.77 9.77 0 0 1-9 6 9.77 9.77 0 0 1-9-6 9.77 9.77 0 0 1 9-6zm0 2a4 4 0 1 0 .001 8.001A4 4 0 0 0 12 8z"/></svg>
          <span class="toggle-text">クイック編集を隠す</span>
        </button>
      </div>
      <div class="card">
        <div class="toolbar">
          <div class="left">
            <svg class="ic" viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm18-11.5a1.003 1.003 0 0 0 0-1.42l-1.59-1.59a1.003 1.003 0 0 0-1.42 0l-1.83 1.83 3.75 3.75L21 5.75z"/></svg>
            <span class="title">Quick Edit</span>
            <span class="badge">beta</span>
          </div>
          <div class="right">
            <button class="btn ghost" data-action="reload" title="Reload (Alt+R)">
              <svg class="ic" viewBox="0 0 24 24" fill="currentColor"><path d="M12 6V3L8 7l4 4V8c2.76 0 5 2.24 5 5a5 5 0 11-5-5z"/></svg>
              <span>Reload</span>
            </button>
            <button class="btn ghost" data-action="cancel" title="Cancel (Esc)">
              <svg class="ic" viewBox="0 0 24 24" fill="currentColor"><path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
              <span>Cancel</span>
            </button>
            <button class="btn primary" data-action="save" title="Save (Ctrl/Cmd+S)" disabled>
              <svg class="ic" viewBox="0 0 24 24" fill="currentColor"><path d="M17 3H5c-1.1 0-2 .9-2 2v14l4-4h10c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2z"/></svg>
              <span>Save</span>
            </button>
          </div>
        </div>
        <div class="content">
          <div class="field f-title">
            <input type="text" class="gpe-title" id="gpe-title" />
            <label for="gpe-title">Title / タイトル</label>
          </div>
        </div>
        <div class="status" aria-live="polite">
          <span class="spinner hidden" aria-hidden="true"></span>
          <svg class="check hidden" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M9 16.17l-3.88-3.88L4 13.41 9 18.41 20.59 6.83 19.17 5.41z"/></svg>
          <span class="text"></span>
        </div>
      </div>
    `;

    shadow.appendChild(style);
    shadow.appendChild(wrap);

    const titleEl = wrap.querySelector('.gpe-title');
    const card = wrap.querySelector('.card');
    const toggleBtn = wrap.querySelector('.toggle-btn');
    const toggleText = wrap.querySelector('.toggle-text');
    const statusText = wrap.querySelector('.status .text');
    const spinner = wrap.querySelector('.status .spinner');
    const check = wrap.querySelector('.status .check');
    const saveBtn = wrap.querySelector('button[data-action="save"]');
    const cancelBtn = wrap.querySelector('button[data-action="cancel"]');
    const reloadBtn = wrap.querySelector('button[data-action="reload"]');

    const ui = {
      host,
      shadow,
      root: wrap,
      title: titleEl,
      status: statusText,
      buttons: { save: saveBtn, cancel: cancelBtn, reload: reloadBtn, toggle: toggleBtn },
      on(action, fn) {
        wrap.addEventListener('click', (e) => {
          const btn = e.target.closest('button[data-action]');
          if (btn && btn.dataset.action === action) fn(e);
        });
      },
      trigger(action) { const btn = wrap.querySelector(`button[data-action="${action}"]`); if (btn) btn.click(); },
      setStatus(msg) { statusText.textContent = msg || ''; },
      setSaving(isSaving) {
        if (isSaving) { spinner.classList.remove('hidden'); check.classList.add('hidden'); }
        else { spinner.classList.add('hidden'); }
        titleEl.disabled = !!isSaving; saveBtn.disabled = !!isSaving || !dirty();
      },
      markSaved() { baseline.title = titleEl.value; updateDirty(); },
      setCollapsed(collapsed) {
        card.classList.toggle('hidden', !!collapsed);
        toggleText.textContent = collapsed ? 'クイック編集を表示' : 'クイック編集を隠す';
      }
    };

    // Initialize values
    titleEl.value = initial.title || '';

    // Floating labels state
    const titleField = wrap.querySelector('.f-title');
    function updateFilled() {
      titleField.classList.toggle('filled', !!titleEl.value.trim());
    }

    // Dirty state tracking
    const baseline = { title: titleEl.value };
    function dirty() { return titleEl.value !== baseline.title; }
    function updateDirty() { saveBtn.disabled = !dirty(); }

    // Wire inputs
    ['input','change'].forEach(ev => titleEl.addEventListener(ev, () => { updateFilled(); updateDirty(); }));
    updateFilled(); updateDirty();

    // Toggle show/hide
    toggleBtn.addEventListener('click', () => {
      const collapsed = !card.classList.contains('hidden');
      ui.setCollapsed(collapsed);
    });
    ui.setCollapsed(false);

    // Keyboard shortcuts within shadow
    shadow.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); ui.trigger('save'); }
      else if (e.key === 'Escape') { e.preventDefault(); ui.trigger('cancel'); }
      else if (e.altKey && (e.key.toLowerCase() === 'r')) { e.preventDefault(); ui.trigger('reload'); }
    });

    // Save success check pulse helper
    ui.pulseCheck = () => { check.classList.remove('hidden'); setTimeout(() => check.classList.add('hidden'), 1200); };

    return ui;
  }

  function injectEditorIntoPopup(popup) {
    if (popup.querySelector(':scope > .gpe-host, .gpe-host')) {
      log('Editor already present in this popup');
      return;
    }

    const initial = {
      title: extractTitleFromPopup(popup),
      description: extractDescriptionFromPopup(popup)
    };
    const ui = createEditorUI(initial);

    // Insert near the bottom of the popup content
    popup.appendChild(ui.host);

    ui.on('reload', () => {
      ui.title.value = extractTitleFromPopup(popup) || ui.title.value;
      ui.setStatus('Reloaded from popup');
      setTimeout(() => ui.setStatus(''), 1200);
    });

    ui.on('cancel', () => {
      ui.host.remove();
    });

    ui.on('save', async () => {
      try {
        ui.setSaving(true);
        ui.setStatus('Opening editor…');
        // Snapshot current route and scroll position(s)
        const routeSnap = snapshotRoute();
        const scrollSnap = snapshotCalendarScroll();
        let editBtn = findEditButton(popup);
        if (editBtn) {
          editBtn.click();
        } else {
          // Fallback: try the keyboard shortcut 'e' to open editor
          log('Edit button not found; trying keyboard fallback');
          focusWithin(popup);
          simulateKey('e');
        }

        // Wait for editor title field
        const titleInput = await waitFor(() => findTitleInput(), { timeout: 20000 });

        ui.setStatus('Updating title…');
        setTextInputValue(titleInput, ui.title.value);

        // Save
        ui.setStatus('Saving…');
        const saveBtn = await waitFor(() => findSaveButton(), { timeout: 12000 });
        triggerClick(saveBtn);

        // Non-blocking: auto-accept "送信/Send" prompt if it appears shortly
        const stopPromptWatch = armAutoSendUpdatesPrompt(6000);

        // Wait for Calendar to become idle (loading finished) with short timeout
        await waitForCalendarIdle({ minQuietMs: 350, maxWaitMs: 12000 });
        stopPromptWatch();

        // Restore route (date/view) if changed, then restore scroll — triggered by idle
        ui.setStatus('Restoring view…');
        await restoreRouteSoft(routeSnap);

        // If route still differs (Calendar jumped to today), navigate back hard and restore via sessionStorage
        const after = snapshotRoute();
        if (after !== routeSnap) {
          setPendingRestore({ url: routeSnap, primaryTop: scrollSnap.primaryTop, win: scrollSnap.win, t: Date.now() });
          location.assign(routeSnap);
          return; // further logic will run after navigation via attemptApplyPendingRestore()
        }

        // Same route → just do in-place scroll restore
        await restoreCalendarScrollWithRetries(scrollSnap);
        if (typeof scrollSnap.primaryTop === 'number') {
          lockCalendarScroll(scrollSnap.primaryTop, 1400);
        }
        ui.setSaving(false);
        ui.setStatus('Saved');
        if (typeof ui.pulseCheck === 'function') ui.pulseCheck();
        if (typeof ui.markSaved === 'function') ui.markSaved();
        setTimeout(() => ui.setStatus(''), 1500);
      } catch (e) {
        warn('Save failed', e);
        ui.setSaving(false);
        ui.setStatus(`Error: ${e.message}`);
      }
    });
  }

  function triggerClick(el) {
    el.focus();
    el.click();
  }

  function simulateKey(key) {
    const evInit = { key, code: key.length === 1 ? 'Key' + key.toUpperCase() : key, bubbles: true, cancelable: true };
    document.activeElement?.dispatchEvent(new KeyboardEvent('keydown', evInit));
    document.activeElement?.dispatchEvent(new KeyboardEvent('keypress', evInit));
    document.activeElement?.dispatchEvent(new KeyboardEvent('keyup', evInit));
  }

  function focusWithin(root) {
    const focusableSel = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
    const el = root.querySelector(focusableSel) || root;
    el.focus();
  }

  // --- Scroll position snapshot/restore ------------------------------------
  function isScrollable(el) {
    if (!el || !(el instanceof HTMLElement)) return false;
    const cs = getComputedStyle(el);
    const canY = (cs.overflowY === 'auto' || cs.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
    const canX = (cs.overflowX === 'auto' || cs.overflowX === 'scroll') && el.scrollWidth > el.clientWidth;
    return canX || canY;
  }

  function getScrollableAncestors(start) {
    const list = [];
    let n = start instanceof Node ? start.parentNode : null;
    let hops = 0;
    while (n && n !== document && hops < 10) {
      if (n instanceof HTMLElement && isScrollable(n)) list.push(n);
      n = n.parentNode;
      hops++;
    }
    // Include page scroll element last
    if (document.scrollingElement) list.push(document.scrollingElement);
    return list;
  }

  function snapshotScroll(contextEl) {
    return {
      win: { x: window.scrollX, y: window.scrollY },
      elems: getScrollableAncestors(contextEl).map(el => ({ el, top: el.scrollTop, left: el.scrollLeft }))
    };
  }

  function restoreScrollOnce(snap) {
    try { window.scrollTo(snap.win.x, snap.win.y); } catch {}
    for (const s of snap.elems) {
      const el = s.el;
      if (!el || !document.documentElement.contains(el)) continue;
      try {
        if (typeof el.scrollTo === 'function') el.scrollTo({ left: s.left, top: s.top });
        else { el.scrollLeft = s.left; el.scrollTop = s.top; }
      } catch {}
    }
  }

  async function restoreScrollWithRetries(snap, tries = 6, delay = 60) {
    for (let i = 0; i < tries; i++) {
      restoreScrollOnce(snap);
      await new Promise(r => setTimeout(r, delay));
    }
  }

  // Calendar-specific primary scroller detection (handles overlay remount)
  function findCalendarScrollCandidates() {
    const sels = [
      'main',
      '[role="main"]',
      'div[aria-label*="Calendar" i]',
      'div[aria-label*="カレンダー" i]',
      'div[aria-label*="Main" i]',
      'div[aria-label*="メイン" i]'
    ];
    const set = new Set();
    for (const sel of sels) document.querySelectorAll(sel)?.forEach(el => set.add(el));
    const all = Array.from(set).filter(el => el instanceof HTMLElement && el.offsetParent !== null);
    const scrollables = all.filter(isScrollable);
    // If none matched, consider any large scrollable in document
    const anyScrollables = scrollables.length ? scrollables : Array.from(document.querySelectorAll('div')).filter(isScrollable);
    return anyScrollables;
  }

  function findPrimaryCalendarScroller() {
    const cands = findCalendarScrollCandidates();
    if (!cands.length) return null;
    // Pick the tallest scroll area as the primary calendar scroller
    return cands.reduce((best, el) => {
      const span = (el.scrollHeight - el.clientHeight) + (el.scrollWidth - el.clientWidth);
      return span > ((best?.span) || -1) ? { el, span } : best;
    }, null)?.el || null;
  }

  function snapshotCalendarScroll() {
    const primary = findPrimaryCalendarScroller();
    return { primaryTop: primary ? primary.scrollTop : null, win: { x: window.scrollX, y: window.scrollY } };
  }

  async function restoreCalendarScrollWithRetries(snap, tries = 10, delay = 80) {
    const prev = history.scrollRestoration;
    try { history.scrollRestoration = 'manual'; } catch {}
    for (let i = 0; i < tries; i++) {
      const primary = findPrimaryCalendarScroller();
      if (primary && typeof snap.primaryTop === 'number') {
        try { primary.scrollTop = snap.primaryTop; } catch {}
      }
      try { window.scrollTo(snap.win.x, snap.win.y); } catch {}
      await new Promise(r => setTimeout(r, delay));
    }
    try { history.scrollRestoration = prev; } catch {}
  }

  function lockCalendarScroll(targetTop, ms = 1200) {
    const start = Date.now();
    const handler = () => {
      const el = findPrimaryCalendarScroller();
      if (!el) return;
      if (typeof targetTop === 'number' && Math.abs(el.scrollTop - targetTop) > 2) {
        try { el.scrollTop = targetTop; } catch {}
      }
      if (Date.now() - start >= ms) clearInterval(timer);
    };
    const timer = setInterval(handler, 80);
    handler();
  }

  // --- Route snapshot/restore (keep same date/view) ------------------------
  function snapshotRoute() {
    return location.pathname + location.search + location.hash;
  }

  async function restoreRouteSoft(prevUrl, waitMs = 120) {
    const cur = location.pathname + location.search + location.hash;
    if (cur === prevUrl) return false;
    try {
      history.replaceState(null, '', prevUrl);
      window.dispatchEvent(new PopStateEvent('popstate'));
    } catch {}
    await new Promise(r => setTimeout(r, waitMs));
    return true;
  }

  // Persisted restore across navigation
  const RESTORE_SS_KEY = 'gpe:restore';
  function setPendingRestore(payload) {
    try { sessionStorage.setItem(RESTORE_SS_KEY, JSON.stringify(payload)); } catch {}
  }
  function consumePendingRestore() {
    try {
      const raw = sessionStorage.getItem(RESTORE_SS_KEY);
      if (!raw) return null;
      sessionStorage.removeItem(RESTORE_SS_KEY);
      return JSON.parse(raw);
    } catch { return null; }
  }

  async function attemptApplyPendingRestore() {
    const rec = consumePendingRestore();
    if (!rec) return;
    // Only apply if URL matches what we expected to return to (best effort)
    const cur = snapshotRoute();
    if (rec.url && !cur.includes(rec.url)) {
      // Different route than expected; still try scroll
    }
    const snap = { primaryTop: rec.primaryTop ?? null, win: rec.win || { x: 0, y: 0 } };
    try {
      await waitFor(() => findPrimaryCalendarScroller(), { timeout: 7000, interval: 80 });
    } catch {}
    await restoreCalendarScrollWithRetries(snap, 12, 80);
    if (typeof snap.primaryTop === 'number') lockCalendarScroll(snap.primaryTop, 1600);
  }

  function isEditorOpen() {
    // Look for large dialog with inputs
    const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], div[role="region"]'));
    return dialogs.some(d => d.offsetParent !== null && d.querySelector('input, textarea, [contenteditable="true"]'));
  }

  function isAnyDialogOpen() {
    const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], div[role="region"]'));
    return dialogs.some(d => d.offsetParent !== null);
  }

  function findUpdatePromptDialog() {
    const dlg = Array.from(document.querySelectorAll('div[role="dialog"], div[role="region"]'))
      .find(d => d.offsetParent !== null && /update|send|guest|更新|送信|ゲスト/i.test(d.textContent || ''));
    return dlg || null;
  }

  function chooseInDialog(dlg, prefs = { action: 'send' }) {
    const buttons = Array.from(dlg.querySelectorAll('button, div[role="button"]')).filter(isVisible);
    const byText = (rxList) => buttons.find(b => rxList.some(rx => rx.test((b.textContent || '').trim())));
    if (prefs.action === 'send') {
      const btn = byText([/^send$/i, /^送信$/, /^更新を送信/, /^ゲストに送信/]);
      if (btn) return btn;
    }
    if (prefs.action === 'dontsend') {
      const btn = byText([/don't\s*send/i, /^送信しない$/]);
      if (btn) return btn;
    }
    // Fallback: primary-looking button
    return buttons.find(b => b.getAttribute('data-mdc-dialog-action') === 'accept') || buttons[0] || null;
  }

  async function handleUpdatePromptPreferSend({ timeout = 8000 } = {}) {
    // Wait briefly for the prompt to appear
    let dlg = null;
    try {
      dlg = await waitFor(() => findUpdatePromptDialog(), { timeout, interval: 150 });
    } catch { /* none */ }
    if (!dlg) return false;
    const btn = chooseInDialog(dlg, { action: 'send' });
    if (btn) {
      triggerClick(btn);
      // Wait for the dialog to close
      try { await waitFor(() => !dlg.isConnected || dlg.offsetParent === null, { timeout: 6000 }); } catch {}
      return true;
    }
    return false;
  }

  // Non-blocking watcher: auto-click "送信/Send" if the prompt appears
  function armAutoSendUpdatesPrompt(durationMs = 6000) {
    const mo = new MutationObserver(() => {
      const dlg = findUpdatePromptDialog();
      if (!dlg) return;
      const btn = chooseInDialog(dlg, { action: 'send' });
      if (btn) triggerClick(btn);
    });
    try { mo.observe(document.body, { childList: true, subtree: true }); } catch {}
    const timer = setTimeout(() => mo.disconnect(), durationMs);
    return () => { clearTimeout(timer); mo.disconnect(); };
  }

  function findToastElement() {
    // Look for aria-live alerts, often contain "Undo/元に戻す" or "Saved/保存"
    const cands = Array.from(document.querySelectorAll('[aria-live="polite"], [aria-live="assertive"], [role="alert"]')).filter(isVisible);
    return cands.find(el => /undo|saved|updated|保存|更新|元に戻す/i.test(el.textContent || '')) || null;
  }

  async function waitForSavedToast({ timeout = 15000 } = {}) {
    try {
      const el = await waitFor(() => findToastElement(), { timeout, interval: 150 });
      // Give it a moment to settle
      await new Promise(r => setTimeout(r, 200));
      return !!el;
    } catch { return false; }
  }

  // Detect Calendar loading/busy state and resolve when it becomes idle
  function findVisibleProgressIndicator() {
    const bars = Array.from(document.querySelectorAll('[role="progressbar"], .progress, .loading'))
      .filter(isVisible);
    // Filter out tiny decorative elements
    return bars.find(el => el.getBoundingClientRect().width > 20 || el.getBoundingClientRect().height > 6) || null;
  }

  async function waitForCalendarIdle({ minQuietMs = 400, maxWaitMs = 12000 } = {}) {
    let last = Date.now();
    const mo = new MutationObserver(() => { last = Date.now(); });
    try { mo.observe(document.body, { childList: true, subtree: true, attributes: true, characterData: false }); } catch {}
    const start = Date.now();
    try {
      while (Date.now() - start < maxWaitMs) {
        const busy = isAnyDialogOpen() || !!findVisibleProgressIndicator();
        const quietEnough = (Date.now() - last) >= minQuietMs;
        if (!busy && quietEnough) { mo.disconnect(); return true; }
        await new Promise(r => setTimeout(r, 120));
      }
    } finally { try { mo.disconnect(); } catch {} }
    return false;
  }

  function findTitleInput() {
    // Try aria-label and placeholder in multiple locales
    const all = Array.from(document.querySelectorAll('input[aria-label], input[placeholder]'))
      .filter(isVisible);
    let m = all.find(el => matchesAny(el.getAttribute('aria-label'), LABELS.title));
    if (m) return m;
    m = all.find(el => matchesAny(el.getAttribute('placeholder'), LABELS.title));
    if (m) return m;
    // Fallback: first text input inside an open editor dialog
    return all.find(el => el.type === 'text' && queryClosestDialog(el));
  }

  function findDescriptionBox() {
    // Description is often a contenteditable region, sometimes a textarea
    const cands = [
      ...document.querySelectorAll('[contenteditable="true"][aria-label], [role="textbox"][aria-label]'),
      ...document.querySelectorAll('textarea[aria-label], textarea[placeholder]')
    ].filter(isVisible);
    let match = cands.find(el => matchesAny(el.getAttribute('aria-label'), LABELS.description));
    if (match) return match;
    match = cands.find(el => matchesAny(el.getAttribute('placeholder'), LABELS.description));
    if (match) return match;
    // Fallback: editable area with longest text
    return cands.sort((a, b) => (b.textContent?.length || b.value?.length || 0) - (a.textContent?.length || a.value?.length || 0))[0] || null;
  }

  function findSaveButton() {
    const cands = Array.from(document.querySelectorAll('div[role="button"], button'))
      .filter(isVisible);
    let btn = cands.find(el => matchesAny(el.getAttribute('aria-label'), LABELS.save));
    if (btn) return btn;
    btn = cands.find(el => matchesAny(el.textContent, LABELS.save));
    if (btn) return btn;
    btn = cands.find(el => matchesAny(el.getAttribute('data-tooltip') || el.getAttribute('title'), LABELS.save));
    return btn || null;
  }

  function onMutations(mutations) {
    for (const m of mutations) {
      for (const node of m.addedNodes) {
        if (!(node instanceof HTMLElement)) continue;
        // Look for quick popup
        const dialogs = findQuickPopupDialogs(node);
        dialogs.forEach(d => {
          try { injectEditorIntoPopup(d); } catch (e) { warn('inject failed', e); }
        });
      }
    }
  }

  function boot() {
    // If we just returned to a saved URL, restore scroll ASAP
    attemptApplyPendingRestore();
    // Initial sweep
    findQuickPopupDialogs().forEach(d => {
      try { injectEditorIntoPopup(d); } catch (e) { warn('inject failed', e); }
    });

    // Observe
    const obs = new MutationObserver(onMutations);
    obs.observe(document.body, { childList: true, subtree: true });
    log('Observer attached');
  }

  // Only on Google Calendar
  if (location.hostname.includes('calendar.google.com')) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', boot);
    } else {
      boot();
    }
  }
})();
