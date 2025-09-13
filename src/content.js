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
        const saveBtn = await waitFor(() => findSaveButton(), { timeout: 15000 });
        triggerClick(saveBtn);

        // Wait until editor closes
        await waitFor(() => !isEditorOpen(), { timeout: 20000 });
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

  function isEditorOpen() {
    // Look for large dialog with inputs
    const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], div[role="region"]'));
    return dialogs.some(d => d.offsetParent !== null && d.querySelector('input, textarea, [contenteditable="true"]'));
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
