// Word Helper Content Script
console.log('[Word Helper] loaded in', window === window.top ? 'main frame' : 'iframe', '|', location.href.slice(0, 60));

let popupEl         = null;
let pillEl          = null;
let _pillWord       = null;
let _pillPos        = null;
let shadowHost      = null;
let shadowRoot      = null;
let lastMouse       = { x: window.innerWidth / 2, y: window.innerHeight / 2 };
let _iHaveSelection = false;
let _lastFocusEl    = null;
let _savedRange     = null;
let _savedSuffix    = '';

document.addEventListener('mousemove', (e) => { lastMouse = { x: e.clientX, y: e.clientY }; }, { passive: true });

if (window !== window.top) {
  document.addEventListener('focusin', (e) => { _lastFocusEl = e.target; }, { capture: true, passive: true });
}


// ── Shadow DOM container ───────────────────────────────────────────────────────
function ensureShadow() {
  if (shadowRoot) return;
  shadowHost = document.createElement('div');
  shadowHost.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:2147483647;pointer-events:none;';
  document.documentElement.appendChild(shadowHost);
  shadowRoot = shadowHost.attachShadow({ mode: 'open' });

  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = chrome.runtime.getURL('content.css');
  shadowRoot.appendChild(link);
}

// Loads content.css into the main document's <head> so liquid-theme popups
// appended to document.body get the same styles as shadow-DOM popups.
function ensureLiquidGlassStylesheet() {
  if (document.getElementById('wh-main-css')) return;
  const link = document.createElement('link');
  link.id  = 'wh-main-css';
  link.rel = 'stylesheet';
  link.href = chrome.runtime.getURL('content.css');
  document.head.appendChild(link);
}

// ── BroadcastChannel: cross-frame coordination ────────────────────────────────
const bc = new BroadcastChannel('word-helper-v1');

bc.addEventListener('message', async (e) => {
  if (e.data.type === 'TRIGGER_COPY') {
    if (window === window.top) return;
    _iHaveSelection = false;
    _savedRange     = null;
    _savedSuffix    = '';

    // getSelection() is empty until execCommand('copy') prods Google Docs into
    // syncing its internal cursor to the DOM — so we still need localCopyRead.
    // The capture-phase preventDefault inside it stops the browser from actually
    // writing to the system clipboard; Google Docs still calls setData() so we
    // can read the text, but the real clipboard is never touched.
    let rawText = window.getSelection()?.toString() ?? '';
    if (!rawText.trim()) rawText = await localCopyRead() ?? '';

    const trimmed = rawText.trim();
    console.log('[Word Helper] TRIGGER_COPY resolved:', JSON.stringify(rawText));

    if (trimmed && /^[a-zA-Z'-]{1,40}$/.test(trimmed)) {
      _iHaveSelection = true;
      // After localCopyRead fires execCommand, Google Docs syncs DOM selection.
      // getSelection() now reflects the actual selection including trailing space.
      const selNow = window.getSelection();
      const rawSel = selNow?.toString() ?? '';
      _savedSuffix = (rawSel || rawText) !== (rawSel || rawText).trimEnd() ? ' ' : '';
      _savedRange  = selNow?.rangeCount > 0 ? selNow.getRangeAt(0).cloneRange() : null;
      window.name  = 'wh-editor-active';
      console.log('[Word Helper] TRIGGER_COPY word:', trimmed, '| suffix:', JSON.stringify(_savedSuffix));
      bc.postMessage({ type: 'FOUND', word: trimmed, pos: e.data.pos, direct: e.data.direct });
    }
  }

  if (e.data.type === 'FOUND' && window === window.top && !popupEl) {
    console.log('[Word Helper] FOUND via broadcast:', e.data.word, '| direct:', !!e.data.direct);
    if (e.data.direct) showPopup(e.data.word, e.data.pos);
    else               showPill(e.data.word, e.data.pos);
  }

  if (e.data.type === 'PASTE_REPLACEMENT' && window !== window.top && _iHaveSelection) {
    _iHaveSelection = false;
    window.name = '';

    const target = document.activeElement;

    // Selection is still "word " (mousedown.preventDefault kept focus in editor).
    const currentSel = window.getSelection()?.toString() ?? '';
    const suffix     = currentSel.endsWith(' ') ? ' ' : _savedSuffix;
    const syn        = matchCase(currentSel.trim(), e.data.syn);
    console.log('[Word Helper] pre-paste sel:', JSON.stringify(currentSel), '| suffix:', JSON.stringify(suffix), '| syn:', syn);

    const dt = new DataTransfer();
    dt.setData('text/plain', syn);
    target.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt }));

    if (suffix === ' ') {
      // Pasting " " via ClipboardEvent gets trimmed; execCommand is intercepted.
      // Dispatching keyboard events mimics the user pressing Space, which Google
      // Docs' own key handler processes — inserting the character into its model.
      await new Promise(r => setTimeout(r, 30));
      for (const type of ['keydown', 'keypress', 'keyup']) {
        target.dispatchEvent(new KeyboardEvent(type, {
          bubbles: true, cancelable: true,
          key: ' ', code: 'Space', keyCode: 32, which: 32,
          ...(type === 'keypress' ? { charCode: 32 } : {}),
        }));
      }
    }
    console.log('[Word Helper] inserted:', JSON.stringify(syn + suffix));
  }
});

// ── Triggers ─────────────────────────────────────────────────────────────────
// Primary trigger is "select, then click the pill": once a single-word
// selection settles, a small glyph appears beside it and opens the card on
// click.  Double-click alone no longer opens the card -- it is Google Docs'
// own select-a-word gesture, so firing on it interrupted ordinary editing.
// Alt+Shift+D and the right-click menu stay "direct": they skip the pill,
// because the user has already expressed the intent explicitly.
const PROBE_DELAY_MS = 180;   // debounce after a selection gesture settles
const DRAG_SLOP_PX   = 3;     // below this, a mouseup is a caret click

let _probeTimer = null;
let _downX = 0, _downY = 0;

window.addEventListener('mouseup',    onMouseUp,   { capture: true });
document.addEventListener('keydown',  onKeyDown,   { capture: true });
document.addEventListener('keyup',    onKeyUp,     { capture: true });
document.addEventListener('mousedown',onMouseDown, { capture: true });
window.addEventListener('scroll',     closePill,   { capture: true, passive: true });
console.log('[Word Helper] listeners ready');

// True when the event originated inside our own pill or card.
function inOurUI(e) {
  const path = e.composedPath?.() ?? [];
  if (pillEl  && path.includes(pillEl))  return true;
  if (popupEl && path.includes(popupEl)) return true;
  return !!popupEl && (!!shadowHost?.contains(e.target) || popupEl.contains(e.target));
}

function scheduleProbe() {
  clearTimeout(_probeTimer);
  _probeTimer = setTimeout(() => triggerLookup(lastMouse, false), PROBE_DELAY_MS);
}

function onMouseUp(e) {
  if (inOurUI(e)) return;
  // A plain caret click selects nothing, so skip the (execCommand-based) probe
  // unless the pointer dragged or this was a double/triple click.
  const moved = Math.hypot(e.clientX - _downX, e.clientY - _downY) >= DRAG_SLOP_PX;
  if (!moved && e.detail < 2) return;
  lastMouse = { x: e.clientX, y: e.clientY };
  scheduleProbe();
}

function onKeyDown(e) {
  if (e.key === 'Escape') { closePopup(); closePill(); return; }
  if (e.altKey && e.shiftKey && !e.ctrlKey && !e.metaKey && e.key.toLowerCase() === 'd') {
    e.preventDefault();
    e.stopPropagation();
    console.log('[Word Helper] Alt+Shift+D pressed');
    closePill();
    triggerLookup(lastMouse, true);
    return;
  }
  // Any other keystroke invalidates the current pill; onKeyUp re-probes if a
  // selection still stands (e.g. the user is extending it with Shift+Arrow).
  closePill();
}

function onKeyUp(e) {
  if (e.key === 'Shift' || (e.shiftKey && e.key.startsWith('Arrow')) ||
      (e.ctrlKey && e.key.toLowerCase() === 'a')) scheduleProbe();
}

function onMouseDown(e) {
  _downX = e.clientX; _downY = e.clientY;
  if (inOurUI(e)) return;
  clearTimeout(_probeTimer);
  closePopup();
  closePill();
}

// ── Core lookup ───────────────────────────────────────────────────────────────
async function triggerLookup(pos, direct = true) {
  const sel     = window.getSelection();
  const rawSel  = sel?.toString() ?? '';
  const trimmed = rawSel.trim();
  console.log('[Word Helper] triggerLookup getSelection:', JSON.stringify(rawSel), '| top:', window === window.top);

  if (trimmed && /^[a-zA-Z'-]{1,40}$/.test(trimmed)) {
    if (window === window.top) {
      if (direct) showPopup(trimmed, pos);
      else        showPill(trimmed, pos);
    } else {
      _iHaveSelection = true;
      _savedSuffix    = rawSel !== rawSel.trimEnd() ? ' ' : '';
      _savedRange     = sel.rangeCount > 0 ? sel.getRangeAt(0).cloneRange() : null;
      window.name     = 'wh-editor-active';
      console.log('[Word Helper] found in iframe getSelection:', trimmed, '| suffix:', JSON.stringify(_savedSuffix));
      bc.postMessage({ type: 'FOUND', word: trimmed, pos, direct });
    }
    return;
  }

  // In an iframe where getSelection() is still empty (Google Docs hasn't synced
  // its cursor to the DOM yet), try a clipboard-safe copy read before giving up.
  if (window !== window.top) {
    const raw = await localCopyRead() ?? '';
    const t   = raw.trim();
    if (t && /^[a-zA-Z'-]{1,40}$/.test(t)) {
      _iHaveSelection = true;
      // After execCommand, Google Docs syncs DOM selection — use that for suffix.
      const selNow = window.getSelection();
      const rawSel = selNow?.toString() ?? '';
      _savedSuffix = (rawSel || raw) !== (rawSel || raw).trimEnd() ? ' ' : '';
      _savedRange  = selNow?.rangeCount > 0 ? selNow.getRangeAt(0).cloneRange() : null;
      window.name  = 'wh-editor-active';
      console.log('[Word Helper] found in iframe localCopyRead:', t, '| suffix:', JSON.stringify(_savedSuffix));
      bc.postMessage({ type: 'FOUND', word: t, pos, direct });
      return;
    }
  }

  console.log('[Word Helper] no word in this frame, broadcasting TRIGGER_COPY');
  bc.postMessage({ type: 'TRIGGER_COPY', pos, direct });
}

// Reads the selected text without touching the system clipboard.
//
// Two-layer approach:
//   1. content-main.js (main world) patches DataTransfer.prototype.setData so
//      that when data-wh-intercept is set, it stores the text in a DOM attribute
//      instead of calling the native C++ setData — leaving the DataTransfer's
//      backing store empty so the browser has nothing to write to the clipboard.
//   2. This isolated-world function sets that flag, fires execCommand('copy'),
//      and also calls e.preventDefault() in window capture phase to block the
//      browser's native-selection fallback.
//   Together: nothing ever reaches the OS clipboard.
function localCopyRead() {
  return new Promise((resolve) => {
    let done = false;

    const cleanup = () => {
      document.documentElement.removeAttribute('data-wh-intercept');
      document.documentElement.removeAttribute('data-wh-copy-text');
    };

    // e.preventDefault() in capture blocks the browser's default "copy native
    // selection to clipboard" path (the fallback when no setData was called).
    const preventWrite = (e) => { e.preventDefault(); };

    const readText = (e) => {
      if (done) return;
      done = true;
      // Text was stored in the DOM attribute by the main-world patch.
      const text = document.documentElement.getAttribute('data-wh-copy-text') ?? '';
      cleanup();
      console.log('[Word Helper] intercepted copy text:', JSON.stringify(text));
      resolve(text || null);
    };

    // Set the intercept flag before registering the copy handler so the
    // main-world patch sees it when Google Docs' handler calls setData().
    document.documentElement.setAttribute('data-wh-intercept', '1');
    document.documentElement.setAttribute('data-wh-copy-text', '');

    window.addEventListener('copy', preventWrite, { capture: true, once: true });
    document.addEventListener('copy', readText,    { capture: false, once: true });

    const ok = document.execCommand('copy');
    console.log('[Word Helper] execCommand(copy):', ok);

    if (!ok) {
      cleanup();
      window.removeEventListener('copy', preventWrite, { capture: true });
      document.removeEventListener('copy', readText);
      done = true;
      resolve(null);
      return;
    }

    setTimeout(() => {
      if (!done) {
        done = true;
        cleanup();
        window.removeEventListener('copy', preventWrite, { capture: true });
        document.removeEventListener('copy', readText);
        resolve(null);
      }
    }, 300);
  });
}

// ── Popup ─────────────────────────────────────────────────────────────────────
// -- Pill --------------------------------------------------------------------
// A 26px affordance shown beside a settled single-word selection.  It lives in
// the shadow root for every theme: it is opaque, so unlike the liquid/frosted
// cards it never needs to see the page behind it through backdrop-filter.
const PILL_SIZE       = 26;
const PILL_HOVER_MS   = 300;   // dwell on the pill to open without clicking

let _pillHoverTimer = null;

function closePill() {
  clearTimeout(_pillHoverTimer);
  pillEl?.remove();
  pillEl    = null;
  _pillWord = null;
  _pillPos  = null;
}

function computePillPos(pos) {
  const GAP = 6;
  let left = pos.x + GAP;
  let top  = pos.y + GAP;
  if (left + PILL_SIZE > window.innerWidth  - 8) left = pos.x - PILL_SIZE - GAP;
  if (top  + PILL_SIZE > window.innerHeight - 8) top  = pos.y - PILL_SIZE - GAP;
  return { top: Math.max(8, top), left: Math.max(8, left) };
}

function showPill(word, pos) {
  if (window !== window.top) return;
  if (popupEl) return;
  // Same word in the same spot: leave the existing pill alone so it does not
  // re-animate every time the probe re-fires on an unchanged selection.
  if (pillEl && _pillWord === word) return;
  closePill();
  ensureShadow();

  _pillWord = word;
  _pillPos  = pos;

  const { top, left } = computePillPos(pos);
  const el = document.createElement('button');
  el.className = 'wh-pill';
  el.type      = 'button';
  el.style.top  = top  + 'px';
  el.style.left = left + 'px';
  el.setAttribute('aria-label', 'Look up "' + word + '"');
  el.title = 'Look up "' + word + '"';
  el.innerHTML =
    '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" ' +
    'stroke="currentColor" stroke-width="2.1" stroke-linecap="round" ' +
    'stroke-linejoin="round" aria-hidden="true">' +
    '<circle cx="10.5" cy="10.5" r="6.5"/><path d="M15.4 15.4 21 21"/></svg>';

  const open = () => {
    const w = _pillWord, p = _pillPos;
    closePill();
    showPopup(w, p);
  };

  // Keep the Google Docs selection alive -- the synonym paste path replaces it,
  // and a focus change here would drop it (same reason the synonym buttons
  // preventDefault on mousedown).
  el.addEventListener('mousedown', e => { e.preventDefault(); e.stopPropagation(); });
  el.addEventListener('click',     e => { e.preventDefault(); e.stopPropagation(); open(); });
  el.addEventListener('mouseenter', () => {
    _pillHoverTimer = setTimeout(open, PILL_HOVER_MS);
  });
  el.addEventListener('mouseleave', () => clearTimeout(_pillHoverTimer));

  shadowRoot.appendChild(el);
  pillEl = el;
  console.log('[Word Helper] pill shown for', word);
}

function closePopup() {
  popupEl?.remove();
  popupEl = null;
}

function computePopupPos(rect) {
  const W = 340, H = 280, GAP = 12;
  let top  = rect.bottom + GAP;
  let left = rect.left;
  if (left + W > window.innerWidth  - 8) left = window.innerWidth  - W - 8;
  if (left < 8) left = 8;
  if (top  + H > window.innerHeight - 8) top  = Math.max(8, rect.top - H - GAP);
  return { top, left };
}

async function showPopup(word, pos) {
  if (window !== window.top) return;
  closePopup();
  closePill();
  ensureShadow();

  const rect = { top: pos.y - 24, bottom: pos.y, left: pos.x, right: pos.x + 10, width: 10, height: 24 };
  const { theme = 'glass' } = await chrome.storage.local.get('theme');

  if (theme === 'liquid' || theme === 'liquidhd' || theme === 'liquid2' || theme === 'frosted') ensureLiquidGlassStylesheet();
  const container = (theme === 'liquid' || theme === 'liquidhd' || theme === 'liquid2' || theme === 'frosted') ? document.body : shadowRoot;

  let screenshotUrl = null;
  if (theme === 'liquid' || theme === 'liquidhd') {
    try {
      const msgType = theme === 'liquidhd' ? 'CAPTURE_SCREEN_HD' : 'CAPTURE_SCREEN';
      const resp = await chrome.runtime.sendMessage({ type: msgType });
      screenshotUrl = resp?.dataUrl ?? null;
      console.log('[Word Helper] screenshot captured:', screenshotUrl ? `${screenshotUrl.length} chars` : 'NULL');
    } catch (e) {
      console.warn('[Word Helper] screenshot failed, using CSS fallback:', e.message);
    }
  }

  popupEl = buildPopup({ state: 'loading', word }, rect, theme, screenshotUrl);
  container.appendChild(popupEl);

  let data;
  try {
    const res = await fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(word)}`);
    if (!res.ok) throw new Error();
    const [entry] = await res.json();
    const meanings = (entry.meanings ?? []).slice(0, 5).map(m => {
      const def    = m.definitions?.[0];
      const synSet = new Set([
        ...(m.synonyms ?? []),
        ...(m.definitions ?? []).flatMap(d => d.synonyms ?? []),
      ]);
      return {
        partOfSpeech: m.partOfSpeech ?? '',
        definition:   def?.definition ?? 'No definition found.',
        example:      def?.example ?? '',
        synonyms:     [...synSet].slice(0, 12),
      };
    });
    data = {
      state: 'loaded', word,
      meanings: meanings.length ? meanings : [{ partOfSpeech: '', definition: 'No definition found.', example: '', synonyms: [] }],
    };
  } catch {
    data = { state: 'loaded', word, meanings: [{ partOfSpeech: '', definition: 'No definition found.', example: '', synonyms: [] }] };
  }

  // Update in-place so the popup never disappears — no DOM remove/re-add, no
  // animation replay, no blank frame between loading state and loaded state.
  if (popupEl) {
    _updatePopupBody(popupEl, data);
  } else {
    popupEl = buildPopup(data, rect, theme, screenshotUrl);
    container.appendChild(popupEl);
  }
}

// Swaps the wh-body content of an existing popup without touching the glass
// layers or triggering a new entry animation.
function _updatePopupBody(el, data) {
  const { word, meanings } = data;
  const multi = meanings.length > 1;

  const meaningHTML = meanings.map((m, i) => `
    <div class="wh-meaning${i === 0 ? '' : ' wh-hidden'}" data-idx="${i}">
      <div class="wh-def">${esc(m.definition)}</div>
      ${m.example ? `<div class="wh-ex">&ldquo;${esc(m.example)}&rdquo;</div>` : ''}
      <div class="wh-divider"></div>
      <div class="wh-syn-label">Synonyms</div>
      ${m.synonyms.length
        ? `<div class="wh-syn-list">${m.synonyms.map(s => `<button class="wh-syn" data-syn="${escAttr(s)}">${esc(s)}</button>`).join('')}</div>`
        : `<div class="wh-no-syn">No synonyms found.</div>`}
    </div>`).join('');

  const body = el.querySelector('.wh-body');
  if (!body) return;

  body.innerHTML = `
    <div class="wh-header">
      <span class="wh-word">${esc(word)}</span>
      ${!multi && meanings[0].partOfSpeech ? `<span class="wh-pos">${esc(meanings[0].partOfSpeech)}</span>` : ''}
      <button class="wh-close" aria-label="Close">&#x2715;</button>
    </div>
    ${multi ? `<div class="wh-tabs">${meanings.map((m, i) =>
      `<button class="wh-tab${i === 0 ? ' wh-tab-active' : ''}" data-idx="${i}">${esc(m.partOfSpeech)}</button>`
    ).join('')}</div>` : ''}
    ${meaningHTML}`;

  el.querySelector('.wh-close').addEventListener('click', closePopup);

  for (const tab of el.querySelectorAll('.wh-tab')) {
    tab.addEventListener('mousedown', e => e.preventDefault());
    tab.addEventListener('click', () => {
      const idx = tab.dataset.idx;
      el.querySelectorAll('.wh-tab').forEach(t => t.classList.remove('wh-tab-active'));
      el.querySelectorAll('.wh-meaning').forEach(m => m.classList.add('wh-hidden'));
      tab.classList.add('wh-tab-active');
      el.querySelector(`.wh-meaning[data-idx="${idx}"]`).classList.remove('wh-hidden');
    });
  }

  for (const btn of el.querySelectorAll('.wh-syn')) {
    btn.addEventListener('mousedown', e => e.preventDefault());
    btn.addEventListener('click', () => {
      bc.postMessage({ type: 'PASTE_REPLACEMENT', syn: btn.dataset.syn });
      closePopup();
    });
  }
}

function buildPopup(data, rect, theme = 'glass', screenshotUrl = null) {
  const el = document.createElement('div');
  el.className = 'wh-popup';
  if (theme === 'dictionary') el.classList.add('wh-theme-dictionary');

  const { top, left } = computePopupPos(rect);
  el.style.top  = `${top}px`;
  el.style.left = `${left}px`;

  const layers = `
    <div class="wh-glass-filter"></div>
    <div class="wh-glass-overlay"></div>
    <div class="wh-glass-specular"></div>`;

  if (data.state === 'loading') {
    el.innerHTML = layers + `
      <div class="wh-body">
        <div class="wh-loading"><span class="wh-spinner"></span>Looking up <strong class="wh-loading-word">${esc(data.word)}</strong></div>
      </div>`;
  } else {
    const { word, meanings } = data;
    const multi = meanings.length > 1;

    const meaningHTML = meanings.map((m, i) => `
      <div class="wh-meaning${i === 0 ? '' : ' wh-hidden'}" data-idx="${i}">
        <div class="wh-def">${esc(m.definition)}</div>
        ${m.example ? `<div class="wh-ex">&ldquo;${esc(m.example)}&rdquo;</div>` : ''}
        <div class="wh-divider"></div>
        <div class="wh-syn-label">Synonyms</div>
        ${m.synonyms.length
          ? `<div class="wh-syn-list">${m.synonyms.map(s => `<button class="wh-syn" data-syn="${escAttr(s)}">${esc(s)}</button>`).join('')}</div>`
          : `<div class="wh-no-syn">No synonyms found.</div>`}
      </div>`).join('');

    el.innerHTML = layers + `
      <div class="wh-body">
        <div class="wh-header">
          <span class="wh-word">${esc(word)}</span>
          ${!multi && meanings[0].partOfSpeech ? `<span class="wh-pos">${esc(meanings[0].partOfSpeech)}</span>` : ''}
          <button class="wh-close" aria-label="Close">&#x2715;</button>
        </div>
        ${multi ? `<div class="wh-tabs">${meanings.map((m, i) =>
          `<button class="wh-tab${i === 0 ? ' wh-tab-active' : ''}" data-idx="${i}">${esc(m.partOfSpeech)}</button>`
        ).join('')}</div>` : ''}
        ${meaningHTML}
      </div>`;

    el.querySelector('.wh-close').addEventListener('click', closePopup);

    for (const tab of el.querySelectorAll('.wh-tab')) {
      tab.addEventListener('mousedown', e => e.preventDefault());
      tab.addEventListener('click', () => {
        const idx = tab.dataset.idx;
        el.querySelectorAll('.wh-tab').forEach(t => t.classList.remove('wh-tab-active'));
        el.querySelectorAll('.wh-meaning').forEach(m => m.classList.add('wh-hidden'));
        tab.classList.add('wh-tab-active');
        el.querySelector(`.wh-meaning[data-idx="${idx}"]`).classList.remove('wh-hidden');
      });
    }

    for (const btn of el.querySelectorAll('.wh-syn')) {
      // preventDefault on mousedown keeps the editor element focused —
      // without this, clicking the button moves focus out of the editor iframe
      // and Google Docs loses its internal cursor before we can paste.
      btn.addEventListener('mousedown', e => e.preventDefault());
      btn.addEventListener('click', () => {
        bc.postMessage({ type: 'PASTE_REPLACEMENT', syn: btn.dataset.syn });
        closePopup();
      });
    }
  }

  if (theme === 'frosted') {
    el.classList.add('wh-theme-frosted');
    // backdrop-filter must be on the popup element itself — isolation:isolate on
    // the popup means any child's backdrop-filter only sees content inside the
    // popup, not the page behind it. Applying it to el directly bypasses that.
    el.style.backdropFilter = 'blur(6px)';
    el.style.webkitBackdropFilter = 'blur(6px)';
    el.style.background = 'rgba(255,255,255,0.35)';
    el.querySelector('.wh-glass-filter').style.display = 'none';
    el.querySelector('.wh-glass-overlay').style.display = 'none';
  }

  if (theme === 'liquid2') {
    // Pure CSS — no screenshot needed. Matches the bubbbly.com panel technique:
    // backdrop-filter on the popup element itself (not a child) so isolation:isolate
    // doesn't block it from seeing the page behind the popup.
    el.classList.add('wh-theme-liquid2');
    el.style.backdropFilter = 'blur(12px) saturate(5) brightness(1.03)';
    el.style.webkitBackdropFilter = 'blur(12px) saturate(5) brightness(1.03)';
    el.style.background = 'rgba(255, 255, 255, 0.30)';
    el.querySelector('.wh-glass-filter').style.display = 'none';
    el.querySelector('.wh-glass-overlay').style.display = 'none';
  }

  if (theme === 'liquid' || theme === 'liquidhd') {
    try {
      el.classList.add('wh-theme-liquid');
      const filterEl  = el.querySelector('.wh-glass-filter');
      const overlayEl = el.querySelector('.wh-glass-overlay');

      if (screenshotUrl) {
        // ── WebGL path: real lens refraction from page screenshot ─────────────
        const { top: pTop, left: pLeft } = computePopupPos(rect);
        if (theme === 'liquidhd') {
          _buildWebGLGlassHD(el, screenshotUrl, pLeft, pTop);
        } else {
          _buildWebGLGlass(el, screenshotUrl, pLeft, pTop);
        }
        // Hide CSS overlay and specular — the WebGL shader provides both the
        // frosted background and the glare/Fresnel highlight.
        // Add wh-webgl class so the ::after pseudo-element (radial gradients
        // centred outside the popup that bleed in as triangular wedges) is also
        // suppressed via CSS .wh-theme-liquid.wh-webgl::after { display:none }.
        overlayEl.style.display = 'none';
        el.querySelector('.wh-glass-specular').style.display = 'none';
        el.classList.add('wh-webgl');
      } else {
        // ── CSS fallback: frosted glass + gentle surface-ripple overlay ────────
        _ensureLiquidGlassDom();
        filterEl.style.backdropFilter = 'blur(26px) saturate(1.9) brightness(1.05)';
        filterEl.style.webkitBackdropFilter = 'blur(26px) saturate(1.9) brightness(1.05)';
        overlayEl.style.background =
          'linear-gradient(158deg,rgba(225,238,255,0.18) 0%,rgba(255,255,255,0.10) 48%,rgba(218,235,255,0.16) 100%)';
        overlayEl.style.filter = 'url("#wh-lg")';
        overlayEl.style.webkitFilter = 'url("#wh-lg")';
      }
    } catch (err) {
      console.error('[Word Helper] liquid glass init error:', err);
    }
  }

  return el;
}


chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'WH_LOOKUP') { closePill(); triggerLookup(lastMouse, true); }
});

// ── Liquid Glass SVG filter ────────────────────────────────────────────────────
// backdrop-filter: url() doesn't propagate SVG displacement to the backdrop in
// Chrome — blur works but feDisplacementMap is silently ignored there.
// Instead we apply filter: url() to the semi-transparent color overlay div,
// making the gradient itself ripple and flow.  The frosted glass underneath
// comes from a plain backdrop-filter: blur() which is fully supported.
function _ensureLiquidGlassDom() {
  if (document.getElementById('wh-lg-svg')) return;
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.id = 'wh-lg-svg';
  svg.setAttribute('style', 'position:fixed;width:0;height:0;pointer-events:none;overflow:hidden;');
  // Low-frequency, slow turbulence — displaces the near-white tint layer gently
  // so the glass surface ripples like undisturbed water, not like a lava lamp.
  svg.innerHTML =
    `<defs>` +
    `<filter id="wh-lg" color-interpolation-filters="sRGB" x="-20%" y="-20%" width="140%" height="140%">` +
    `<feTurbulence type="fractalNoise" baseFrequency="0.016 0.012" numOctaves="3" seed="5" result="noise">` +
    `<animate attributeName="baseFrequency" calcMode="spline" keyTimes="0;0.5;1" keySplines="0.45 0 0.55 1;0.45 0 0.55 1" values="0.016 0.012;0.022 0.016;0.016 0.012" dur="14s" repeatCount="indefinite"/>` +
    `</feTurbulence>` +
    `<feDisplacementMap in="SourceGraphic" in2="noise" scale="20" xChannelSelector="R" yChannelSelector="G"/>` +
    `</filter>` +
    `</defs>`;
  document.body.appendChild(svg);
}

// ── WebGL liquid-glass renderer ───────────────────────────────────────────────
// Sets up a WebGL canvas inside the popup's glass-filter layer.
// The fragment shader implements: lens-refraction at edges + multi-tap frost blur
// + saturation boost + specular highlight — all sampled from a page screenshot.
function _buildWebGLGlass(popupEl, screenshotUrl, cssLeft, cssTop) {
  const PW = 340, PH = 280;

  const canvas = document.createElement('canvas');
  canvas.width  = PW;
  canvas.height = PH;
  // Must fill the popup; z-index 0 puts it below overlay/specular/body layers
  canvas.style.cssText =
    'position:absolute;top:0;left:0;width:100%;height:100%;' +
    'border-radius:inherit;pointer-events:none;z-index:0;display:block;';

  const filterEl = popupEl.querySelector('.wh-glass-filter');
  filterEl.style.cssText += ';backdrop-filter:none;-webkit-backdrop-filter:none;';
  filterEl.appendChild(canvas);

  console.log('[Word Helper] WebGL glass: starting, url length =', screenshotUrl?.length);

  const img = new Image();
  img.onerror = (e) => console.error('[Word Helper] WebGL screenshot img failed:', e);
  img.onload = () => {
    console.log('[Word Helper] WebGL glass: img loaded', img.naturalWidth, 'x', img.naturalHeight);
    const gl = canvas.getContext('webgl', { alpha: false, antialias: false });
    if (!gl) { console.error('[Word Helper] WebGL context unavailable'); return; }

    const SW = window.innerWidth, SH = window.innerHeight;
    // Popup rect in normalised screen UV (0–1), accounting for UNPACK_FLIP_Y
    const rx = cssLeft / SW;
    const ry = cssTop  / SH;
    const rw = PW / SW;
    const rh = PH / SH;

    // ── shaders ───────────────────────────────────────────────────────────────
    const vs = gl.createShader(gl.VERTEX_SHADER);
    gl.shaderSource(vs,
      'attribute vec2 a;varying vec2 v;' +
      'void main(){v=a*.5+.5;gl_Position=vec4(a,0.,1.);}');
    gl.compileShader(vs);

    const fs = gl.createShader(gl.FRAGMENT_SHADER);
    gl.shaderSource(fs, `
      precision mediump float;
      uniform sampler2D u;
      uniform vec4 r;   /* x=leftUV y=topUV z=widthUV w=heightUV */
      varying vec2 v;   /* (0,0)=canvas bottom-left  (1,1)=canvas top-right */
      #define PI 3.14159265359

      float sdRB(vec2 p, vec2 b, float rad) {
        vec2 q = abs(p) - b + rad;
        return length(max(q,0.)) + min(max(q.x,q.y),0.) - rad;
      }

      void main() {
        vec2 base = vec2(r.x + v.x*r.z, r.y + (1.-v.y)*r.w);
        vec2 c    = v - .5;
        vec2 p    = vec2(c.x*1.214, c.y);

        /* Pixel-space SDF — exactly matches the CSS 340×280 popup with border-radius:20px.
           d_px < 0 inside, > 0 outside (in the corner cuts), 0 on the rim.
           Using pixel space avoids the old "margin" bug where an undersized SDF
           caused maximum-refraction artefacts along the top/bottom edge strips. */
        vec2  px   = c * vec2(340., 280.);
        float d_px = sdRB(px, vec2(170.,140.), 20.);
        vec2 nrm   = normalize(vec2(
          sdRB(px+vec2(1.,0.),vec2(170.,140.),20.) - sdRB(px-vec2(1.,0.),vec2(170.,140.),20.),
          sdRB(px+vec2(0.,1.),vec2(170.,140.),20.) - sdRB(px-vec2(0.,1.),vec2(170.,140.),20.)
        ));

        /* ── Snell's law refraction (liquid-glass-studio STEP 3-9)
           IOR=1.50, 28 px rim band.  step(d_px,0.)=1 inside, 0 outside —
           prevents the corner-area pixels from getting wrong maximum refraction. */
        float nPx = max(0., -d_px);
        float xR  = clamp(1.-nPx/28., 0., 1.);
        float thI = asin(pow(xR, 2.));
        float eF  = step(d_px, 0.) * max(0., -tan(asin(clamp(sin(thI)/1.50,-1.,1.)) - thI));

        /* ── Per-channel chromatic dispersion (N_R=0.965, N_G=1.0, N_B=1.035) */
        vec2 ks  = vec2(r.z/1.214, -r.w);
        vec2 bv  = -nrm*eF*(20./280.);
        vec2 uvR = clamp(base + bv*ks*1.20, 0., 1.);
        vec2 uvG = clamp(base + bv*ks,       0., 1.);
        vec2 uvB = clamp(base + bv*ks*0.80, 0., 1.);

        /* ── Two-zone frost blur — per-channel for dispersion */
        float centreBlend = smoothstep(.15,.45,length(p));
        float rimMask     = smoothstep(-16.,0.,d_px);
        float br  = mix(.010,.002, max(centreBlend, rimMask));
        vec4 sR=vec4(0.), sG=vec4(0.), sB=vec4(0.);
        for (int i=0; i<16; i++) {
          float a  = float(i)*.3927;
          float ri = (mod(float(i),2.)==0.) ? br : br*.55;
          vec2  os = vec2(cos(a),sin(a))*ri;
          sR += texture2D(u, uvR+os);
          sG += texture2D(u, uvG+os);
          sB += texture2D(u, uvB+os);
        }
        vec4 col = vec4(sR.r/16., sG.g/16., sB.b/16., 1.);

        /* Saturation boost */
        float lum = dot(col.rgb, vec3(.299,.587,.114));
        col.rgb = mix(vec3(lum), col.rgb, 1.6)*1.05;

        /* Very light center tint */
        float center = 1.-centreBlend;
        col.rgb = mix(col.rgb, vec3(.96,.97,1.), center*.14);
        col.rgb = mix(col.rgb, vec3(.85,.93,1.), .02+center*.04);

        /* ── Fresnel: bright 20 px rim glow ─────────────────────────────────
           rimPx = 1 at the rim (d_px=0), fades to 0 at 20 px inside.
           NOTE: d_px < 0 inside, so the correct formula is (1 + d_px/20).
           clamp(-d_px/20) would be inverted: 0 at rim, 1 deep inside. */
        float rimPx = clamp(1. + d_px/20., 0., 1.);
        col.rgb = mix(col.rgb, vec3(1.), rimPx*rimPx*0.72);

        /* ── Directional glare (liquid-glass-studio sine-angle, upper-left peak) */
        float nrmAngle = atan(nrm.y, nrm.x);
        if (nrmAngle < 0.) nrmAngle += 2.*PI;
        float glareAngle = (nrmAngle - PI/4.) * 2.;
        float glareAng   = clamp(pow((0.5+sin(glareAngle)*0.5)*1.2*0.85, 1.1), 0., 1.);
        float glareGeo   = clamp(pow(1.+d_px/1500.*pow(500./30.,2.)+0.25, 5.), 0., 1.);
        col.rgb += (glareAng*glareGeo + rimPx*rimPx*glareAng*0.7)*0.95;

        col.rgb = clamp(col.rgb, 0., 1.);
        gl_FragColor = vec4(col.rgb, 1.);
      }
    `);
    gl.compileShader(fs);

    const prog = gl.createProgram();
    gl.attachShader(prog, vs);
    gl.attachShader(prog, fs);
    gl.linkProgram(prog);
    gl.useProgram(prog);

    // Full-screen quad (-1,-1) to (1,1)
    const buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(gl.ARRAY_BUFFER,
      new Float32Array([-1,-1, 1,-1, -1,1, -1,1, 1,-1, 1,1]),
      gl.STATIC_DRAW);
    const aLoc = gl.getAttribLocation(prog, 'a');
    gl.enableVertexAttribArray(aLoc);
    gl.vertexAttribPointer(aLoc, 2, gl.FLOAT, false, 0, 0);

    // Screenshot texture — no UNPACK_FLIP_Y; the formula r.y+(1-v.y)*r.w
    // already accounts for WebGL's y-origin being at the bottom.
    const tex = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, tex);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, img);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);

    gl.uniform1i(gl.getUniformLocation(prog, 'u'), 0);
    gl.uniform4f(gl.getUniformLocation(prog, 'r'), rx, ry, rw, rh);
    gl.viewport(0, 0, PW, PH);
    gl.drawArrays(gl.TRIANGLES, 0, 6);
    const err = gl.getError();
    console.log('[Word Helper] WebGL glass: draw complete, GL error =', err, '| rect UV:', rx.toFixed(3), ry.toFixed(3), rw.toFixed(3), rh.toFixed(3));
  };
  img.src = screenshotUrl;
}

// ── WebGL liquid-glass renderer — HiDPI variant ───────────────────────────────
// Identical effect to _buildWebGLGlass but renders at physical pixel resolution
// (canvas.width = PW * devicePixelRatio) and captures a quality-100 JPEG.
// All pixel-space shader constants are scaled by u_dpr so the geometry is
// identical to the standard version — just sharper on HiDPI screens.
function _buildWebGLGlassHD(popupEl, screenshotUrl, cssLeft, cssTop) {
  const PW = 340, PH = 280;
  const dpr = window.devicePixelRatio || 1;
  const CW  = Math.round(PW * dpr);
  const CH  = Math.round(PH * dpr);

  const canvas = document.createElement('canvas');
  canvas.width  = CW;
  canvas.height = CH;
  canvas.style.cssText =
    'position:absolute;top:0;left:0;width:100%;height:100%;' +
    'border-radius:inherit;pointer-events:none;z-index:0;display:block;';

  const filterEl = popupEl.querySelector('.wh-glass-filter');
  filterEl.style.cssText += ';backdrop-filter:none;-webkit-backdrop-filter:none;';
  filterEl.appendChild(canvas);

  const img = new Image();
  img.onerror = (e) => console.error('[Word Helper] WebGL HD screenshot img failed:', e);
  img.onload = () => {
    const gl = canvas.getContext('webgl', { alpha: false, antialias: false });
    if (!gl) { console.error('[Word Helper] WebGL HD context unavailable'); return; }

    const SW = window.innerWidth, SH = window.innerHeight;
    const rx = cssLeft / SW, ry = cssTop / SH;
    const rw = PW / SW,     rh = PH / SH;

    const vs = gl.createShader(gl.VERTEX_SHADER);
    gl.shaderSource(vs,
      'attribute vec2 a;varying vec2 v;' +
      'void main(){v=a*.5+.5;gl_Position=vec4(a,0.,1.);}');
    gl.compileShader(vs);

    const fs = gl.createShader(gl.FRAGMENT_SHADER);
    gl.shaderSource(fs, `
      precision highp float;
      uniform sampler2D u;
      uniform vec4 r;
      uniform float u_dpr;
      varying vec2 v;
      #define PI 3.14159265359

      float sdRB(vec2 p, vec2 b, float rad) {
        vec2 q = abs(p) - b + rad;
        return length(max(q,0.)) + min(max(q.x,q.y),0.) - rad;
      }

      void main() {
        vec2 base = vec2(r.x + v.x*r.z, r.y + (1.-v.y)*r.w);
        vec2 c    = v - .5;
        vec2 p    = vec2(c.x*1.214, c.y);

        vec2  px   = c * vec2(340.*u_dpr, 280.*u_dpr);
        float d_px = sdRB(px, vec2(170.*u_dpr, 140.*u_dpr), 20.*u_dpr);
        vec2 nrm   = normalize(vec2(
          sdRB(px+vec2(1.,0.),vec2(170.*u_dpr,140.*u_dpr),20.*u_dpr) - sdRB(px-vec2(1.,0.),vec2(170.*u_dpr,140.*u_dpr),20.*u_dpr),
          sdRB(px+vec2(0.,1.),vec2(170.*u_dpr,140.*u_dpr),20.*u_dpr) - sdRB(px-vec2(0.,1.),vec2(170.*u_dpr,140.*u_dpr),20.*u_dpr)
        ));

        float nPx = max(0., -d_px);
        float xR  = clamp(1.-nPx/(28.*u_dpr), 0., 1.);
        float thI = asin(pow(xR, 2.));
        float eF  = step(d_px, 0.) * max(0., -tan(asin(clamp(sin(thI)/1.50,-1.,1.)) - thI));

        vec2 ks  = vec2(r.z/1.214, -r.w);
        vec2 bv  = -nrm*eF*(20./280.);
        vec2 uvR = clamp(base + bv*ks*1.20, 0., 1.);
        vec2 uvG = clamp(base + bv*ks,       0., 1.);
        vec2 uvB = clamp(base + bv*ks*0.80, 0., 1.);

        float centreBlend = smoothstep(.15,.45,length(p));
        float rimMask     = smoothstep(-16.*u_dpr, 0., d_px);
        float br  = mix(.010,.002, max(centreBlend, rimMask));
        vec4 sR=vec4(0.), sG=vec4(0.), sB=vec4(0.);
        for (int i=0; i<16; i++) {
          float a  = float(i)*.3927;
          float ri = (mod(float(i),2.)==0.) ? br : br*.55;
          vec2  os = vec2(cos(a),sin(a))*ri;
          sR += texture2D(u, uvR+os);
          sG += texture2D(u, uvG+os);
          sB += texture2D(u, uvB+os);
        }
        vec4 col = vec4(sR.r/16., sG.g/16., sB.b/16., 1.);

        float lum = dot(col.rgb, vec3(.299,.587,.114));
        col.rgb = mix(vec3(lum), col.rgb, 1.6)*1.05;

        float center = 1.-centreBlend;
        col.rgb = mix(col.rgb, vec3(.96,.97,1.), center*.14);
        col.rgb = mix(col.rgb, vec3(.85,.93,1.), .02+center*.04);

        float rimPx = clamp(1. + d_px/(20.*u_dpr), 0., 1.);
        col.rgb = mix(col.rgb, vec3(1.), rimPx*rimPx*0.72);

        float nrmAngle = atan(nrm.y, nrm.x);
        if (nrmAngle < 0.) nrmAngle += 2.*PI;
        float glareAngle = (nrmAngle - PI/4.) * 2.;
        float glareAng   = clamp(pow((0.5+sin(glareAngle)*0.5)*1.2*0.85, 1.1), 0., 1.);
        float glareGeo   = clamp(pow(1.+d_px/(1500.*u_dpr)*pow(500./30.,2.)+0.25, 5.), 0., 1.);
        col.rgb += (glareAng*glareGeo + rimPx*rimPx*glareAng*0.7)*0.95;

        col.rgb = clamp(col.rgb, 0., 1.);
        gl_FragColor = vec4(col.rgb, 1.);
      }
    `);
    gl.compileShader(fs);

    const prog = gl.createProgram();
    gl.attachShader(prog, vs);
    gl.attachShader(prog, fs);
    gl.linkProgram(prog);
    gl.useProgram(prog);

    const buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(gl.ARRAY_BUFFER,
      new Float32Array([-1,-1, 1,-1, -1,1, -1,1, 1,-1, 1,1]),
      gl.STATIC_DRAW);
    const aLoc = gl.getAttribLocation(prog, 'a');
    gl.enableVertexAttribArray(aLoc);
    gl.vertexAttribPointer(aLoc, 2, gl.FLOAT, false, 0, 0);

    const tex = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, tex);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, img);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);

    gl.uniform1i(gl.getUniformLocation(prog, 'u'), 0);
    gl.uniform4f(gl.getUniformLocation(prog, 'r'), rx, ry, rw, rh);
    gl.uniform1f(gl.getUniformLocation(prog, 'u_dpr'), dpr);
    gl.viewport(0, 0, CW, CH);
    gl.drawArrays(gl.TRIANGLES, 0, 6);
  };
  img.src = screenshotUrl;
}

function matchCase(original, syn) {
  if (!original || !syn) return syn;
  const isLetter = c => c.toLowerCase() !== c.toUpperCase();
  const firstLetter = [...original].find(isLetter) ?? '';
  if (!firstLetter) return syn;
  if (original.split('').filter(isLetter).every(c => c === c.toUpperCase())) {
    return syn.toUpperCase();
  }
  if (firstLetter === firstLetter.toUpperCase()) {
    return syn[0].toUpperCase() + syn.slice(1);
  }
  return syn;
}

const esc     = s => (s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const escAttr = s => (s ?? '').replace(/"/g,'&quot;');
