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
const DRAG_SLOP_PX   = 3;     // below this, a pointerup is a caret click

let _probeTimer = null;
let _downX = 0, _downY = 0;

// Pointer events, not mouse events.  When a page cancels pointerdown, the
// browser stops sending that press's compatibility mousedown/mouseup at all --
// and Google Docs' editor manages its own selection, so the mouse events these
// triggers used to rely on never arrived in edit mode.  pointerdown/up always
// fire.  They carry no click count, though, so double/triple-click selection is
// caught from 'click', which also still fires after a cancelled pointerdown.
window.addEventListener('pointerdown', onPointerDown, { capture: true });
window.addEventListener('pointerup',   onPointerUp,   { capture: true });
window.addEventListener('click',       onClick,       { capture: true });
document.addEventListener('keydown',   onKeyDown,     { capture: true });
document.addEventListener('keyup',     onKeyUp,       { capture: true });
window.addEventListener('scroll',      closePill,     { capture: true, passive: true });
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

// A plain caret click selects nothing, so the (execCommand-based) probe only
// runs after a drag or a multi-click.  Both paths can fire for one gesture;
// scheduleProbe's debounce collapses them into a single lookup.
function onPointerUp(e) {
  if (!e.isPrimary || e.button !== 0 || inOurUI(e)) return;
  if (Math.hypot(e.clientX - _downX, e.clientY - _downY) < DRAG_SLOP_PX) return;
  lastMouse = { x: e.clientX, y: e.clientY };
  scheduleProbe();
}

function onClick(e) {
  if (e.detail < 2 || inOurUI(e)) return;
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

function onPointerDown(e) {
  if (!e.isPrimary) return;
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
// True while showPopup is awaiting storage and popupEl is still
// null.  Without it, releasing Shift after Alt+Shift+D fires a selection probe
// whose pill lands on top of the card that is already on its way.
let _popupPending = false;

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
  if (popupEl || _popupPending) return;
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
    const r = el.getBoundingClientRect();
    const origin = { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    // Hand the pill to its launch animation.  Clearing pillEl first keeps the
    // closePill() inside showPopup from yanking it out mid-flight.
    clearTimeout(_pillHoverTimer);
    pillEl = null; _pillWord = null; _pillPos = null;
    el.classList.add('wh-pill-launch');
    setTimeout(() => el.remove(), 300);
    showPopup(w, p, origin);
  };

  // Keep the Google Docs selection alive -- the synonym paste path replaces it,
  // and a focus change here would drop it (same reason the synonym buttons
  // preventDefault on mousedown).
  el.addEventListener('pointerdown', e => { e.preventDefault(); e.stopPropagation(); });
  el.addEventListener('mousedown',   e => { e.preventDefault(); e.stopPropagation(); });
  el.addEventListener('click',       e => { e.preventDefault(); e.stopPropagation(); open(); });
  el.addEventListener('mouseenter', () => {
    _pillHoverTimer = setTimeout(open, PILL_HOVER_MS);
  });
  el.addEventListener('mouseleave', () => clearTimeout(_pillHoverTimer));

  shadowRoot.appendChild(el);
  pillEl = el;
  console.log('[Word Helper] pill shown for', word);
}

// Close with a short shrink-and-fade rather than vanishing.  The element is
// detached from popupEl first, so a new card can open while this one leaves.
const LEAVE_MS = 150;

function closePopup() {
  const el = popupEl;
  popupEl = null;
  if (!el) return;
  el.classList.add('wh-leaving');
  // animationend bubbles up from the card's children too, so match the card.
  el.addEventListener('animationend', e => { if (e.target === el) el.remove(); });
  // Fallback: animationend never fires if the animation is skipped.
  setTimeout(() => el.remove(), LEAVE_MS + 250);
}

const POPUP_W = 340, POPUP_H = 280;   // must match .wh-popup in content.css

function computePopupPos(rect) {
  const W = POPUP_W, H = POPUP_H, GAP = 12;
  let top  = rect.bottom + GAP;
  let left = rect.left;
  if (left + W > window.innerWidth  - 8) left = window.innerWidth  - W - 8;
  if (left < 8) left = 8;
  if (top  + H > window.innerHeight - 8) top  = Math.max(8, rect.top - H - GAP);
  return { top, left };
}

// ── Dictionary lookup ─────────────────────────────────────────────────────────
// dictionaryapi.dev is primary (it has example sentences); Datamuse is the
// fallback.  The primary's host does go down -- it has been seen returning
// Cloudflare 522s after ~20s -- so it gets a short timeout, and after a failure
// it is skipped for a while so only the first lookup of an outage pays for it.
const DICT_TIMEOUT_MS    = 3000;
const PRIMARY_BACKOFF_MS = 5 * 60 * 1000;
let _primaryDownUntil    = 0;

const noEntry = () => Object.assign(new Error('no entry'), { noEntry: true });

async function lookupWord(word) {
  if (Date.now() >= _primaryDownUntil) {
    try {
      return await _lookupPrimary(word);
    } catch (err) {
      if (!err.noEntry) {
        _primaryDownUntil = Date.now() + PRIMARY_BACKOFF_MS;
        console.warn('[Word Helper] dictionaryapi.dev failed, using Datamuse:', err.name, err.message);
      }
      // A 404 falls through too: Datamuse's coverage is wider.
    }
  }
  return _lookupDatamuse(word);
}

async function _lookupPrimary(word) {
  const res = await fetch(
    `https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(word)}`,
    { signal: AbortSignal.timeout(DICT_TIMEOUT_MS) });
  if (res.status === 404) throw noEntry();
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
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
  if (!meanings.length) throw noEntry();
  return meanings;
}

const DATAMUSE_POS = { n: 'noun', v: 'verb', adj: 'adjective', adv: 'adverb' };

async function _lookupDatamuse(word) {
  const q      = encodeURIComponent(word);
  const signal = AbortSignal.timeout(DICT_TIMEOUT_MS);
  const [defRes, synRes] = await Promise.all([
    fetch(`https://api.datamuse.com/words?sp=${q}&md=dp&max=1`, { signal }),
    fetch(`https://api.datamuse.com/words?rel_syn=${q}&md=p&max=40`, { signal }),
  ]);
  if (!defRes.ok || !synRes.ok) throw new Error(`Datamuse HTTP ${defRes.status}/${synRes.status}`);
  const [entry] = await defRes.json();
  const syns    = await synRes.json();
  const lower   = word.toLowerCase();
  // sp= is a spelling match, so a typo comes back as a different word.
  if (!entry || entry.word.toLowerCase() !== lower || !entry.defs?.length) throw noEntry();

  // defs arrive as "n	An institution ..."; keep the first per part of speech.
  const byPos = new Map();
  for (const raw of entry.defs) {
    const tab = raw.indexOf('	');
    const tag = raw.slice(0, tab);
    if (!byPos.has(tag)) byPos.set(tag, raw.slice(tab + 1).trim());
  }
  return [...byPos].slice(0, 5).map(([tag, definition]) => ({
    partOfSpeech: DATAMUSE_POS[tag] ?? '',
    definition,
    example: '',
    // Synonyms are tagged by part of speech, so "educate" lands under the verb
    // and "academy" under the noun.  Drop derivatives like "schoolhouse".
    synonyms: syns
      .filter(sy => (sy.tags ?? []).includes(tag) && !sy.word.toLowerCase().includes(lower))
      .map(sy => sy.word)
      .slice(0, 12),
  }));
}

// Themes whose card mounts in document.body instead of the shadow root, so a
// backdrop-filter on the card can see the page behind it.
const BODY_THEMES = new Set(['liquidlive', 'frosted']);

// Retired themes, mapped to their successor so a saved choice keeps working.
const LEGACY_THEMES = { liquid: 'liquidlive', liquidhd: 'liquidlive', liquid2: 'liquidlive' };

// origin: viewport point the card grows out of -- the pill's centre when opened
// from the pill, otherwise the pointer.
async function showPopup(word, pos, origin = pos) {
  if (window !== window.top) return;
  closePopup();
  closePill();
  ensureShadow();
  _popupPending = true;
  try {
    await _showPopupInner(word, pos, origin);
  } finally {
    _popupPending = false;
  }
}

async function _showPopupInner(word, pos, origin) {
  const rect = { top: pos.y - 24, bottom: pos.y, left: pos.x, right: pos.x + 10, width: 10, height: 24 };
  const { theme: stored = 'glass' } = await chrome.storage.local.get('theme');
  const theme = LEGACY_THEMES[stored] ?? stored;

  const onBody = BODY_THEMES.has(theme);
  if (onBody) ensureLiquidGlassStylesheet();
  const container = onBody ? document.body : shadowRoot;

  popupEl = buildPopup({ state: 'loading', word }, rect, theme, origin);
  container.appendChild(popupEl);

  let data;
  try {
    data = { state: 'loaded', word, meanings: await lookupWord(word) };
  } catch (err) {
    // noEntry means both sources answered and neither knows the word.  Any
    // other error is the services failing -- saying "No definition found"
    // then would tell the user their word doesn't exist.
    if (!err.noEntry) console.warn('[Word Helper] all dictionary sources failed:', err.name, err.message);
    data = {
      state: 'loaded', word,
      meanings: [{
        partOfSpeech: '', example: '', synonyms: [],
        definition: err.noEntry
          ? 'No definition found.'
          : "Couldn't reach the dictionary service. Try again in a moment.",
      }],
    };
  }

  // Update in-place so the popup never disappears — no DOM remove/re-add, no
  // animation replay, no blank frame between loading state and loaded state.
  if (popupEl) {
    _updatePopupBody(popupEl, data);
  } else {
    popupEl = buildPopup(data, rect, theme, origin);
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

  // Restart the swap animation so the definition fades in over the spinner.
  body.classList.remove('wh-body-swap');
  void body.offsetWidth;
  body.classList.add('wh-body-swap');

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

function buildPopup(data, rect, theme = 'glass', origin = null) {
  const el = document.createElement('div');
  el.className = 'wh-popup';
  if (theme === 'dictionary') el.classList.add('wh-theme-dictionary');

  const { top, left } = computePopupPos(rect);
  el.style.top  = `${top}px`;
  el.style.left = `${left}px`;

  if (theme === 'dictionary' && origin) {
    // Unroll like a scroll: a thin rolled bar grows out from under the pill,
    // then the sheet unrolls away from it.  Anchor the card on the sides
    // nearest the pill, so the growth starts there whichever way
    // computePopupPos flipped the card to fit on screen.
    el.classList.add('wh-unfurl');
    if (origin.y > top + POPUP_H / 2) {
      el.classList.add('wh-unfurl-up');
      el.style.top    = 'auto';
      el.style.bottom = `${window.innerHeight - top - POPUP_H}px`;
    }
    if (origin.x > left + POPUP_W / 2) {
      el.classList.add('wh-unfurl-right');
      el.style.left  = 'auto';
      el.style.right = `${window.innerWidth - left - POPUP_W}px`;
    }
  } else {
    // Spring out of the point the user acted on, so the card reads as coming
    // from the pill rather than appearing beside it.
    el.classList.add('wh-spring');
    if (origin) el.style.transformOrigin = `${origin.x - left}px ${origin.y - top}px`;
  }

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

  if (el.classList.contains('wh-unfurl')) {
    // The paper roller riding the unrolling edge.  A sibling of .wh-body, so it
    // survives _updatePopupBody swapping the body's contents.
    el.insertAdjacentHTML('beforeend', '<div class="wh-scroll-rod" aria-hidden="true"></div>');
  }

  if (theme === 'liquidlive') {
    // The browser refracts the live page through the backdrop filter.  It goes
    // on the card itself -- isolation:isolate makes the card a backdrop root,
    // so on a child it would only ever see the card's own interior.
    el.classList.add('wh-theme-liquidlive');
    _lgEnsureFilter();
    el.style.backdropFilter = 'url(#wh-lg-live-f)';
    el.querySelector('.wh-glass-filter').style.display  = 'none';
    el.querySelector('.wh-glass-overlay').style.display = 'none';
  }

  return el;
}


chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'WH_LOOKUP') { closePill(); triggerLookup(lastMouse, true); }
});

// ── Liquid glass via SVG backdrop filter ──────────────────────
// Chromium lets backdrop-filter reference an SVG filter, so the browser itself
// refracts the live page behind the card: no page capture, no WebGL texture,
// and no broad host permission.  The displacement map encodes a rounded-rect bezel:
// flat (0.5 grey = no shift) across the face, bending toward the rim.
const LG_W = 340, LG_H = 280, LG_R = 20;   // must match .wh-popup in CSS
const LG_BEZEL = 28;                        // px of curved rim
const LG_SCALE = 40;                        // feDisplacementMap scale; max shift is half this
const LG_BLUR  = 2.5;                       // frost on the face: enough to keep card text legible
const LG_CHROMA = 0.2;                      // R/B displace this much more/less than G
const LG_SAT   = 1.1;                       // saturation boost; higher turns the rim's colour split
                                            // into hard stripes along table rules

function _lgSdf(px, py) {
  // Signed distance to the rounded rect, negative inside.
  const qx = Math.abs(px - LG_W / 2) - (LG_W / 2 - LG_R);
  const qy = Math.abs(py - LG_H / 2) - (LG_H / 2 - LG_R);
  const ox = Math.max(qx, 0), oy = Math.max(qy, 0);
  return Math.hypot(ox, oy) + Math.min(Math.max(qx, qy), 0) - LG_R;
}

function _lgDisplacementMap(bezel = LG_BEZEL) {
  const c   = document.createElement('canvas');
  c.width   = LG_W; c.height = LG_H;
  const ctx = c.getContext('2d');
  const img = ctx.createImageData(LG_W, LG_H);
  const d   = img.data;
  for (let y = 0; y < LG_H; y++) {
    for (let x = 0; x < LG_W; x++) {
      const px = x + 0.5, py = y + 0.5;
      const s  = _lgSdf(px, py);
      let dx = 0, dy = 0;
      if (s < 0 && s > -bezel) {
        // Outward normal from the SDF gradient.
        const gx = _lgSdf(px + 0.5, py) - _lgSdf(px - 0.5, py);
        const gy = _lgSdf(px, py + 0.5) - _lgSdf(px, py - 0.5);
        const gl = Math.hypot(gx, gy) || 1;
        // 0 at the rim, 1 where the bezel meets the flat face.  A circular
        // bezel's slope is steepest at the rim, so refraction is too.
        const t   = -s / bezel;
        const mag = Math.pow(1 - t, 2);
        // Sample inward (toward the centre): a convex rim magnifies, and it
        // never asks for backdrop pixels outside the card's own box, which the
        // filter region would clip.
        dx = -(gx / gl) * mag;
        dy = -(gy / gl) * mag;
      }
      const i = (y * LG_W + x) * 4;
      d[i]     = Math.round(128 + 127 * dx);
      d[i + 1] = Math.round(128 + 127 * dy);
      d[i + 2] = 128;
      d[i + 3] = 255;
    }
  }
  ctx.putImageData(img, 0, 0);
  return c.toDataURL('image/png');
}

function _lgEnsureFilter({ id = 'wh-lg-live', scale = LG_SCALE, blur = LG_BLUR,
                          bezel = LG_BEZEL, chroma = LG_CHROMA, sat = LG_SAT,
                          doc = document } = {}) {
  if (doc.getElementById(id)) return;
  const map = _lgDisplacementMap(bezel);
  const disp = (scale, res) =>
    `<feDisplacementMap in="frost" in2="map" scale="${scale}" ` +
    `xChannelSelector="R" yChannelSelector="G" result="${res}"/>`;
  const only = (src, row, res) => {
    const rows = ['0 0 0 0 0', '0 0 0 0 0', '0 0 0 0 0'];
    rows[row] = ['1 0 0 0 0', '0 1 0 0 0', '0 0 1 0 0'][row];
    return `<feColorMatrix in="${src}" type="matrix" values="${rows.join(' ')} 0 0 0 1 0" result="${res}"/>`;
  };
  const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.id = id;
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('style', 'position:fixed;width:0;height:0;pointer-events:none;overflow:hidden;');
  // Chromatic aberration: red refracts harder than green, blue softer.
  svg.innerHTML =
    `<filter id="${id}-f" x="0" y="0" width="${LG_W}" height="${LG_H}" ` +
    `filterUnits="userSpaceOnUse" primitiveUnits="userSpaceOnUse" ` +
    `color-interpolation-filters="sRGB">` +
      `<feImage href="${map}" x="0" y="0" width="${LG_W}" height="${LG_H}" ` +
      `preserveAspectRatio="none" result="map"/>` +
      `<feGaussianBlur in="SourceGraphic" stdDeviation="${blur}" result="frost"/>` +
      disp(scale * (1 + chroma), 'dr') + disp(scale, 'dg') + disp(scale * (1 - chroma), 'db') +
      only('dr', 0, 'r') + only('dg', 1, 'g') + only('db', 2, 'b') +
      `<feComposite in="r"  in2="g" operator="arithmetic" k2="1" k3="1" result="rg"/>` +
      `<feComposite in="rg" in2="b" operator="arithmetic" k2="1" k3="1" result="rgb"/>` +
      `<feColorMatrix in="rgb" type="saturate" values="${sat}"/>` +
    `</filter>`;
  doc.body.appendChild(svg);
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
