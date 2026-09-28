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

// The pill and card live in the shadow root, styled by a <link> that loads
// asynchronously.  Created lazily on the first pill, the first pill rendered
// unstyled -- static, in the top-left corner -- until the sheet arrived.  So
// the top frame sets it up as soon as there is a document to attach to.
if (window === window.top) {
  if (document.documentElement) ensureShadow();
  else document.addEventListener('readystatechange', ensureShadow, { once: true });
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

    // Ask Docs for its current selection first (localCopyRead).  This frame's
    // DOM selection is only synced when Docs is prodded, so after one lookup it
    // can keep holding the previous word -- every later pill then offered that
    // word.  The capture-phase preventDefault inside localCopyRead stops the
    // browser from writing to the system clipboard; Docs still calls setData()
    // so the text can be read, but the real clipboard is never touched.
    let rawText = await localCopyRead() ?? '';
    if (!rawText.trim()) rawText = window.getSelection()?.toString() ?? '';

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

  if (e.data.type === 'CLOSE' && window === window.top) { closePopup(); closePill(); }

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
const PROBE_DELAY_MS = 120;   // debounce after a selection gesture settles
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
  if (e.key === 'Escape') {
    closePopup(); closePill();
    // In Docs the key lands in the hidden text frame, but the card lives in
    // the top frame.
    if (window !== window.top) bc.postMessage({ type: 'CLOSE' });
    return;
  }
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
  const isWord = t => t && /^[a-zA-Z'-]{1,40}$/.test(t);

  if (window === window.top) {
    const rawSel  = window.getSelection()?.toString() ?? '';
    const trimmed = rawSel.trim();
    console.log('[Word Helper] triggerLookup getSelection:', JSON.stringify(rawSel), '| top: true');
    if (isWord(trimmed)) {
      if (direct) showPopup(trimmed, pos);
      else        showPill(trimmed, pos);
      return;
    }
    console.log('[Word Helper] no word in this frame, broadcasting TRIGGER_COPY');
    bc.postMessage({ type: 'TRIGGER_COPY', pos, direct });
    return;
  }

  // In a frame (Docs' hidden text frame): ask Docs for its selection first --
  // this frame's DOM selection can still hold the previous word (see the
  // TRIGGER_COPY handler) -- and fall back to the DOM selection.
  const raw = (await localCopyRead()) || window.getSelection()?.toString() || '';
  const t   = raw.trim();
  if (isWord(t)) {
    _iHaveSelection = true;
    // After execCommand, Google Docs syncs DOM selection — use that for suffix.
    const selNow = window.getSelection();
    const rawSel = selNow?.toString() ?? '';
    _savedSuffix = (rawSel || raw) !== (rawSel || raw).trimEnd() ? ' ' : '';
    _savedRange  = selNow?.rangeCount > 0 ? selNow.getRangeAt(0).cloneRange() : null;
    window.name  = 'wh-editor-active';
    console.log('[Word Helper] found in frame:', t, '| suffix:', JSON.stringify(_savedSuffix));
    bc.postMessage({ type: 'FOUND', word: t, pos, direct });
    return;
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
  lookupCached(word);                          // prefetch: see lookupCached
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
// dictionaryapi.dev is primary.  Its host does go down -- it has been seen
// returning Cloudflare 522s after ~20s -- so it gets a short timeout, and after
// a failure it is skipped for a while so only the first lookup of an outage
// pays for it.
//
// The fallback is Wiktionary, which dictionaryapi.dev is itself built from:
// one page per spelling (so "grad" never picks up the "Grad" rocket launcher),
// senses in Wiktionary's order, and example sentences.  Synonyms come from the
// same sense's synonym line where editors have written one, topped up from
// Datamuse.  Datamuse's own definitions are the last resort: they merge letter
// cases, come in no useful order, and keep editorial labels like
// "(India, Canada, US)".
const DICT_TIMEOUT_MS    = 3000;
const PRIMARY_BACKOFF_MS = 5 * 60 * 1000;
let _primaryDownUntil    = 0;

const noEntry = () => Object.assign(new Error('no entry'), { noEntry: true });

// Wiktionary first: it is Wikimedia infrastructure, the tuned sense and
// synonym handling lives in that path, and it is the largest of the three
// sources -- dictionaryapi.dev is built from a subset of it -- so its "no such
// word" is final.  The others only step in when Wiktionary can't be reached.
async function lookupWord(word) {
  try {
    return await _lookupWiktionary(word);
  } catch (err) {
    if (err.noEntry) throw err;
    console.warn('[Word Helper] Wiktionary failed:', err.name, err.message);
  }
  // The fallbacks race rather than queue: dictionaryapi.dev, when down, hangs
  // until its timeout, and Datamuse answers in a couple of hundred ms.
  const primary = Date.now() < _primaryDownUntil
    ? Promise.reject(new Error('backed off'))
    : _lookupPrimary(word).catch(err => {
        if (!err.noEntry) {
          _primaryDownUntil = Date.now() + PRIMARY_BACKOFF_MS;
          console.warn('[Word Helper] dictionaryapi.dev failed:', err.name, err.message);
        }
        throw err;
      });
  try {
    return await Promise.any([primary, _lookupDatamuse(word)]);
  } catch (agg) {
    throw agg.errors.find(e => !e.noEntry) ?? agg.errors[0];
  }
}

// ── Lookup cache and prefetch ─────────────────────────────────────────────────
// A lookup starts the moment the pill appears, not when it is clicked:
// reaching for the pill takes a few hundred milliseconds, which covers the
// network, so the card usually opens with its content already there.  Results
// are also kept across tabs and reloads, so a word seen before is instant.
const LOOKUP_STORE_PREFIX = 'lk3:';            // bump when result shapes change
const LOOKUP_STORE_MAX    = 400;
const LOOKUP_TTL_MS       = 14 * 24 * 60 * 60 * 1000;
const _lookups = new Map();                     // word -> { promise, settled, meanings, error }

function lookupCached(word) {
  const key = word.toLowerCase();
  const hit = _lookups.get(key);
  // A failed lookup is retried; "no such word" is remembered.
  if (hit && !(hit.settled && hit.error && !hit.error.noEntry)) return hit;

  const entry = { promise: null, settled: false, meanings: null, error: null };
  entry.promise = (async () => {
    const stored = await _readStoredLookup(key);
    if (stored) return stored;
    const meanings = await lookupWord(word);
    _writeStoredLookup(key, meanings);
    return meanings;
  })().then(m => { entry.meanings = m; entry.settled = true; return m; },
            e => { entry.error = e;    entry.settled = true; throw e; });
  entry.promise.catch(() => {});                // failures are read from entry.error
  _lookups.set(key, entry);
  if (_lookups.size > 100) _lookups.delete(_lookups.keys().next().value);
  return entry;
}

async function _readStoredLookup(key) {
  try {
    const id = LOOKUP_STORE_PREFIX + key;
    const rec = (await chrome.storage.local.get(id))[id];
    return rec && Date.now() - rec.t < LOOKUP_TTL_MS ? rec.m : null;
  } catch { return null; }                       // storage is an optimisation only
}

async function _writeStoredLookup(key, meanings) {
  try {
    const id = LOOKUP_STORE_PREFIX + key;
    const { lkIndex = [] } = await chrome.storage.local.get('lkIndex');
    const index   = lkIndex.filter(k => k !== id).concat(id);
    const evicted = index.splice(0, Math.max(0, index.length - LOOKUP_STORE_MAX));
    await chrome.storage.local.set({ [id]: { t: Date.now(), m: meanings }, lkIndex: index });
    if (evicted.length) await chrome.storage.local.remove(evicted);
  } catch { /* storage is an optimisation only */ }
}

// The card's data for a settled lookup.
function _cardData(word, entry) {
  if (entry.meanings) return { state: 'loaded', word, meanings: entry.meanings };
  const err = entry.error ?? {};
  // noEntry means the sources answered and none knows the word.  Any other
  // error is the services failing -- "No definition found" would then tell the
  // user their word doesn't exist.
  if (!err.noEntry) console.warn('[Word Helper] all dictionary sources failed:', err.name, err.message);
  return {
    state: 'loaded', word,
    meanings: [{
      partOfSpeech: '', example: '', synonyms: [],
      definition: err.noEntry
        ? 'No definition found.'
        : "Couldn't reach the dictionary service. Try again in a moment.",
    }],
  };
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

// ── Wiktionary fallback ───────────────────────────────────────────────────────
// No custom headers (Wikimedia suggests an Api-User-Agent, but only
// suggests): any non-simple header makes each request a CORS preflight, and
// preflights are cached per URL -- so every new word would pay an extra round
// trip before the real one.
const WIKT_TIMEOUT_MS = 2500;   // a stalled request falls back instead of hanging the card
const MAX_SYNONYMS    = 8;

// Parts of speech worth a tab.  Leaves out "Symbol" (big: an ISO language
// code), "Letter", "Prefix" -- and "Proper noun", since lookups are lowercased.
const WORD_POS = new Set(['noun', 'verb', 'adjective', 'adverb', 'pronoun', 'preposition',
                          'conjunction', 'interjection', 'determiner', 'article', 'numeral', 'particle']);
// Synonyms carrying these qualifiers read as wrong in modern prose.
const DATED_SYNONYM = /archaic|obsolete|dated|dialect|rare|poetic|nonstandard|regional|scotland|northern|slang|vulgar/i;

// One lookup can reach the same URL more than once ("decisions" follows its
// noun and verb senses to "decision"), so JSON responses are kept briefly.
const _jsonCache = new Map();
function _cachedJson(url, init, check) {
  if (_jsonCache.has(url)) return _jsonCache.get(url);
  const p = fetch(url, init).then(check);
  p.catch(() => _jsonCache.delete(url));
  _jsonCache.set(url, p);
  if (_jsonCache.size > 60) _jsonCache.delete(_jsonCache.keys().next().value);
  return p;
}

function _wiktFetch(url) {
  return _cachedJson(url, { signal: AbortSignal.timeout(WIKT_TIMEOUT_MS) }, res => {
    if (res.status === 404) throw noEntry();
    if (!res.ok) throw new Error(`Wiktionary HTTP ${res.status}`);
    return res.json();
  });
}

// Everything a lookup needs comes from the page source, in one request:
// senses in Wiktionary's order, their text, example sentences, synonym lines,
// and whether a sense only points at another entry ("past of run", "clipping
// of graduate").  (The REST definition endpoint renders pages on demand and
// took 0.5-2s for words outside Wikimedia's cache.)  The source is read with
// action=query&prop=revisions, which returns it from storage as-is:
// action=parse&prop=wikitext returns the same text but still runs the parser,
// and measured 0.4-2.3s on first fetch against 0.14-0.31s for this.
// Returns [{ pos, senses: [{ plain, example, syns, formOf }] }, ...], or throws
// noEntry when the page or its English section doesn't exist.
async function _wiktSource(title) {
  const data = await _wiktFetch('https://en.wiktionary.org/w/api.php?action=query&prop=revisions' +
    '&rvprop=content&rvslots=main&redirects=1&format=json&formatversion=2&origin=*' +
    `&titles=${encodeURIComponent(title)}`);
  if (data.error) throw new Error(`Wiktionary API ${data.error.code}`);
  const pg = data.query?.pages?.[0];
  if (!pg || pg.missing || pg.invalid) throw noEntry();
  const text = pg.revisions?.[0]?.slots?.main?.content ?? '';
  // Leading newline: on pages like "schools" the English section is line one.
  const en   = (('\n' + text).split(/\n==English==\n/)[1] ?? '').split(/\n==[^=]/)[0];
  if (!en) throw noEntry();
  const blocks = [];
  const re = /\n(===+)\s*([A-Za-z][A-Za-z ]*?)\s*\1\n([\s\S]*?)(?=\n===|$)/g;
  for (let m; (m = re.exec(en)); ) {
    const senses = [];
    for (const line of m[3].split('\n')) {
      // A top-level sense: "#" not followed by ":", "*" or "#" (a space is optional).
      if (/^#(?![:*#])/.test(line)) {
        senses.push({ plain: _wikiPlain(line.slice(1)), example: '', syns: [], formOf: _parseFormOf(line) });
        continue;
      }
      const cur = senses[senses.length - 1];
      if (!cur) continue;
      if (/^#:\s*\{\{(?:syn|synonyms)\|en\|/.test(line)) cur.syns.push(..._parseSynTemplate(line));
      else if (!cur.example && /^#:\s*\{\{(?:ux|uxi|usex)\|en\|/.test(line)) cur.example = _wikiExample(line);
    }
    blocks.push({ pos: m[2].toLowerCase(), senses });
  }
  return blocks;
}

// "#: {{ux|en|Our children attend a public '''school''' nearby.}}" -> the sentence
function _wikiExample(line) {
  const inner = line.replace(/^#:\s*\{\{(?:ux|uxi|usex)\|en\|/, '').replace(/\}\}\s*$/, '');
  const text  = _wikiPlain(inner.split(/\|(?:t|translation|q|qq|ref|inline)=/)[0]);
  return text.length <= 140 ? text : '';
}

// Wikitext -> readable text: "{{lb|en|informal}} Of great size, [[large]]." ->
// "Of great size, large."  Labels, qualifiers and sense ids are dropped (that
// is what kept "(India, Canada, US)" off the card); templates that display a
// word keep it.
const _ENTITIES = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', '#39': "'", ndash: '–', mdash: '—', hellip: '…' };
function _wikiPlain(w) {
  let t = w.replace(/<!--[\s\S]*?-->/g, '')
           .replace(/<ref[^>]*\/>/g, '').replace(/<ref[^>]*>[\s\S]*?<\/ref>/g, '');
  // {{1|rapidly}} renders "Rapidly": case templates keep their word, recased.
  t = t.replace(/\{\{(1|cap|ucfirst|lcfirst)\|([^{}|]*)\}\}/g, (_, fn, word) =>
    fn === 'lcfirst' ? word.charAt(0).toLowerCase() + word.slice(1)
                     : word.charAt(0).toUpperCase() + word.slice(1));
  // {{l|en|expedition}}, {{w|Paris}}, {{n-g|Used to ...}}: keep the displayed text.
  t = t.replace(/\{\{(?:l|m|ll|l-self|w|gloss|vern|taxlink|glossary|n-g|ngd|non-gloss definition|non-gloss)\|([^{}]*)\}\}/g, (_, args) => {
    const pos = args.split('|').filter(x => !x.includes('='));
    return pos.length > 1 && /^[a-z]{2,3}(-[a-z]+)?$/.test(pos[0]) ? pos.at(-1) : pos[0] ?? '';
  });
  let prev;
  do { prev = t; t = t.replace(/\{\{[^{}]*\}\}/g, ''); } while (t !== prev);
  return t.replace(/\[\[(?:[^|\]]*\|)?([^\]]*)\]\]/g, '$1').replace(/<[^>]+>/g, '')
          .replace(/'{2,}/g, '').replace(/&(#39|[a-z]+);/g, (m, e) => _ENTITIES[e] ?? m)
          .replace(/\s+([,.;:])/g, '$1').replace(/\s+/g, ' ').trim();
}

// {{past of|en|run}}, {{infl of|en|run||past}}, {{clipping of|en|graduate}} ...
function _parseFormOf(line) {
  const m = line.match(/\{\{(?:en-)?([a-z][a-z -]*?) of\|(?:en\|)?([^|}<]+)([^}]*)\}\}/i);
  if (!m) return null;
  const kind   = m[1].toLowerCase();
  const target = m[2].replace(/#.*$/, '').trim();              // "graduate#Noun" -> "graduate"
  if (/misspelling|misconstruction|eye dialect|pronunciation spelling/.test(kind)) return { target, type: 'misspelling' };
  if (/clipping|abbreviation|initialism|acronym|short|contraction|ellipsis/.test(kind))  return { target, type: 'short' };
  if (/alternative|alt|obsolete|archaic|dated|nonstandard|rare|standard spelling|letter-case/.test(kind))
    return { target, type: 'alt', caseOnly: /letter-case/.test(kind) };
  return { target, type: 'inflection', form: _inflectionForm(kind + ' ' + m[3].toLowerCase()) };
}

// Which form a pointer names: "simple past", "infl of ... ||3|s|pres", ...
function _inflectionForm(t) {
  if (/plural|\bp\b|\|p\|/.test(t) && !/past|pres/.test(t)) return 'plural';
  if (/comparative|\bcomp\b/.test(t))                          return 'comparative';
  if (/superlative|\bsupd?\b/.test(t))                          return 'superlative';
  if (/present participle|ing form|gerund|pres\|ptcp|\bing\b/.test(t)) return 'ing';
  if (/third-person|3\|s|\b3s\b/.test(t))                     return '3sg';
  if (/past participle/.test(t) && !/simple past|past tense/.test(t)) return 'pp';
  if (/past|\bspast\b/.test(t))                                 return 'past';
  return null;
}

// ── Inflecting synonyms ───────────────────────────────────────────────────────
// A synonym has to match the selected word's form to be pasted in its place:
// "ran" needs "sprinted", not "sprint"; "schools" needs "academies".  Regular
// English rules, plus the irregular verbs common enough to turn up as synonyms.
const IRREGULAR = {
  arise:['arose','arisen'], be:['was','been'], bear:['bore','borne'], beat:['beat','beaten'], become:['became','become'],
  begin:['began','begun'], bend:['bent','bent'], bind:['bound','bound'], bite:['bit','bitten'], bleed:['bled','bled'],
  blow:['blew','blown'], break:['broke','broken'], breed:['bred','bred'], bring:['brought','brought'], build:['built','built'],
  burst:['burst','burst'], buy:['bought','bought'], cast:['cast','cast'], catch:['caught','caught'], choose:['chose','chosen'],
  cling:['clung','clung'], come:['came','come'], cost:['cost','cost'], creep:['crept','crept'], cut:['cut','cut'],
  deal:['dealt','dealt'], dig:['dug','dug'], do:['did','done'], draw:['drew','drawn'], drink:['drank','drunk'],
  drive:['drove','driven'], eat:['ate','eaten'], fall:['fell','fallen'], feed:['fed','fed'], feel:['felt','felt'],
  fight:['fought','fought'], find:['found','found'], flee:['fled','fled'], fling:['flung','flung'], fly:['flew','flown'],
  forbid:['forbade','forbidden'], forget:['forgot','forgotten'], forgive:['forgave','forgiven'], freeze:['froze','frozen'],
  get:['got','gotten'], give:['gave','given'], go:['went','gone'], grind:['ground','ground'], grow:['grew','grown'],
  hang:['hung','hung'], have:['had','had'], hear:['heard','heard'], hide:['hid','hidden'], hit:['hit','hit'],
  hold:['held','held'], hurt:['hurt','hurt'], keep:['kept','kept'], kneel:['knelt','knelt'], know:['knew','known'],
  lay:['laid','laid'], lead:['led','led'], leap:['leapt','leapt'], leave:['left','left'], lend:['lent','lent'],
  let:['let','let'], lie:['lay','lain'], light:['lit','lit'], lose:['lost','lost'], make:['made','made'],
  mean:['meant','meant'], meet:['met','met'], pay:['paid','paid'], put:['put','put'], quit:['quit','quit'],
  read:['read','read'], ride:['rode','ridden'], ring:['rang','rung'], rise:['rose','risen'], run:['ran','run'],
  say:['said','said'], see:['saw','seen'], seek:['sought','sought'], sell:['sold','sold'], send:['sent','sent'],
  set:['set','set'], shake:['shook','shaken'], shine:['shone','shone'], shoot:['shot','shot'], show:['showed','shown'],
  shrink:['shrank','shrunk'], shut:['shut','shut'], sing:['sang','sung'], sink:['sank','sunk'], sit:['sat','sat'],
  sleep:['slept','slept'], slide:['slid','slid'], sling:['slung','slung'], speak:['spoke','spoken'], speed:['sped','sped'],
  spend:['spent','spent'], spin:['spun','spun'], split:['split','split'], spread:['spread','spread'], spring:['sprang','sprung'],
  stand:['stood','stood'], steal:['stole','stolen'], stick:['stuck','stuck'], sting:['stung','stung'], stride:['strode','stridden'],
  strike:['struck','struck'], strive:['strove','striven'], swear:['swore','sworn'], sweep:['swept','swept'], swim:['swam','swum'],
  swing:['swung','swung'], take:['took','taken'], teach:['taught','taught'], tear:['tore','torn'], tell:['told','told'],
  think:['thought','thought'], throw:['threw','thrown'], thrust:['thrust','thrust'], tread:['trod','trodden'],
  understand:['understood','understood'], wake:['woke','woken'], wear:['wore','worn'], weave:['wove','woven'],
  weep:['wept','wept'], win:['won','won'], wind:['wound','wound'], wring:['wrung','wrung'], write:['wrote','written'],
};
const IRREGULAR_PLURAL = { man:'men', woman:'women', child:'children', person:'people', foot:'feet', tooth:'teeth',
  mouse:'mice', goose:'geese', ox:'oxen', criterion:'criteria', phenomenon:'phenomena', analysis:'analyses',
  crisis:'crises', thesis:'theses', life:'lives', knife:'knives', wife:'wives', leaf:'leaves', half:'halves',
  wolf:'wolves', self:'selves', shelf:'shelves', thief:'thieves', loaf:'loaves' };

const _cons  = c => /[bcdfghjklmnpqrstvwxz]/.test(c);
// stop -> stopped, plan -> planned; one short syllable ending consonant-vowel-consonant.
const _doubles = w => /^[^aeiou]*[aeiou][bdgklmnprtv]$/.test(w);
const _sForm = w =>
  /(s|x|z|ch|sh)$/.test(w) ? w + 'es' : (/y$/.test(w) && _cons(w.at(-2)) ? w.slice(0, -1) + 'ies' : w + 's');
const _edForm = w =>
  /e$/.test(w) ? w + 'd' : (/y$/.test(w) && _cons(w.at(-2)) ? w.slice(0, -1) + 'ied'
                         : (_doubles(w) ? w + w.at(-1) + 'ed' : w + 'ed'));
const _ingForm = w =>
  /ie$/.test(w) ? w.slice(0, -2) + 'ying' : (/[^eoy]e$/.test(w) ? w.slice(0, -1) + 'ing'
                                          : (_doubles(w) ? w + w.at(-1) + 'ing' : w + 'ing'));
// glad -> gladder, happy -> happier, but elated -> more elated.  One syllable,
// or two ending in -y/-le/-er/-ow, take the suffix; anything longer takes
// "more"/"most".
const _syllables = w => Math.max(1, (w.match(/[aeiouy]+/g) ?? []).length - (/[^l]e$/.test(w) ? 1 : 0));
const _erForm = (w, suf) => {
  const n = _syllables(w);
  if (n > 2 || (n === 2 && !/(y|le|er|ow)$/.test(w))) return (suf === 'er' ? 'more ' : 'most ') + w;
  if (/e$/.test(w)) return w + suf.slice(1);
  if (/y$/.test(w) && _cons(w.at(-2))) return w.slice(0, -1) + 'i' + suf;
  return _doubles(w) ? w + w.at(-1) + suf : w + suf;
};

// "overtake" conjugates like "take": a known prefix on an irregular verb.
function _irregular(w) {
  if (IRREGULAR[w]) return IRREGULAR[w];
  const m = w.match(/^(over|under|out|re|mis|with|fore|up|be|for)(.+)$/);
  const base = m && IRREGULAR[m[2]];
  return base ? base.map(f => m[1] + f) : null;
}

function inflect(phrase, form) {
  // "cheer up" inflects its verb; a noun phrase inflects its last word.
  const words = phrase.split(' ');
  const idx   = form === 'plural' ? words.length - 1 : 0;
  const w     = words[idx].toLowerCase();
  let out;
  switch (form) {
    case 'plural':      out = IRREGULAR_PLURAL[w] ?? _sForm(w); break;
    case '3sg':         out = w === 'be' ? 'is' : w === 'have' ? 'has' : _sForm(w); break;
    case 'past':        out = _irregular(w)?.[0] ?? _edForm(w); break;
    case 'pp':          out = _irregular(w)?.[1] ?? _edForm(w); break;
    case 'ing':         out = _ingForm(w); break;
    case 'comparative': out = _erForm(w, 'er'); break;
    case 'superlative': out = _erForm(w, 'est'); break;
    default:            return phrase;
  }
  words[idx] = out;
  return words.join(' ');
}

function _parseSynTemplate(line) {
  const inner = line.replace(/^#:\s*\{\{(?:syn|synonyms)\|en\|/, '').replace(/\}\}.*$/, '');
  const out = [];
  for (const raw of inner.split('|')) {
    if (!raw || raw.includes('=') || raw.startsWith('Thesaurus:')) continue;
    const mods = raw.match(/<[^>]*>/g)?.join(' ') ?? '';          // e.g. stour<q:archaic>
    if (DATED_SYNONYM.test(mods)) continue;
    const w = raw.replace(/<[^>]*>/g, '').replace(/\[\[|\]\]/g, '').trim();
    if (w) out.push(w);
  }
  return out;
}

const _DM_TAG = { noun: 'n', verb: 'v', adjective: 'adj', adverb: 'adv' };
const _dmFreq = d => parseFloat((d.tags ?? []).find(t => t.startsWith('f:'))?.slice(2) ?? '0');
// A candidate's main part of speech, not any it can take: "went" can be a
// noun, but offering it as a synonym for a noun is wrong.
const _dmPrimary = d => (d.tags ?? []).find(t => ['n', 'v', 'adj', 'adv'].includes(t));

function _datamuse(query, max = 40) {
  return _cachedJson(`https://api.datamuse.com/words?${query}&md=pf&max=${max}`,
                     { signal: AbortSignal.timeout(DICT_TIMEOUT_MS) }, res => {
    if (!res.ok) throw new Error(`Datamuse HTTP ${res.status}`);
    return res.json();
  });
}

// The start of a definition, as a reverse-dictionary query: "To move swiftly",
// "The act of deciding".  Long definitions return noise, so they're skipped.
function _definitionClause(text) {
  const c = text.replace(/^Short for “[^”]*”:\s*/, '').replace(/\([^)]*\)/g, '')
                .split(/[;.]/)[0].trim();
  return c && c.split(/\s+/).length <= 8 ? c : null;
}

// "schools", "schooling", "schoolhouse" -- but not "graduate" for "grad".
const _DERIVED = /^(s|es|d|ed|ing|ings|er|ers|est|ly|ness|ful|time|house|hood|ship|like|room|work)$/;
function _isDerivative(candidate, heads) {
  const c = candidate.toLowerCase();
  return heads.some(h => {
    h = h.toLowerCase();
    return c === h || (c.startsWith(h) && _DERIVED.test(c.slice(h.length)));
  });
}

// Builds the synonym row for one sense.  The sources and how far each is
// trusted were settled by comparing their output on real words:
//
//  - curated:  the sense's own synonym line on Wiktionary.  Right sense, but
//              sometimes thin, and sometimes obscure ("lickety-split").
//  - pool:     Datamuse "means like" for the word.  Its scores come in tiers:
//              ~1.0 for words it holds as synonyms, 0.75 merely related, 0.5
//              noise, so only the top tier is used.  Reliable for adjectives
//              and adverbs, and for ordinary nouns and verbs, but for a word
//              with dozens of senses it mixes them ("unraveled" for run).
//  - reverse:  the reverse dictionary on the definition itself ("To move
//              swiftly" -> dash, sprint, race).  Sense-specific but drifts, so
//              only words the pool also contains are taken.
//
// Adjectives and adverbs: curated words the pool also knows, in the pool's
// order; then the pool's top tier; then the remaining curated words.
// Nouns and verbs: curated first, in Wiktionary's order; if that leaves fewer
// than three, the reverse-dictionary intersection for many-sensed words, and
// the pool's top tier for the rest.  That tier describes whichever sense
// Datamuse treats as core, and its parts of speech show which: for "school"
// it is half verbs (cultivate, educate), for "big" nearly all adjectives.  So
// a secondary tab only uses it when its part of speech makes up a real share
// of the tier -- otherwise it gets outliers ("bragging" for big's noun).  For
// a many-sensed word it is never used ("pass" for walked).  An empty row is
// better than a wrong one.
// Candidates must be commonish single words whose main part of speech fits.
function _mergeSynonyms(heads, curated, pool, reverse, pos,
                        { allowPool = true, polysemous = false, primary = true } = {}) {
  const out  = [];
  const push = w => {
    if (out.length < MAX_SYNONYMS && !out.some(o => o.toLowerCase() === w.toLowerCase()) &&
        !_isDerivative(w, heads)) out.push(w);
  };
  const tag     = _DM_TAG[pos];
  const top     = pool[0]?.score ?? 0;
  const common  = d => _dmFreq(d) >= 0.5 && !/\s/.test(d.word);
  const topTier = allowPool ? pool.filter(d => d.score >= top * 0.95 && common(d) && _dmPrimary(d) === tag) : [];

  if (pos === 'adjective' || pos === 'adverb') {
    const known = new Map(pool.map((d, i) => [d.word.toLowerCase(), i]));
    curated.filter(w => known.has(w.toLowerCase()))
           .sort((x, y) => known.get(x.toLowerCase()) - known.get(y.toLowerCase()))
           .forEach(push);
    topTier.forEach(d => push(d.word));
    curated.forEach(push);
    return out;
  }

  curated.forEach(push);
  if (out.length >= 3 || !tag) return out;
  const cap = out.length ? 5 : MAX_SYNONYMS;
  if (polysemous && reverse) {
    const related = new Set(pool.map(d => d.word));
    for (const d of reverse) {
      if (out.length >= cap) break;
      if ((d.tags ?? []).includes(tag) && common(d) && related.has(d.word)) push(d.word);
    }
  } else if (!polysemous) {
    const tier  = pool.filter(d => d.score >= top * 0.95 && common(d));
    const share = tier.length ? tier.filter(d => _dmPrimary(d) === tag).length / tier.length : 0;
    if (primary || share >= 0.3)
      for (const d of topTier) { if (out.length >= cap) break; push(d.word); }
  }
  return out;
}

async function _lookupWiktionary(word) {
  // Titles are case-sensitive, and a capitalised selection is usually just the
  // start of a sentence: "School" should find school, not a proper noun.
  const lower  = word.toLowerCase();
  const titles = lower === word ? [word] : [lower, word];
  // The Datamuse list only depends on the title, so it rides the same round trip.
  _datamuse(`ml=${encodeURIComponent(lower)}`, 100).catch(() => {});
  let title, source;
  for (const t of titles) {
    try {
      source = await _wiktSource(t);
      if (!source.some(b => WORD_POS.has(b.pos) && b.senses.length)) throw noEntry();
      title = t; break;
    } catch (err) { if (!err.noEntry || t === titles[titles.length - 1]) throw err; }
  }

  // The first real sense of each part of speech, in page order.
  const picks = [];
  for (const blk of source) {
    if (!WORD_POS.has(blk.pos) || picks.some(p => p.pos === blk.pos)) continue;
    for (const sense of blk.senses) {
      if (sense.formOf?.caseOnly) continue;   // "letter-case form of Grad": a different word
      if (!/[a-z]/i.test(sense.plain)) sense.plain = '';   // "{{misspelling of|en|the}}." -> "."
      if (!sense.formOf && !sense.plain) continue;
      picks.push({ pos: blk.pos, def: { text: sense.plain, example: sense.example },
                   formOf: sense.formOf, syns: sense.syns, nSenses: blk.senses.length });
      break;
    }
  }
  // "ran" is mostly "past of run"; that reading goes ahead of the rare noun.
  picks.sort((x, y) => (y.formOf?.type === 'inflection') - (x.formOf?.type === 'inflection'));
  // For a word that is mainly an inflection, its own Datamuse list is made of
  // other inflections ("ran" -> "came", "went"), so its minor senses don't
  // draw on it.
  const inflectedWord = picks.some(p => p.formOf?.type === 'inflection');

  const meanings = await Promise.all(picks.slice(0, 5).map(async (pick, idx) => {
    let text = pick.def.text, example = pick.def.example, curated = pick.syns;
    let lemma = title, form = null, nSenses = pick.nSenses;
    const heads = [title];
    const f = pick.formOf;
    if (f && f.type !== 'misspelling') {
      // Follow the pointer and show the real entry's first sense.
      try {
        _datamuse(`ml=${encodeURIComponent(f.target)}`, 100).catch(() => {});   // same round trip
        const tSource = await _wiktSource(f.target);
        const tSrc    = tSource.find(x => x.pos === pick.pos) ?? tSource.find(x => WORD_POS.has(x.pos));
        const tSense  = tSrc?.senses.find(s => !s.formOf && /[a-z]/i.test(s.plain));
        const td      = tSense && { text: tSense.plain, example: tSense.example };
        if (td) {
          example = example || td.example;
          if (f.type === 'short') {
            // "grad" -> the full word is the best synonym there is.
            text    = `Short for “${f.target}”: ${td.text.charAt(0).toLowerCase()}${td.text.slice(1)}`;
            curated = [f.target, ...curated, ...(tSense?.syns ?? [])];
          } else {
            text  = td.text;
            heads.push(f.target);
            if (curated.length < 3) curated = tSense?.syns ?? curated;
            nSenses = tSrc?.senses.length ?? nSenses;
            // An inflected word's own Datamuse list is noise ("ran" -> "came,
            // went"): draw on the lemma's, then put every synonym into the
            // selected word's form.
            if (f.type === 'inflection') { lemma = f.target; form = f.form; }
          }
        }
      } catch { /* keep what the page itself said */ }
    }
    if (!text && f) text = `${f.type === 'misspelling' ? 'Misspelling' : 'Form'} of “${f.target}”.`;
    if (text && !/[.!?)]$/.test(text)) text += '.';
    if (!text) return null;

    // Many senses means the word-level lists blur them; "run" has dozens.
    const polysemous = nSenses >= 10;
    const clause = polysemous ? _definitionClause(text) : null;
    // The word's list is fetched deep (100) so a many-sensed word's real
    // synonyms ("sprint" for run), ranked below its other senses, are there.
    const [pool, reverse] = await Promise.all([
      _datamuse(`ml=${encodeURIComponent(lemma)}`, 100).catch(() => []),
      clause ? _datamuse(`ml=${encodeURIComponent(clause)}`).catch(() => null) : null,
    ]);
    let synonyms = _mergeSynonyms([...heads, lemma], curated, pool, reverse, pick.pos, {
      allowPool: !inflectedWord || lemma !== title,
      polysemous,
      primary: idx === 0,
    });
    if (form) synonyms = synonyms.map(w => inflect(w, form)).filter(w => !_isDerivative(w, heads));
    return { partOfSpeech: pick.pos, definition: text, example, synonyms };
  }));

  const found = meanings.filter(Boolean);
  if (!found.length) throw noEntry();
  return found;
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

  const entry = lookupCached(word);
  if (entry.settled) {
    // Prefetched while the pointer travelled to the pill: no spinner at all.
    popupEl = buildPopup(_cardData(word, entry), rect, theme, origin);
    container.appendChild(popupEl);
    return;
  }

  popupEl = buildPopup({ state: 'loading', word }, rect, theme, origin);
  container.appendChild(popupEl);
  // The card is on screen, so nothing is "pending" any more: the network wait
  // must not block pills elsewhere (showPill bails while this flag is set).
  _popupPending = false;
  const card = popupEl;
  await entry.promise.catch(() => {});
  // Closed, or replaced by another word's card, while loading: this result
  // belongs to neither, so drop it rather than write "A" into "B"'s card.
  if (popupEl !== card) return;

  // Update in-place so the popup never disappears — no DOM remove/re-add, no
  // animation replay, no blank frame between loading state and loaded state.
  _updatePopupBody(card, _cardData(word, entry));
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
