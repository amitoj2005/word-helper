<div align="center">

<img src="icon128.png" width="96" alt="">

# Word Helper

**Look up any word in Google Docs without leaving the document.**

Select a word, click the pill that appears, and get its definition, pronunciation,
and part of speech — plus one-click synonyms that replace the word in place.

</div>

---

## Demo

<!-- TODO: record a ~10s screen capture of select -> pill -> card -> synonym swap,
     save it as docs/demo.gif, and uncomment the line below. -->
<!-- ![Word Helper in action](docs/demo.gif) -->

*Screen recording coming soon.*

---

## Features

- **Non-intrusive trigger.** Select a word and a small pill appears beside it; click it (or hover for 300 ms) to open the card. Double-click deliberately does *nothing* — that's Google Docs' own select-a-word gesture, and hijacking it interrupted ordinary editing.
- **One-click synonym replacement.** Click a synonym and it replaces the selected word directly in the document, matching the original capitalisation and preserving a trailing space.
- **Never touches your clipboard.** Reading the selection out of Google Docs requires provoking a copy, so the extension intercepts it before anything reaches the OS clipboard. Details below.
- **Five popup themes**, chosen from the toolbar button.
- **Keyboard and context-menu access** — `Alt+Shift+D`, or right-click → *Look up "…"*.

### Themes

| Theme | Look |
|---|---|
| **Modern Glass** *(default)* | Blue-white gradient, rendered in a shadow root |
| **Old Dictionary** | Parchment and serif type |
| **Liquid Glass** | Live WebGL refraction of the page behind the card |
| **Liquid Glass HD** | Same, rendered at device pixel ratio from a quality-100 capture |
| **Frosted Glass** | `backdrop-filter` blur on the card itself |

---

## Install

Not yet on the Chrome Web Store. To run it locally:

1. Clone or download this repository.
2. Open `chrome://extensions` and switch on **Developer mode**.
3. Click **Load unpacked** and select the project folder.
4. Open any Google Doc and select a word.

---

## How it works

The interesting problems here aren't the dictionary lookup — they're getting text
*out of* Google Docs and putting a synonym back *into* it.

### Reading the selection without touching the clipboard

Google Docs renders its document to canvas, so there is no DOM text node to read
and `window.getSelection()` is usually empty. The only reliable way to get the
selected text is to trigger a copy — but silently overwriting the user's
clipboard every time they select a word is unacceptable.

The fix is a two-layer interception:

1. **`content-main.js`** runs in the page's MAIN world at `document_start`, early
   enough to patch `DataTransfer.prototype.setData` before Docs' own scripts
   load. When a flag attribute is set on `<html>`, it stashes the text in a DOM
   attribute instead of forwarding it to the native `DataTransfer`, so the C++
   backing store stays empty.
2. **`content.js`** (isolated world) sets that flag, fires
   `execCommand('copy')`, and calls `preventDefault()` on the copy event in the
   capture phase to block the browser's native-selection fallback.

Docs still calls `setData()` and the text is still readable — but the backing
store is empty and the default is prevented, so nothing is ever committed to the
system clipboard.

### Coordinating across frames

The Docs editor lives in an iframe, while the popup must render in the top frame.
The two coordinate over a `BroadcastChannel`: the iframe that owns the selection
announces the word it found, and the top frame renders the card. A `direct` flag
rides along on those messages so an explicit trigger (`Alt+Shift+D`, context
menu) skips the pill and opens the card immediately, while a plain selection
stops at the pill.

Selection probing is debounced and gated — a mouseup only counts if the pointer
dragged at least 3 px or it was a multi-click — so a plain caret click never
fires the `execCommand`-based read.

### Writing the synonym back

Replacement dispatches a synthetic `paste` `ClipboardEvent` at the editor. A
trailing space needs separate handling: pasting `" "` gets trimmed, and
`execCommand` is intercepted, so the extension dispatches a `keydown`/`keypress`/
`keyup` triple for Space and lets Docs' own key handler insert it into the model.

### The liquid glass themes

`liquid` and `liquidhd` capture the visible tab via `chrome.tabs.captureVisibleTab`
and use it as a WebGL texture, refracting it through a signed-distance-field
rounded rectangle with a Fresnel rim and chromatic aberration. The SDF's
half-extents have to match the CSS box exactly, or the mismatch shows up as
bright artifacts along the edges.

These two themes and `frosted` mount into `document.body` rather than a shadow
root — `.wh-popup` sets `isolation: isolate`, which makes it a backdrop root, so
a child's `backdrop-filter` would see nothing but the popup's own interior.

---

## Permissions

| Permission | Why |
|---|---|
| `https://docs.google.com/document/*` | Inject the content scripts |
| `https://api.dictionaryapi.dev/*` | Definitions and pronunciations (no API key, no account) |
| `<all_urls>` | Required by `captureVisibleTab`, which the liquid-glass themes use to sample the page behind the card. Chrome demands it even when the tab already matches a narrower host permission. |
| `storage` | Remember the selected theme |
| `contextMenus` | The right-click *Look up "…"* entry |
| `tabs` | Resolve the window ID for `captureVisibleTab` |

No analytics, no tracking, no account. The only network request is to
`api.dictionaryapi.dev` for the word you looked up.

---

## Development

No build step — it's plain JavaScript loaded directly by Chrome.

| Path | Role |
|---|---|
| `content.js` | Triggers, pill, card, cross-frame coordination, WebGL themes |
| `content-main.js` | MAIN-world `setData` patch |
| `content.css` | All popup and pill styles |
| `background.js` | Service worker: tab capture and the context menu |
| `popup.html` / `popup.js` | Theme picker |
| `assets/make_icon.py` | Regenerates `icon16/48/128.png` |

### Regenerating the icons

```sh
python assets/make_icon.py
```

Crops `assets/icon-source.png` to its tile, masks the surround to transparency at
the tile's corner radius, and writes the three sizes. The mask is downsampled
with `Image.BOX` rather than LANCZOS — ringing on an alpha channel pushes edge
values past 0/255. Requires Pillow.

---

## License

[MIT](LICENSE)
