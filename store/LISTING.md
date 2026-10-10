# Chrome Web Store submission — Word Helper

Everything the developer dashboard asks for, ready to paste. Upload
`dist/word-helper-<version>.zip` (built with `python scripts/build.py`).

---

## Store listing tab

**Name**
```
Word Helper
```

**Summary** (from the manifest; 132 characters max)
```
Select a word in Google Docs to see its definition and swap it for a synonym.
```

**Category:** Productivity  **Language:** English

**Description**
```
Look up any word without leaving your Google Doc.

Select a word and a small magnifier appears beside it. Click it and a card opens right next to the word with its definition, an example sentence and synonyms. Click a synonym and it replaces the word in your document.

• Instant: the lookup starts the moment the magnifier appears, so the card usually opens with everything already there. Words you've looked up before appear immediately.
• Synonyms that fit: synonyms come from the sense of the word you're reading, and match its form, so "ran" offers "sprinted" and "schools" offers "academies".
• Keyboard and right-click: press Alt+Shift+D, or right-click a selected word and choose Look up.
• Four card styles: Modern Glass, Old Dictionary (it unrolls like a scroll), Liquid Glass (it bends the page behind it) and Frosted Glass. Switch any time from the toolbar button.
• Private by design: only the single word you look up is sent to public dictionary services. No accounts, no analytics, no tracking, no ads. Your clipboard is never read or changed.

Definitions and synonyms come from Wiktionary (CC BY-SA 4.0) and Datamuse, with the Free Dictionary API as a fallback.

Open source: https://github.com/amitoj2005/word-helper
```

**Graphics** (all in `store/`)

| Field | File | Size |
|---|---|---|
| Store icon | `../icon128.png` | 128×128 |
| Screenshot 1 | `screenshot-1-lookup.png` | 1280×800 |
| Screenshot 2 | `screenshot-2-synonyms.png` | 1280×800 |
| Screenshot 3 | `screenshot-3-styles.png` | 1280×800 |
| Screenshot 4 | `screenshot-4-dictionary.png` | 1280×800 |
| Small promo tile | `promo-small-440x280.png` | 440×280 |

The screenshots are real captures of the extension, shown on a plain sample
document rather than a Google Docs window.

**Homepage URL:** `https://github.com/amitoj2005/word-helper`
**Support URL:** `https://github.com/amitoj2005/word-helper/issues`

---

## Privacy practices tab

**Single purpose**
```
Word Helper shows the definition, an example and synonyms for a word the user selects in a Google Docs document, and lets them replace that word with a synonym.
```

**Permission justifications**

`storage`
```
Saves the user's chosen card style, and caches recent lookups (word, definition, synonyms) on the device so a repeated word appears instantly. Nothing stored is transmitted.
```

`contextMenus`
```
Adds a "Look up" item to the right-click menu for a selected word in Google Docs, as an alternative to the on-page button.
```

Host permission / content scripts (`https://docs.google.com/document/*`)
```
The extension's content scripts run only on Google Docs documents, where it detects the selected word, shows the definition card beside it, and inserts a chosen synonym. It requests no host permissions and runs on no other site.

Google Docs exposes the selected text only through its copy handler, so the extension triggers a copy and intercepts it before anything reaches the system clipboard. content-main.js runs in the page's main world solely to do that interception: it overrides DataTransfer.setData to capture the text instead of writing it to the clipboard. The user's clipboard is never read or changed.
```

**Remote code:** No, I am not using remote code.
(No external scripts, no eval; the extension only fetches dictionary data as JSON.)

**Data usage** — tick only:
- ☑ **Website content** — the single word the user selects, sent to public
  dictionary services (en.wiktionary.org, api.datamuse.com, and
  api.dictionaryapi.dev as a fallback) to retrieve its definition and synonyms.

Leave every other category unticked (no personally identifiable information,
health, financial, authentication, personal communications, location, web
history or user activity).

Tick all three certifications:
- ☑ I do not sell or transfer user data to third parties, outside of the approved use cases
- ☑ I do not use or transfer user data for purposes that are unrelated to my item's single purpose
- ☑ I do not use or transfer user data to determine creditworthiness or for lending purposes

**Privacy policy URL**
```
https://github.com/amitoj2005/word-helper/blob/main/PRIVACY.md
```

---

## Distribution tab

Visibility: **Public** (or **Unlisted** for a soft launch). Regions: all.
Pricing: free.

---

## Before you submit

1. `python scripts/build.py` → `dist/word-helper-1.0.zip`.
2. Load the zip's contents in a clean Chrome profile (`chrome://extensions` →
   Developer mode → Load unpacked) and try it on a document you're editing.
3. Register a developer account ($5, one-time) with 2-step verification on.
4. New item → upload the zip → fill in the three tabs above → Submit for review.

For updates, bump `version` in `manifest.json`, rebuild, and upload the new zip
to the same item.
