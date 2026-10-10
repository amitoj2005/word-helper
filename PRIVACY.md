# Privacy Policy — Word Helper

**Effective date:** October 10, 2026

Word Helper is a Chrome extension that shows the definition and synonyms of a
word you select in Google Docs. This policy explains exactly what data it
handles. The short version: **it sends only the single word you look up to
public dictionary services, keeps nothing about you, and has no analytics,
tracking, accounts or ads.**

## What leaves your device

When you look up a word, that one word — and nothing else from your document —
is sent to these public dictionary services to fetch its definition and
synonyms:

| Service | Operator | Used for |
|---|---|---|
| `en.wiktionary.org` | Wikimedia Foundation | Definitions, examples and synonyms (primary source) |
| `api.datamuse.com` | Datamuse | Additional synonyms |
| `api.dictionaryapi.dev` | Free Dictionary API | Fallback only, if Wiktionary can't be reached |

A lookup happens when the lookup button appears beside a word you've selected
(so the result is ready when you click it), when you press Alt+Shift+D, or when
you choose **Look up** from the right-click menu.

As with any web request, these services also receive the technical information
every browser sends — your IP address and browser version — and handle it under
their own policies:
[Wikimedia Foundation Privacy Policy](https://foundation.wikimedia.org/wiki/Policy:Privacy_policy),
[Datamuse](https://www.datamuse.com/api/),
[Free Dictionary API](https://dictionaryapi.dev/).

Word Helper does not send the rest of your document, your document's title or
address, your Google account details, or anything else.

## What stays on your device

Stored in Chrome's local extension storage, on your computer only:

- **Your chosen card style.**
- **Recent lookups** — the words you've looked up and their definitions and
  synonyms, so a repeated word appears instantly. At most 400 words, each kept
  for 14 days. Removing the extension deletes them.

This data is never transmitted anywhere.

## Your clipboard

To read the selected word from Google Docs, the extension briefly triggers a
copy action and intercepts it before anything reaches your clipboard. Your
clipboard is never read and never changed.

## What Word Helper does not do

- No analytics, telemetry, crash reporting or usage tracking.
- No accounts, sign-in or personal information.
- No advertising, and no selling or sharing of data.
- No access to any site other than Google Docs documents.

## Changes

If this policy changes, the updated version will be published here with a new
effective date.

## Contact

Questions or concerns: open an issue at
<https://github.com/amitoj2005/word-helper/issues>.
