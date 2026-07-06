# SheetSense

AI functions for Google Sheets, powered by the **Gemini API** (Google AI Studio,
free tier). Use them like native formulas and fill them down a column:

| Function | What it does |
|---|---|
| `=SUMMARIZE(text, [maxWords])` | Concise summary (default **40** words). |
| `=SENTIMENT(text)` | Returns exactly `Positive`, `Negative`, or `Neutral`. |
| `=CLASSIFY(text, "a,b,c")` | Returns exactly **one** of the comma-separated labels. |
| `=EXTRACT(text, field)` | Pulls a single value; returns `N/A` if it isn't present. |
| `=ASK(prompt, [context])` | Free-form answer; with `context`, answers **only** from it. |

Any argument can be a cell, a literal string, or a **range** (it's flattened to
text). Empty input returns `""` without calling the API. Errors never break the
cell — they come back as a readable string starting with `⚠ `.

> This is Google Apps Script code — paste it into the Apps Script editor. It does
> not run locally.

---

## Setup

1. **Get a Gemini API key** (free): <https://aistudio.google.com/apikey>.

2. **Open the script editor.** In your Google Sheet: **Extensions ▸ Apps Script**.

3. **Add the code.**
   - Replace the contents of the default `Code.gs` with this repo's [`Code.gs`](Code.gs).
   - Show the manifest: **Project Settings ▸** check **"Show `appsscript.json`
     manifest file in editor"**, then open `appsscript.json` and replace its
     contents with this repo's [`appsscript.json`](appsscript.json).
   - **Save** (Ctrl/Cmd + S).

4. **Reload the Google Sheet.** A **SheetSense** menu appears next to *Help*.

5. **Set your API key.** **SheetSense ▸ Set / update Gemini API key** → paste the
   key. It's stored in **Script Properties** (a server-only secret) — never in a
   cell, never in the code.

6. **Authorize once — this is the important step.** **SheetSense ▸ Test connection**.
   - Google shows an authorization prompt the first time. Approve it (choose your
     account → *Advanced* → *Go to …(unsafe)* if shown, since the script is
     unverified and yours → *Allow*).
   - You should then see **"Connection OK"**.

   **Why this step exists:** custom cell functions *are allowed* to make external
   requests, but they *cannot* pop the authorization dialog. Running any menu item
   that calls the API grants the `script.external_request` scope once — after that,
   the `=SUMMARIZE(...)` etc. cell functions work.

7. **Use the functions** in any cell (see below).

---

## Usage

```text
=SUMMARIZE(A2)                     Summarize the text in A2 (≤ 40 words).
=SUMMARIZE(A2, 15)                 Summarize in ≤ 15 words.
=SUMMARIZE(A2:A10)                 Summarize a whole range as one block of text.

=SENTIMENT(B2)                     -> Positive | Negative | Neutral

=CLASSIFY(C2, "Bug,Feature,Question")
=CLASSIFY(C2, D1:F1)               Labels can come from a range too.

=EXTRACT(E2, "email")              -> jane@acme.com   (or "N/A")
=EXTRACT(E2, "invoice total")

=ASK("Reply in one word: is this urgent? " & F2)
=ASK("What is the refund policy?", G2:G50)   Answer only from the context range.
```

Fill any of these down a column just like a normal formula. Identical inputs are
served from cache, so recalculating a sheet does not re-bill the API.

---

## Menu reference (SheetSense)

- **Set / update Gemini API key** — stores the key in Script Properties.
- **Test connection** — makes one real API call; grants the external-request
  authorization (do this once after setting the key).
- **Set model** — override the model id (stored in Script Properties). Default
  `gemini-2.5-flash`. Other options: `gemini-2.5-flash-lite`, `gemini-2.5-pro`.
- **Clear cache** — bumps a `CACHE_VERSION` property to invalidate cached answers
  (CacheService has no bulk clear). Cells re-query on their next recalculation.
- **About** — function list and credit.

---

## How it works

- **Secret handling** — the API key lives only in **Script Properties** (readable
  server-side inside custom functions), never in a cell or in the source.
- **Caching** — successful responses are cached with **CacheService** (document
  cache, falling back to script cache) for **6 hours**, keyed by an **MD5 hash** of
  `CACHE_VERSION + model + function-tag + prompt`.
- **Retry** — bounded exponential backoff (500 ms → 1 s → 2 s, max **3** attempts)
  on transient **HTTP 429 / 500 / 503** only, using `muteHttpExceptions: true`.
  Kept short so functions return within the ~30 s custom-function limit.
- **Errors** — never thrown to the cell; returned as `⚠ …` strings (no key set,
  content blocked, HTTP code + short message, etc.).
- **Model quirk** — `gemini-2.5-flash*` models "think" by default, which can
  consume the output budget and slow responses; SheetSense sends
  `thinkingConfig.thinkingBudget: 0` for the flash tiers to stay fast.

---

## Gemini API

- Uses the **AI Studio Generative Language API** with **API-key auth** (free tier)
  — *not* Vertex AI, so there are no service accounts.
- Endpoint:
  `POST https://generativelanguage.googleapis.com/v1beta/models/{MODEL}:generateContent?key={KEY}`
- `generationConfig.temperature = 0.2`.
- Model ids can change over time. If calls fail with a **404 "model not found"**,
  pick a current id from <https://ai.google.dev/gemini-api/docs/models> and set it
  via **SheetSense ▸ Set model**.

---

## OAuth scopes

The manifest requests only the two scopes these features need:

- `https://www.googleapis.com/auth/script.external_request` — call the Gemini API.
- `https://www.googleapis.com/auth/script.container.ui` — build the menu and show
  prompts/dialogs.

---

Built by **Aqsa Siddiqui**.
