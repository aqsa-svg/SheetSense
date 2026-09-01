# SheetSense

**Type `=SENTIMENT(A2)` into a spreadsheet cell and drag it down a thousand rows — AI as a formula, with no add-on to install and nothing to pay for.**

<!-- SCREENSHOT: a sheet with =SENTIMENT and =CLASSIFY filled down a column.
     Replace with: ![SheetSense in a sheet](docs/screenshot-sheet.png) -->
> 📸 *Screenshot placeholder — a column of `=SENTIMENT()` and `=CLASSIFY()` results filled down beside raw text.*

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

## The problem

Getting an LLM to touch spreadsheet data normally means exporting to CSV, writing a script, and pasting results back — or installing a paid marketplace add-on that wants access to all your files. Neither works for the actual use case, which is "I have 800 rows of feedback and I want a sentiment column."

The awkward part is that Google Sheets recalculates custom functions aggressively and in bulk. Fill a formula down 1,000 rows and you have issued 1,000 API calls, and free-tier Gemini will rate-limit you long before the column finishes. **So the interesting engineering here isn't the prompt — it's not burning your quota.**

## Results

Measured by reading the source; this is Apps Script, so it runs inside Google Sheets rather than locally.

| Measured | Value |
|---|---|
| Custom functions exposed to cells | **5** — `SUMMARIZE`, `SENTIMENT`, `CLASSIFY`, `EXTRACT`, `ASK` |
| Menu actions | **6** — set API key, test connection, set model, clear cache, about, onOpen |
| Source size | **661 lines** Apps Script (V8 runtime) |
| Default model | `gemini-2.5-flash` (switchable to `-flash-lite` / `-pro` from the menu) |
| Temperature | `0.2` — chosen for repeatable cell values |
| Retry policy | up to **3 attempts**, backoff 500 → 1000 → 2000 ms |
| Retried status codes | **429, 500, 503 only** — everything else fails fast |
| Response cache | MD5-keyed, **6 h TTL** (`CacheService` hard maximum) |
| API calls for empty input | **0** — short-circuits before `UrlFetchApp` |
| OAuth scopes requested | **2** — `script.external_request`, `script.container.ui` |

**[TODO] Cache hit rate on a real fill-down.** The caching is the whole quota argument and its effect is unmeasured. *To measure:* fill `=SENTIMENT()` down 500 rows with ~30% duplicate text, then compare the Apps Script execution log's `UrlFetchApp` call count against 500. Report calls saved.

**[TODO] Median latency per cell**, and how many rows can be filled before free-tier rate limiting kicks in.

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

## Design decisions and tradeoffs

**Cache first, because Sheets recalculates without asking.** Successful responses are cached under an MD5 of `(function, model, prompt)` for six hours — `CacheService`'s hard ceiling. A sheet that recalculates on open, or a column with repeated text, costs one API call per *distinct* input rather than one per cell. The tradeoff is that a cached cell won't reflect a model or prompt change for up to six hours, which is why "Clear cache" exists as a menu item.

**Clearing the cache bumps a version counter instead of deleting keys.** `CacheService` has no bulk-clear, so `CACHE_VERSION` is part of every cache key and incrementing it orphans every existing entry at once. Cheap and instant; the cost is that the stale entries sit there until their TTL expires rather than being freed.

**Empty input never reaches the API.** Filling a formula down a column always overshoots into blank rows. Short-circuiting empty input to `""` before `UrlFetchApp` means a 1,000-row fill over 800 rows of data costs 800 calls, not 1,000. Small decision, direct quota saving.

**Retry only on 429/500/503, fail fast on everything else.** Transient limits and server errors are worth a bounded retry with exponential backoff; a 400 or a 403 is a bad request or a bad key and retrying it just burns the 30-second cell budget three times over. `MAX_ATTEMPTS = 3` is set by that budget, not by optimism.

**Errors return a readable string, never an exception.** A thrown error in a custom function poisons the cell with `#ERROR!` and tells the user nothing. Returning `⚠ <reason>` keeps the sheet intact and diagnosable. The tradeoff: `ISERROR()` won't catch it, because as far as Sheets is concerned the cell succeeded.

**Temperature 0.2, not 0.** Cell values should be stable across recalculation, so temperature is low — but not zero, which in practice gives no meaningful additional determinism and can make short outputs degenerate.

**AI Studio API key, not Vertex AI.** Vertex would mean service accounts and a GCP project. An AI Studio key pasted into Script Properties keeps setup to two menu clicks on the free tier. The cost is no IAM, no per-user auth, and one shared key per sheet.

**The key lives in Script Properties, never in a cell.** A key in a cell is visible to every viewer, travels with a copy of the sheet, and lands in version history.

## Known limitations

- **Quota is shared and unmetered.** One API key serves the whole sheet, and there's no per-user or per-day accounting. A large fill-down by one editor can rate-limit everyone.
- **Cache hit rate and latency are unmeasured.** See the `[TODO]`s above — the central efficiency claim has no numbers behind it.
- **Cached values can be up to 6 hours stale**, including after a model change, until the cache is cleared from the menu.
- **Errors are invisible to spreadsheet error handling.** `⚠ …` is a normal string, so `IFERROR`/`ISERROR` won't trap it.
- **A 30-second per-cell ceiling** is imposed by Apps Script, which is what caps retries at 3. Very long inputs can time out.
- **Anyone who can edit the sheet can read the API key** via the Apps Script editor. Script Properties keeps it out of cells, not out of the project.
- **`CLASSIFY` output is coerced to your label list**, so a genuinely ambiguous input still returns one of the labels rather than admitting uncertainty.
- **No tests.** Apps Script has no local test harness here, and nothing is verified automatically.
- **Not deployable as a shareable add-on.** This is source to paste into one spreadsheet's script project, not a published Workspace Marketplace add-on.
