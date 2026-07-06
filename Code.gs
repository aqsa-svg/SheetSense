/**
 * ============================================================================
 *  SheetSense — AI functions for Google Sheets, powered by the Gemini API
 * ============================================================================
 *
 *  Custom cell functions (use them like native formulas; fill down a column):
 *    =SUMMARIZE(text, [maxWords])  Concise summary (default 40 words).
 *    =SENTIMENT(text)              Exactly one of: Positive, Negative, Neutral.
 *    =CLASSIFY(text, "a,b,c")      Returns exactly one of the given labels.
 *    =EXTRACT(text, field)         Pulls one value; "N/A" if not present.
 *    =ASK(prompt, [context])       Free-form answer; if context is supplied,
 *                                  the answer is drawn only from that context.
 *
 *  For every function: a range / 2D array argument is flattened to plain text,
 *  and an empty input returns "" WITHOUT calling the API (no wasted quota).
 *
 *  ---------------------------------------------------------------------------
 *  Apps Script constraints this file is designed around
 *  ---------------------------------------------------------------------------
 *  - Custom functions MAY call UrlFetchApp (external requests are not treated
 *    as personal-data access), so the functions above are real =FORMULA() calls
 *    — not a menu-only tool.
 *  - Custom functions CANNOT trigger the OAuth authorization prompt. So the
 *    "Test connection" menu item makes a normal (menu-triggered) Gemini call;
 *    running it once grants the external_request scope, after which the cell
 *    functions work. See the README.
 *  - Custom functions must return quickly (~30s), so retries are bounded (3).
 *
 *  ---------------------------------------------------------------------------
 *  Design
 *  ---------------------------------------------------------------------------
 *  SECRET:  The Gemini API key lives in Script Properties (a server-only store,
 *           readable inside custom functions but never exposed in a cell). It is
 *           set through a menu prompt — never hardcoded and never typed into a
 *           cell. The model id and cache version also live in Script Properties.
 *
 *  CACHING: Successful responses are cached with CacheService (document cache,
 *           falling back to script cache). The key is an MD5 hash of
 *           CACHE_VERSION + model + function-tag + prompt, with a 6-hour TTL
 *           (CacheService maximum). Identical inputs therefore reuse the cached
 *           answer, so recalculating the sheet does not re-bill the API. Because
 *           CacheService has no bulk-clear, "Clear cache" bumps the CACHE_VERSION
 *           property, which changes every future key and orphans the old ones.
 *
 *  RETRY:   Requests use muteHttpExceptions:true and retry with exponential
 *           backoff (500ms, 1000ms, 2000ms) up to 3 attempts, but ONLY on
 *           transient HTTP 429 / 500 / 503. Other codes fail fast.
 *
 *  ERRORS:  Nothing is ever thrown to the cell. Every failure path returns a
 *           readable string prefixed with "⚠ " (no key set, blocked content,
 *           HTTP code + short message, parse error, etc.).
 *
 *  Gemini:  AI Studio Generative Language API with API-key auth (free tier) —
 *           NOT Vertex AI, so there are no service accounts. Endpoint:
 *           POST .../v1beta/models/{MODEL}:generateContent?key={KEY}
 *
 *  Built by Aqsa Siddiqui.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Configuration constants
// ---------------------------------------------------------------------------

/** Base URL for the AI Studio Generative Language API (append MODEL + method). */
var ENDPOINT_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';

/**
 * Default model. gemini-2.5-flash is fast and free-tier friendly.
 * NOTE: model ids can be deprecated/renamed over time — if calls start failing
 * with a 404 "model not found", set a current id via the "Set model" menu item.
 * Check the live list at: https://ai.google.dev/gemini-api/docs/models
 */
var DEFAULT_MODEL = 'gemini-2.5-flash';

var TEMPERATURE           = 0.2;    // low temperature => stable, repeatable output
var MAX_ATTEMPTS          = 3;      // bounded so a cell returns within ~30s
var INITIAL_BACKOFF_MS    = 500;    // doubles each retry: 500 -> 1000 -> 2000
var CACHE_TTL_SECONDS     = 21600;  // 6 hours (CacheService hard maximum)
var DEFAULT_SUMMARY_WORDS = 40;

// Script Property keys (server-side storage).
var PROP_API_KEY       = 'GEMINI_API_KEY';
var PROP_MODEL         = 'GEMINI_MODEL';
var PROP_CACHE_VERSION = 'CACHE_VERSION';

// ===========================================================================
// Menu / authorization (these run in a normal, authorizable context)
// ===========================================================================

/**
 * Builds the SheetSense menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SheetSense')
    .addItem('Set / update Gemini API key', 'ss_setApiKey')
    .addItem('Test connection', 'ss_testConnection')
    .addItem('Set model', 'ss_setModel')
    .addSeparator()
    .addItem('Clear cache', 'ss_clearCache')
    .addItem('About', 'ss_about')
    .addToUi();
}

/**
 * Prompts for the Gemini API key and stores it in Script Properties.
 * The key is a server-only secret — it never appears in a cell.
 */
function ss_setApiKey() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'SheetSense',
    'Paste your Gemini API key.\n' +
    'Get one free at https://aistudio.google.com/apikey\n\n' +
    'It is stored in Script Properties (server-only) — never in a cell.',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var key = (resp.getResponseText() || '').trim();
  if (!key) {
    ui.alert('SheetSense', 'No key entered — nothing was saved.', ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty(PROP_API_KEY, key);
  ui.alert('SheetSense',
    'API key saved.\n\nNext: run SheetSense ▸ "Test connection" once to authorize ' +
    'external requests. After that, the =SUMMARIZE / =SENTIMENT / … cell ' +
    'functions will work.',
    ui.ButtonSet.OK);
}

/**
 * Makes a real Gemini call from a menu (authorizable) context. Running this
 * once grants the script.external_request scope — the one-time authorization
 * that custom functions cannot trigger on their own.
 */
function ss_testConnection() {
  var ui = SpreadsheetApp.getUi();
  var key = getApiKey_();
  if (!key) {
    ui.alert('SheetSense',
      'Set your API key first: SheetSense ▸ "Set / update Gemini API key".',
      ui.ButtonSet.OK);
    return;
  }
  var model = getModel_();
  // Bypass the cache on purpose — we want a genuine round-trip to the API.
  var result = fetchGemini_(model, key, 'Reply with the single word: OK');
  if (result.ok) {
    ui.alert('SheetSense',
      'Connection OK using model "' + model + '".\n' +
      'Model replied: ' + result.text + '\n\n' +
      'Your cell functions (=SUMMARIZE, =SENTIMENT, =CLASSIFY, =EXTRACT, =ASK) ' +
      'are ready to use.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('SheetSense', 'Connection failed:\n\n' + result.message, ui.ButtonSet.OK);
  }
}

/**
 * Lets the user override the model id (stored in Script Properties).
 */
function ss_setModel() {
  var ui = SpreadsheetApp.getUi();
  var current = getModel_();
  var resp = ui.prompt(
    'SheetSense',
    'Enter the Gemini model id (current: ' + current + ').\n' +
    'Examples: gemini-2.5-flash, gemini-2.5-pro, gemini-2.5-flash-lite.\n' +
    'Model list: https://ai.google.dev/gemini-api/docs/models',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var model = (resp.getResponseText() || '').trim();
  if (!model) {
    ui.alert('SheetSense', 'No model entered — keeping "' + current + '".', ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty(PROP_MODEL, model);
  ui.alert('SheetSense',
    'Model set to "' + model + '".\nTip: run "Test connection" to confirm it works.',
    ui.ButtonSet.OK);
}

/**
 * CacheService has no bulk-clear, so we invalidate every existing entry by
 * bumping CACHE_VERSION — future keys will differ and old ones simply expire.
 */
function ss_clearCache() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var v = parseInt(props.getProperty(PROP_CACHE_VERSION) || '1', 10);
  if (isNaN(v)) v = 1;
  v += 1;
  props.setProperty(PROP_CACHE_VERSION, String(v));
  ui.alert('SheetSense',
    'Cache invalidated (version bumped to ' + v + ').\n\n' +
    'Cells will re-query the API on their next recalculation. To force one, ' +
    'edit and re-enter a formula, or delete and retype it.',
    ui.ButtonSet.OK);
}

/**
 * Shows the function reference and credit.
 */
function ss_about() {
  var ui = SpreadsheetApp.getUi();
  var msg =
    'SheetSense — AI functions for Google Sheets (Gemini API)\n\n' +
    'Custom functions:\n' +
    '  =SUMMARIZE(text, [maxWords])   concise summary (default 40 words)\n' +
    '  =SENTIMENT(text)               Positive / Negative / Neutral\n' +
    '  =CLASSIFY(text, "a,b,c")       returns exactly one label\n' +
    '  =EXTRACT(text, field)          pulls one value; "N/A" if absent\n' +
    '  =ASK(prompt, [context])        free-form answer\n\n' +
    'Current model: ' + getModel_() + '\n\n' +
    'Built by Aqsa Siddiqui.';
  ui.alert('SheetSense', msg, ui.ButtonSet.OK);
}

// ===========================================================================
// Custom functions (real spreadsheet formulas)
// ===========================================================================

/**
 * Summarizes text using the Gemini API.
 *
 * @param {string|Array} text The text (or a range) to summarize.
 * @param {number} [maxWords] Maximum words in the summary. Default 40.
 * @return {string} A concise summary, or "" for empty input.
 * @customfunction
 */
function SUMMARIZE(text, maxWords) {
  var t = flattenInput_(text);
  if (!t) return '';

  var words = Number(maxWords);
  if (!words || words <= 0) words = DEFAULT_SUMMARY_WORDS;
  words = Math.floor(words);

  var prompt =
    'Summarize the following text in at most ' + words + ' words. ' +
    'Reply with only the summary — no preamble, labels, or quotation marks.\n\n' +
    'Text:\n' + t;

  return generate_('SUMMARIZE', prompt);
}

/**
 * Classifies the sentiment of text as Positive, Negative, or Neutral.
 *
 * @param {string|Array} text The text (or a range) to analyze.
 * @return {string} "Positive", "Negative", or "Neutral"; "" for empty input.
 * @customfunction
 */
function SENTIMENT(text) {
  var t = flattenInput_(text);
  if (!t) return '';

  var labels = ['Positive', 'Negative', 'Neutral'];
  var prompt =
    'Classify the sentiment of the following text as exactly one of: ' +
    labels.join(', ') + '. Reply with only that single word.\n\n' +
    'Text:\n' + t;

  var out = generate_('SENTIMENT', prompt);
  if (isError_(out) || out === '') return out;
  return normalizeToLabels_(out, labels);
}

/**
 * Classifies text into exactly one of the provided labels.
 *
 * @param {string|Array} text The text (or a range) to classify.
 * @param {string} labels Comma-separated labels, e.g. "Bug,Feature,Question".
 * @return {string} One of the given labels; "" for empty input.
 * @customfunction
 */
function CLASSIFY(text, labels) {
  var t = flattenInput_(text);
  if (!t) return '';

  var labelStr = flattenInput_(labels);
  var labelArr = labelStr.split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0; });
  if (!labelArr.length) {
    return '⚠ CLASSIFY needs comma-separated labels, e.g. CLASSIFY(A2,"Bug,Feature,Question").';
  }

  var prompt =
    'Classify the following text into exactly one of these labels: ' +
    labelArr.join(', ') + '. Reply with only the single most appropriate label, ' +
    'exactly as written above.\n\n' +
    'Text:\n' + t;

  var out = generate_('CLASSIFY', prompt);
  if (isError_(out) || out === '') return out;
  return normalizeToLabels_(out, labelArr);
}

/**
 * Extracts a single field value from text.
 *
 * @param {string|Array} text The text (or a range) to read.
 * @param {string} field The field to pull, e.g. "email" or "invoice total".
 * @return {string} The extracted value, or "N/A" if it is not present.
 * @customfunction
 */
function EXTRACT(text, field) {
  var t = flattenInput_(text);
  if (!t) return '';

  var f = flattenInput_(field);
  if (!f) return '⚠ EXTRACT needs a field name, e.g. EXTRACT(A2,"email").';

  var prompt =
    'From the text below, extract the value of "' + f + '". ' +
    'Reply with only the value and nothing else. ' +
    'If it is not present, reply with exactly: N/A.\n\n' +
    'Text:\n' + t;

  var out = generate_('EXTRACT', prompt);
  if (isError_(out)) return out;
  return out === '' ? 'N/A' : out;
}

/**
 * Answers a free-form prompt. If context is supplied, the answer is drawn only
 * from that context.
 *
 * @param {string|Array} prompt The question or instruction.
 * @param {string|Array} [context] Optional source text to answer strictly from.
 * @return {string} The model's answer, or "" for an empty prompt.
 * @customfunction
 */
function ASK(prompt, context) {
  var q = flattenInput_(prompt);
  if (!q) return '';

  var ctx = flattenInput_(context);
  var fullPrompt;
  if (ctx) {
    fullPrompt =
      'Answer the question using ONLY the information in the context below. ' +
      'If the answer is not contained in the context, reply exactly: I don\'t know.\n\n' +
      'Context:\n' + ctx + '\n\n' +
      'Question:\n' + q;
  } else {
    fullPrompt = q;
  }

  return generate_('ASK', fullPrompt);
}

// ===========================================================================
// Core: cache -> call -> return (never throws)
// ===========================================================================

/**
 * Central entry point for every custom function. Handles the empty/no-key
 * guards, cache lookup, the API call, and caching of successful results.
 * Always returns a string — an answer, "", or a "⚠ …" error message.
 *
 * @param {string} funcTag Short tag identifying the calling function.
 * @param {string} prompt Fully-built prompt to send to Gemini.
 * @return {string}
 */
function generate_(funcTag, prompt) {
  if (!prompt) return '';

  var key = getApiKey_();
  if (!key) {
    return '⚠ No API key set. Use the SheetSense menu ▸ "Set / update Gemini API key".';
  }

  var model = getModel_();
  var cacheKey = makeCacheKey_(funcTag, model, prompt);
  var cache = getCache_();

  // 1) Serve from cache if we have it (avoids re-billing on recalculation).
  if (cache) {
    var cached = null;
    try { cached = cache.get(cacheKey); } catch (e) { /* ignore cache read errors */ }
    if (cached !== null && cached !== undefined) return cached;
  }

  // 2) Call the API (with bounded retry).
  var result = fetchGemini_(model, key, prompt);
  if (!result.ok) {
    return '⚠ ' + result.message;   // never throw to the cell
  }

  // 3) Cache the success (best-effort; values over ~100KB are simply not cached).
  if (cache) {
    try { cache.put(cacheKey, result.text, CACHE_TTL_SECONDS); } catch (e) { /* ignore */ }
  }
  return result.text;
}

/**
 * Performs the Gemini generateContent request with exponential-backoff retry
 * on transient errors (429 / 500 / 503). Uses muteHttpExceptions so non-2xx
 * responses are inspected rather than thrown.
 *
 * @param {string} model Model id, e.g. "gemini-2.5-flash".
 * @param {string} key Gemini API key.
 * @param {string} promptText Prompt to send.
 * @return {{ok: boolean, text: (string|undefined), message: (string|undefined)}}
 */
function fetchGemini_(model, key, promptText) {
  var url = ENDPOINT_BASE + encodeURIComponent(model) +
            ':generateContent?key=' + encodeURIComponent(key);

  var generationConfig = { temperature: TEMPERATURE };
  // gemini-2.5-flash* models "think" by default, which adds latency and cost.
  // These tasks are short and deterministic, so disable thinking to comfortably
  // stay under the ~30s custom-function limit. Only the flash tiers accept a
  // zero budget; gemini-2.5-pro and older 1.5 models would reject it, so guard.
  if (model.indexOf('2.5-flash') !== -1) {
    generationConfig.thinkingConfig = { thinkingBudget: 0 };
  }

  var payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: generationConfig
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var delay = INITIAL_BACKOFF_MS;
  for (var attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    var resp;
    try {
      resp = UrlFetchApp.fetch(url, options);
    } catch (e) {
      // Network-level failure (rare). Retry if attempts remain.
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(delay); delay *= 2; continue; }
      return { ok: false, message: 'Request failed: ' + truncate_(String(e && e.message || e), 160) };
    }

    var code = resp.getResponseCode();
    var body = resp.getContentText();

    if (code === 200) {
      return parseResponse_(body);
    }

    // Retry only on transient server/rate errors, and only if attempts remain.
    if ((code === 429 || code === 500 || code === 503) && attempt < MAX_ATTEMPTS) {
      Utilities.sleep(delay);
      delay *= 2;
      continue;
    }

    // Non-retryable code, or retries exhausted.
    return { ok: false, message: 'HTTP ' + code + ': ' + shortError_(body) };
  }

  return { ok: false, message: 'No response after ' + MAX_ATTEMPTS + ' attempts.' };
}

/**
 * Turns a 200 response body into {ok, text} or {ok:false, message}. Detects
 * blocked / empty content so those surface as readable messages, not crashes.
 *
 * @param {string} body Raw JSON response text.
 * @return {{ok: boolean, text: (string|undefined), message: (string|undefined)}}
 */
function parseResponse_(body) {
  var data;
  try {
    data = JSON.parse(body);
  } catch (e) {
    return { ok: false, message: 'Could not parse API response.' };
  }

  // Whole prompt blocked before generation.
  if (data.promptFeedback && data.promptFeedback.blockReason) {
    return { ok: false, message: 'Prompt blocked (' + data.promptFeedback.blockReason + ').' };
  }

  var candidates = data.candidates;
  if (!candidates || !candidates.length) {
    return { ok: false, message: 'No content returned.' };
  }

  var cand = candidates[0];
  var finish = cand.finishReason;
  if (finish === 'SAFETY' || finish === 'RECITATION' || finish === 'BLOCKLIST' || finish === 'PROHIBITED_CONTENT') {
    return { ok: false, message: 'Response blocked (' + finish + ').' };
  }

  var text = '';
  if (cand.content && cand.content.parts) {
    text = cand.content.parts.map(function (p) { return p.text || ''; }).join('').trim();
  }
  if (!text) {
    return { ok: false, message: 'Empty response' + (finish ? ' (' + finish + ')' : '') + '.' };
  }
  return { ok: true, text: text };
}

// ===========================================================================
// Helpers
// ===========================================================================

/**
 * Flattens any custom-function argument (string, number, or a range / nested
 * 2D array) into a single trimmed line of text. Empty cells are dropped.
 *
 * @param {*} input
 * @return {string}
 */
function flattenInput_(input) {
  if (input === null || input === undefined) return '';
  var parts = [];
  (function walk(v) {
    if (v === null || v === undefined || v === '') return;
    if (Array.isArray(v)) { v.forEach(walk); return; }
    parts.push(String(v));
  })(input);
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/**
 * Normalizes raw model output back to the exact provided label set:
 *   1) case-insensitive exact match,
 *   2) then containment (either direction),
 *   3) else the raw (cleaned) output.
 *
 * @param {string} output Raw model text.
 * @param {string[]} labels Allowed labels in their canonical casing.
 * @return {string}
 */
function normalizeToLabels_(output, labels) {
  // Trim surrounding whitespace, straight quotes, and periods. We deliberately
  // do NOT put a literal backtick in this character class: a backtick in the
  // source makes some editors think a template string is starting and paint the
  // rest of the file red. Backtick-wrapped output (e.g. the model returning a
  // fenced label) still matches via the containment check below.
  var cleaned = String(output).replace(/^[\s"'.]+|[\s"'.]+$/g, '').trim();
  var lower = cleaned.toLowerCase();

  // 1) exact (case-insensitive)
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].toLowerCase() === lower) return labels[i];
  }
  // 2) containment either direction
  for (var j = 0; j < labels.length; j++) {
    var ll = labels[j].toLowerCase();
    if (lower.indexOf(ll) !== -1 || ll.indexOf(lower) !== -1) return labels[j];
  }
  // 3) fall back to the raw output
  return cleaned;
}

/**
 * Builds the MD5-hashed cache key from CACHE_VERSION + model + tag + prompt.
 *
 * @param {string} funcTag
 * @param {string} model
 * @param {string} prompt
 * @return {string} 32-char hex digest (safe as a CacheService key).
 */
function makeCacheKey_(funcTag, model, prompt) {
  var raw = getCacheVersion_() + '|' + model + '|' + funcTag + '|' + prompt;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return bytesToHex_(digest);
}

/**
 * Converts a signed byte array (from computeDigest) to a lowercase hex string.
 *
 * @param {number[]} bytes
 * @return {string}
 */
function bytesToHex_(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    var s = b.toString(16);
    hex += (s.length === 1 ? '0' : '') + s;
  }
  return hex;
}

/**
 * Returns the document cache, falling back to the script cache. Both are
 * readable inside custom functions.
 *
 * @return {Cache}
 */
function getCache_() {
  try {
    var dc = CacheService.getDocumentCache();
    if (dc) return dc;
  } catch (e) { /* not in a document context */ }
  try {
    return CacheService.getScriptCache();
  } catch (e2) {
    return null;
  }
}

/** @return {string} The stored API key, or "" if unset. */
function getApiKey_() {
  var k = PropertiesService.getScriptProperties().getProperty(PROP_API_KEY);
  return k ? k.trim() : '';
}

/** @return {string} The configured model id, or the default. */
function getModel_() {
  var m = PropertiesService.getScriptProperties().getProperty(PROP_MODEL);
  return (m && m.trim()) ? m.trim() : DEFAULT_MODEL;
}

/** @return {string} The cache version (default "1"). */
function getCacheVersion_() {
  var v = PropertiesService.getScriptProperties().getProperty(PROP_CACHE_VERSION);
  return (v && v.trim()) ? v.trim() : '1';
}

/** @return {boolean} True if a returned string is one of our "⚠ …" errors. */
function isError_(s) {
  return typeof s === 'string' && s.charAt(0) === '⚠';  // ⚠
}

/**
 * Pulls a short, human-readable message out of a Gemini error body.
 *
 * @param {string} body
 * @return {string}
 */
function shortError_(body) {
  try {
    var data = JSON.parse(body);
    if (data && data.error && data.error.message) {
      return truncate_(data.error.message, 160);
    }
  } catch (e) { /* fall through */ }
  return truncate_(body, 160);
}

/**
 * Truncates a string to n characters with an ellipsis.
 *
 * @param {string} s
 * @param {number} n
 * @return {string}
 */
function truncate_(s, n) {
  s = String(s == null ? '' : s).trim();
  return s.length > n ? s.substring(0, n) + '…' : s;
}
