/**
 * ============================================================================
 *  SPG SHOPPER VOICE — Backend  (Google Apps Script)
 * ----------------------------------------------------------------------------
 *  Role of this file:
 *    - JSON API for the mobile front-end (hosted statically, see /web)
 *    - Google Sheets: SPG_MASTER, PRODUCT_MASTER, FEEDBACK_LOG
 *    - Google Drive : audio + transcripts (separate folder trees)
 *
 *  ZERO-COST DESIGN: there is no Google Cloud project, no billing account and
 *  no paid API. Sinhala speech recognition happens in the browser (Chrome's
 *  free Web Speech API); this backend only stores what the phone sends it.
 * ============================================================================
 */

const CONFIG = {
  // ---- Auto-filled by setupProject(). You may also paste them here. -------
  SPREADSHEET_ID: '',
  ROOT_FOLDER_ID: '',
  AUDIO_FOLDER_ID: '',
  TRANSCRIPT_FOLDER_ID: '',

  // ---- Server-side transcription -------------------------------------------
  //  SPGs only record. A background job transcribes and writes the .txt files.
  //  'NONE'   = no automatic transcription (transcripts stay blank)
  //  'GEMINI' = Gemini API. Cheapest by far; free tier needs no card BUT Google
  //             may use free-tier audio to improve its models. Paid tier does not.
  //  'CHIRP'  = Google Cloud Speech-to-Text v2, chirp_2, si-LK. Dedicated ASR.
  //  'GROQ'   = free Whisper. Weak at Sinhala (>30% WER). Last resort.
  //  Keys live in Script Properties, never in this file:
  //    GEMINI_API_KEY   GROQ_API_KEY      (CHIRP needs no key — see README)
  TRANSCRIBE_ENGINE: 'GEMINI',

  GEMINI_MODEL: 'gemini-2.5-flash',
  GEMINI_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/',

  GCP_PROJECT_ID: '',                  // CHIRP only
  STT_LOCATION: 'asia-southeast1',     // the only region serving Sinhala Chirp
  STT_MODEL: 'chirp_2',
  STT_FALLBACK_MODEL: 'chirp',
  SYNC_LIMIT_SECONDS: 58,              // CHIRP recognize() hard limit is 60s

  GROQ_MODEL: 'whisper-large-v3',
  GROQ_LANGUAGE: 'si',

  // ---- Background job ------------------------------------------------------
  BATCH_MAX_MINUTES: 4,                // stay under the 6-minute execution limit
  MAX_TRANSCRIBE_ATTEMPTS: 3,

  // ---- App rules ----------------------------------------------------------
  LANGUAGE_CODE: 'si-LK',        // used by the browser recogniser + file header
  MAX_RECORDING_SECONDS: 120,
  TIMEZONE: 'Asia/Colombo',
  APP_NAME: 'SPG Shopper Voice'
};

const SHEETS = {
  SPG: 'SPG_MASTER',
  PRODUCT: 'PRODUCT_MASTER',
  LOG: 'FEEDBACK_LOG'
};

const LOG_HEADERS = [
  'FEEDBACK_ID','SPG_CODE','SPG_NAME','PRODUCT_CODE','PRODUCT_NAME',
  'DATE','TIME','DURATION','AUDIO_FILE_ID','AUDIO_URL',
  'TRANSCRIPT_FILE_ID','TRANSCRIPT_URL','TRANSCRIPT','STATUS','CONSENT','ATTEMPTS'
];

/** Row states, in the order a recording moves through them. */
const ST = {
  DRAFT: 'DRAFT',                       // audio saved, SPG has not confirmed yet
  PENDING: 'PENDING_TRANSCRIPTION',     // confirmed, waiting for the batch job
  COMPLETED: 'COMPLETED',
  FAILED: 'TRANSCRIPTION_FAILED',
  DISCARDED: 'DISCARDED'
};

/* ===========================================================================
 *  CONFIG HELPERS  (Script Properties override the constants above)
 * ======================================================================== */
function cfg_(key) {
  const p = PropertiesService.getScriptProperties().getProperty(key);
  return (p !== null && p !== '') ? p : CONFIG[key];
}
function setCfg_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

/* ===========================================================================
 *  WEB ENTRY POINTS
 * ======================================================================== */

/**
 * GET — serves the app shell (desktop/debug only; the microphone does NOT
 * work inside the Apps Script iframe, see README) or answers simple API calls.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action) return jsonOut_(route_(p.action, p));
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST — the real API. The browser sends Content-Type: text/plain so the
 * request stays a CORS "simple request" and no preflight is needed
 * (Apps Script cannot answer OPTIONS preflights).
 */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'BAD_REQUEST', message: 'Could not read the request.' });
  }
  return jsonOut_(route_(body.action, body));
}

/** Used by google.script.run when the page is served from Apps Script. */
function apiBridge(payloadJson) {
  const body = JSON.parse(payloadJson || '{}');
  return JSON.stringify(route_(body.action, body));
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===========================================================================
 *  ROUTER
 * ======================================================================== */
function route_(action, data) {
  try {
    switch (action) {
      case 'ping':               return { ok: true, app: CONFIG.APP_NAME, time: nowParts_().iso };
      case 'validateSPG':        return validateSPG(data.spgCode);
      case 'getProducts':        return getProducts();
      case 'bootstrap':          return bootstrap(data.spgCode);
      case 'processFeedback':    return processFeedback(data);
      case 'confirmFeedback':    return confirmFeedback(data.feedbackId, data.spgCode, data.transcript);
      case 'discardFeedback':    return discardFeedback(data.feedbackId, data.spgCode);
      case 'getSPGHistory':      return getSPGHistory(data.spgCode, data.limit);
      case 'getFeedbackById':    return getFeedbackById(data.feedbackId, data.spgCode);
      case 'getAudio':           return getAudio(data.feedbackId, data.spgCode);
      default:
        return { ok: false, error: 'UNKNOWN_ACTION', message: 'That action is not available.' };
    }
  } catch (err) {
    console.error(action + ' :: ' + (err && err.stack ? err.stack : err));
    return { ok: false, error: 'SERVER_ERROR', message: 'Something went wrong. Please try again.' };
  }
}

/* ===========================================================================
 *  SHEET ACCESS
 * ======================================================================== */
function ss_() {
  const id = cfg_('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID is not set. Run setupProject() once.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}
function rows_(name) {
  const values = sheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(function (h) { return String(h).trim().toUpperCase(); });
  return values.slice(1).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}
function isTrue_(v) {
  return String(v).trim().toUpperCase() === 'TRUE' || v === true;
}

/* ===========================================================================
 *  1. LOGIN
 * ======================================================================== */
function validateSPG(spgCode) {
  const code = String(spgCode || '').trim().toUpperCase();
  if (!code) return { ok: false, error: 'EMPTY_CODE', message: 'Please enter your SPG code.' };

  const match = rows_(SHEETS.SPG).filter(function (r) {
    return String(r.SPG_CODE || '').trim().toUpperCase() === code;
  })[0];

  if (!match) return { ok: false, error: 'INVALID_SPG', message: 'We could not find that SPG code. Please check and try again.' };
  if (!isTrue_(match.ACTIVE)) return { ok: false, error: 'INACTIVE_SPG', message: 'This SPG code is not active. Please contact your supervisor.' };

  return {
    ok: true,
    spg: { code: code, name: String(match.SPG_NAME || code), region: String(match.REGION || '') }
  };
}

/* ===========================================================================
 *  2. PRODUCTS
 * ======================================================================== */
function getProducts() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('products_v1');
  if (hit) return { ok: true, products: JSON.parse(hit) };

  const list = rows_(SHEETS.PRODUCT)
    .filter(function (r) { return isTrue_(r.ACTIVE) && String(r.PRODUCT_CODE || '').trim(); })
    .map(function (r) {
      return { code: String(r.PRODUCT_CODE).trim(), name: String(r.PRODUCT_NAME || r.PRODUCT_CODE).trim() };
    });

  cache.put('products_v1', JSON.stringify(list), 300);
  return { ok: true, products: list };
}

/** One round trip on login: validate + products + today's count. */
function bootstrap(spgCode) {
  const v = validateSPG(spgCode);
  if (!v.ok) return v;
  const p = getProducts();
  const h = getSPGHistory(v.spg.code, 20);
  return {
    ok: true,
    spg: v.spg,
    products: p.products || [],
    todayCount: h.todayCount || 0,
    history: h.items || [],
    maxSeconds: Number(cfg_('MAX_RECORDING_SECONDS')),
    languageCode: cfg_('LANGUAGE_CODE')
  };
}

/* ===========================================================================
 *  3. FEEDBACK ID
 * ======================================================================== */
function createFeedbackId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const day = nowParts_().compactDate;              // 20260910
    const props = PropertiesService.getScriptProperties();
    const key = 'SEQ_' + day;
    const next = Number(props.getProperty(key) || '0') + 1;
    props.setProperty(key, String(next));
    return 'FB-' + day + '-' + ('000000' + next).slice(-6);
  } finally {
    lock.releaseLock();
  }
}

function nowParts_() {
  const tz = cfg_('TIMEZONE');
  const d = new Date();
  return {
    date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
    time: Utilities.formatDate(d, tz, 'HH:mm:ss'),
    compactDate: Utilities.formatDate(d, tz, 'yyyyMMdd'),
    yyyy: Utilities.formatDate(d, tz, 'yyyy'),
    mm: Utilities.formatDate(d, tz, 'MM'),
    dd: Utilities.formatDate(d, tz, 'dd'),
    iso: Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ssXXX")
  };
}

/* ===========================================================================
 *  4. DRIVE
 * ======================================================================== */
function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function datedFolder_(rootId, parts) {
  const root = DriveApp.getFolderById(rootId);
  return childFolder_(childFolder_(childFolder_(root, parts.yyyy), parts.mm), parts.dd);
}

function saveAudioToDrive(bytes, mimeType, fileName, parts) {
  const folder = datedFolder_(cfg_('AUDIO_FOLDER_ID'), parts);
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
  return { id: file.getId(), url: file.getUrl() };
}

function saveTranscriptToDrive(text, fileName, parts) {
  const folder = datedFolder_(cfg_('TRANSCRIPT_FOLDER_ID'), parts);
  const file = folder.createFile(Utilities.newBlob(text, 'text/plain', fileName));
  return { id: file.getId(), url: file.getUrl() };
}

function buildTranscriptText_(rec) {
  return [
    'Shopper Feedback',
    '',
    'Feedback ID: ' + rec.feedbackId,
    'SPG Code: '    + rec.spgCode,
    'SPG Name: '    + rec.spgName,
    'Product: '     + rec.productName + ' (' + rec.productCode + ')',
    'Date: '        + rec.date,
    'Time: '        + rec.time,
    'Duration: '    + rec.duration + ' sec',
    'Language: '    + cfg_('LANGUAGE_CODE'),
    'Source: '      + (rec.source || 'Confirmed by the SPG'),
    'Audio file: '  + (rec.hasAudio ? 'saved separately' : 'not recorded'),
    'Consent: '     + (rec.consent ? 'TRUE' : 'FALSE'),
    '',
    'Transcript:',
    '',
    rec.transcript || '(no transcript)',
    ''
  ].join('\n');
}

function guessExt_(mimeType) {
  const m = String(mimeType || '').toLowerCase();
  if (m.indexOf('webm') > -1) return '.webm';
  if (m.indexOf('ogg') > -1) return '.ogg';
  if (m.indexOf('mp4') > -1 || m.indexOf('m4a') > -1 || m.indexOf('aac') > -1) return '.m4a';
  if (m.indexOf('wav') > -1) return '.wav';
  return '.audio';
}

function safeName_(s) {
  return String(s || '').replace(/[^A-Za-z0-9඀-෿]+/g, '').slice(0, 40) || 'Product';
}

/* ===========================================================================
 *  5. TRANSCRIPTION ENGINES
 * ======================================================================== */

function setGeminiKey(key) { return saveKey_('GEMINI_API_KEY', key); }
function setGroqKey(key)   { return saveKey_('GROQ_API_KEY', key); }
function saveKey_(name, key) {
  PropertiesService.getScriptProperties().setProperty(name, String(key || '').trim());
  return name + ' saved.';
}
function key_(name) { return PropertiesService.getScriptProperties().getProperty(name) || ''; }

/**
 * Transcribe one audio blob. Returns { transcript, ok, detail }.
 * Never throws — a failure must never cost us the recording.
 */
function transcribeAudio_(blob, durationSec) {
  const engine = String(cfg_('TRANSCRIBE_ENGINE') || 'NONE').toUpperCase();
  try {
    if (engine === 'GEMINI') return transcribeGemini_(blob);
    if (engine === 'CHIRP')  return transcribeChirp_(blob, durationSec);
    if (engine === 'GROQ')   return transcribeGroq_(blob);
    return { ok: false, transcript: '', detail: 'No engine configured.' };
  } catch (err) {
    return { ok: false, transcript: '', detail: String(err && err.message ? err.message : err) };
  }
}

/* ---------- Gemini ------------------------------------------------------- */
const GEMINI_PROMPT =
  'Transcribe this audio verbatim in Sinhala script (සිංහල). ' +
  'This is a shopper in a Sri Lankan shop giving feedback about a product. ' +
  'Output ONLY the transcript text — no translation, no explanation, no quotation marks, no preamble. ' +
  'Keep English words the speaker actually used in English. Do not summarise or correct the speaker. ' +
  'If no speech is audible, output exactly: [no speech]';

function transcribeGemini_(blob) {
  const key = key_('GEMINI_API_KEY');
  if (!key) return { ok: false, transcript: '', detail: 'GEMINI_API_KEY is not set. Run setGeminiKey().' };

  const url = cfg_('GEMINI_ENDPOINT') + cfg_('GEMINI_MODEL') + ':generateContent?key=' + encodeURIComponent(key);
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{
        role: 'user',
        parts: [
          { text: GEMINI_PROMPT },
          { inline_data: { mime_type: blob.getContentType() || 'audio/webm',
                           data: Utilities.base64Encode(blob.getBytes()) } }
        ]
      }],
      generationConfig: { temperature: 0, candidateCount: 1 }
    })
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code === 429) return { ok: false, transcript: '', detail: 'Gemini rate limit / daily quota reached.' };
  if (code === 400 && body.indexOf('API key') > -1) return { ok: false, transcript: '', detail: 'Gemini rejected the API key.' };
  if (code !== 200) return { ok: false, transcript: '', detail: 'Gemini HTTP ' + code + ': ' + body.slice(0, 200) };

  let text = '';
  try {
    const c = (JSON.parse(body).candidates || [])[0];
    text = ((c && c.content && c.content.parts) || []).map(function (p) { return p.text || ''; }).join('').trim();
  } catch (e) {
    return { ok: false, transcript: '', detail: 'Could not read the Gemini response.' };
  }
  if (!text || text === '[no speech]') return { ok: false, transcript: '', detail: 'No speech detected in the recording.' };
  return { ok: true, transcript: text };
}

/* ---------- Google Cloud Speech-to-Text v2 (Chirp) ----------------------- */
function transcribeChirp_(blob, durationSec) {
  const proj = cfg_('GCP_PROJECT_ID');
  if (!proj) return { ok: false, transcript: '', detail: 'GCP_PROJECT_ID is not set.' };
  if (Number(durationSec) > Number(cfg_('SYNC_LIMIT_SECONDS'))) {
    return { ok: false, transcript: '', detail: 'Recording is longer than ' + cfg_('SYNC_LIMIT_SECONDS') +
             's. Chirp needs Cloud Storage staging for clips that long — use GEMINI instead.' };
  }

  const loc = cfg_('STT_LOCATION');
  const url = 'https://' + loc + '-speech.googleapis.com/v2/projects/' + proj +
              '/locations/' + loc + '/recognizers/_:recognize';
  const models = [cfg_('STT_MODEL'), cfg_('STT_FALLBACK_MODEL')];
  let last = '';

  for (let i = 0; i < models.length; i++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        config: {
          autoDecodingConfig: {},
          model: models[i],
          languageCodes: [cfg_('LANGUAGE_CODE')],
          features: { enableAutomaticPunctuation: true }
        },
        content: Utilities.base64Encode(blob.getBytes())
      })
    });
    if (res.getResponseCode() !== 200) { last = 'HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 180); continue; }
    const results = (JSON.parse(res.getContentText()) || {}).results || [];
    const text = results.map(function (r) {
      return (r.alternatives && r.alternatives[0] && r.alternatives[0].transcript) || '';
    }).join(' ').replace(/\s+/g, ' ').trim();
    if (text) return { ok: true, transcript: text };
    last = 'empty transcript';
  }
  return { ok: false, transcript: '', detail: last };
}

/* ---------- Groq / Whisper ---------------------------------------------- */
function transcribeGroq_(blob) {
  const key = key_('GROQ_API_KEY');
  if (!key) return { ok: false, transcript: '', detail: 'GROQ_API_KEY is not set.' };

  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/audio/transcriptions', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + key },
    muteHttpExceptions: true,
    payload: { file: blob, model: cfg_('GROQ_MODEL'), language: cfg_('GROQ_LANGUAGE'),
               response_format: 'json', temperature: '0' }
  });
  const code = res.getResponseCode();
  if (code === 429) return { ok: false, transcript: '', detail: 'Groq rate limit reached.' };
  if (code !== 200) return { ok: false, transcript: '', detail: 'Groq HTTP ' + code + ': ' + res.getContentText().slice(0, 180) };
  const text = String((JSON.parse(res.getContentText()) || {}).text || '').trim();
  return text ? { ok: true, transcript: text } : { ok: false, transcript: '', detail: 'Groq returned nothing.' };
}

/* ===========================================================================
 *  6. THE BACKGROUND JOB
 * ======================================================================== */

/**
 * Transcribes every recording waiting in FEEDBACK_LOG and writes its .txt file.
 * Safe to run repeatedly: it only touches PENDING_TRANSCRIPTION rows, stops
 * before the Apps Script execution limit, and leaves the rest for the next run.
 */
function runTranscriptionBatch() {
  const deadline = Date.now() + Number(cfg_('BATCH_MAX_MINUTES')) * 60 * 1000;
  const maxTries = Number(cfg_('MAX_TRANSCRIBE_ATTEMPTS'));
  const sh = sheet_(SHEETS.LOG);
  const values = sh.getDataRange().getValues();
  const col = {};
  LOG_HEADERS.forEach(function (h, i) { col[h] = i; });

  let done = 0, failed = 0, skipped = 0;

  for (let i = 1; i < values.length; i++) {
    if (Date.now() > deadline) { skipped = values.length - i; break; }

    const status = String(values[i][col.STATUS] || '');
    if (status !== ST.PENDING && status !== ST.FAILED) continue;

    const attempts = Number(values[i][col.ATTEMPTS] || 0);
    if (attempts >= maxTries) continue;

    const fileId = String(values[i][col.AUDIO_FILE_ID] || '').trim();
    if (!fileId) { updateLogCells_(i + 1, { STATUS: ST.FAILED, ATTEMPTS: maxTries }); failed++; continue; }

    const row = {};
    LOG_HEADERS.forEach(function (h, c) { row[h] = values[i][c]; });

    let out;
    try {
      out = transcribeAudio_(DriveApp.getFileById(fileId).getBlob(), Number(row.DURATION) || 0);
    } catch (err) {
      out = { ok: false, transcript: '', detail: String(err) };
    }

    if (!out.ok) {
      console.warn(row.FEEDBACK_ID + ' transcription failed: ' + out.detail);
      updateLogCells_(i + 1, { STATUS: ST.FAILED, ATTEMPTS: attempts + 1 });
      failed++;
      continue;
    }

    // Write the transcript as its own file, in its own folder tree.
    const d = String(row.DATE);
    const parts = { yyyy: d.slice(0, 4), mm: d.slice(5, 7), dd: d.slice(8, 10) };
    const fileName = row.FEEDBACK_ID + '_' + row.SPG_CODE + '_' + safeName_(row.PRODUCT_NAME) + '.txt';

    try {
      const tx = saveTranscriptToDrive(buildTranscriptText_({
        feedbackId: row.FEEDBACK_ID, spgCode: row.SPG_CODE, spgName: row.SPG_NAME,
        productCode: row.PRODUCT_CODE, productName: row.PRODUCT_NAME,
        date: row.DATE, time: row.TIME, duration: row.DURATION,
        hasAudio: true, consent: true,
        source: 'Automatic transcription (' + cfg_('TRANSCRIBE_ENGINE') + ')',
        transcript: out.transcript
      }), fileName, parts);

      updateLogCells_(i + 1, {
        TRANSCRIPT: out.transcript,
        TRANSCRIPT_FILE_ID: tx.id,
        TRANSCRIPT_URL: tx.url,
        STATUS: ST.COMPLETED,
        ATTEMPTS: attempts + 1
      });
      done++;
    } catch (err) {
      console.error(row.FEEDBACK_ID + ' transcript file failed: ' + err);
      updateLogCells_(i + 1, { TRANSCRIPT: out.transcript, STATUS: ST.FAILED, ATTEMPTS: attempts + 1 });
      failed++;
    }
  }

  const msg = 'Batch done — transcribed ' + done + ', failed ' + failed +
              (skipped ? ', ' + skipped + ' left for the next run' : '');
  console.log(msg);
  return msg;
}

/** Clear the attempt counter so failed rows are retried on the next run. */
function retryAllFailed() {
  const sh = sheet_(SHEETS.LOG);
  const values = sh.getDataRange().getValues();
  const statusCol = LOG_HEADERS.indexOf('STATUS');
  let n = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][statusCol]) === ST.FAILED) {
      updateLogCells_(i + 1, { STATUS: ST.PENDING, ATTEMPTS: 0 });
      n++;
    }
  }
  return 'Reset ' + n + ' failed rows. Run runTranscriptionBatch() or wait for the next trigger.';
}

/** Run once: transcribes new recordings every 15 minutes. */
function installTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runTranscriptionBatch').timeBased().everyMinutes(15).create();
  return 'Trigger installed — transcription runs every 15 minutes.';
}

/** Alternative: one pass each night at 1 a.m. */
function installNightlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runTranscriptionBatch').timeBased().atHour(1).everyDays(1)
    .inTimezone(cfg_('TIMEZONE')).create();
  return 'Nightly trigger installed — transcription runs at about 01:00.';
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runTranscriptionBatch') ScriptApp.deleteTrigger(t);
  });
  return 'Triggers removed.';
}

/** Editor helper: is the chosen engine actually reachable? */
function testEngine() {
  const engine = String(cfg_('TRANSCRIBE_ENGINE')).toUpperCase();
  const out = ['TRANSCRIBE_ENGINE : ' + engine];
  if (engine === 'GEMINI') {
    out.push('GEMINI_MODEL      : ' + cfg_('GEMINI_MODEL'));
    out.push('GEMINI_API_KEY    : ' + (key_('GEMINI_API_KEY') ? 'set' : 'NOT SET'));
    if (key_('GEMINI_API_KEY')) {
      const r = UrlFetchApp.fetch(cfg_('GEMINI_ENDPOINT') + cfg_('GEMINI_MODEL') +
        ':generateContent?key=' + encodeURIComponent(key_('GEMINI_API_KEY')), {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'Reply with OK' }] }] })
      });
      out.push('API reachable     : HTTP ' + r.getResponseCode());
    }
  } else if (engine === 'CHIRP') {
    out.push('GCP_PROJECT_ID    : ' + (cfg_('GCP_PROJECT_ID') || 'NOT SET'));
    out.push('Region / model    : ' + cfg_('STT_LOCATION') + ' / ' + cfg_('STT_MODEL'));
  } else if (engine === 'GROQ') {
    out.push('GROQ_API_KEY      : ' + (key_('GROQ_API_KEY') ? 'set' : 'NOT SET'));
  }
  const pending = rows_(SHEETS.LOG).filter(function (r) {
    return String(r.STATUS) === ST.PENDING || String(r.STATUS) === ST.FAILED; }).length;
  out.push('Waiting to process: ' + pending + ' recording(s)');
  const txt = out.join('\n');
  console.log(txt);
  return txt;
}

/* ===========================================================================
 *  7. MAIN FLOW
 * ======================================================================== */

/**
 * Step 1 — called as soon as the SPG stops recording.
 * Saves the audio to Drive FIRST (so it can never be lost), then logs the row
 * Nothing else happens here — the SPG must not wait on a transcription API.
 * STATUS = DRAFT until she taps Save.
 */
function processFeedback(data) {
  const v = validateSPG(data.spgCode);
  if (!v.ok) return v;

  if (!data.consent) {
    return { ok: false, error: 'NO_CONSENT', message: 'Please confirm that the shopper agreed to the recording.' };
  }
  if (!data.audioBase64 && !data.noAudio) {
    return { ok: false, error: 'NO_AUDIO', message: 'No recording was received. Please record again.' };
  }

  const maxS = Number(cfg_('MAX_RECORDING_SECONDS'));
  const duration = Math.min(Math.round(Number(data.duration) || 0), maxS);
  const mimeType = String(data.mimeType || 'audio/webm');
  let transcript = String(data.transcript || '').trim();
  const parts = nowParts_();
  const feedbackId = createFeedbackId();

  const productCode = String(data.productCode || '').trim();
  const productName = String(data.productName || productCode).trim();

  // ---- Persist the audio immediately (when there is any) ------------------
  let audio = { id: '', url: '' };
  if (data.audioBase64) {
    try {
      const bytes = Utilities.base64Decode(data.audioBase64);
      const fileName = feedbackId + '_' + v.spg.code + '_' + safeName_(productName) + guessExt_(mimeType);
      audio = saveAudioToDrive(bytes, mimeType, fileName, parts);
    } catch (err) {
      console.error('audio save failed: ' + err);
      return { ok: false, error: 'AUDIO_UPLOAD_FAILED', message: 'We could not save the recording. Please check your connection and try again.' };
    }
  }

  // ---- Log ----------------------------------------------------------------
  const row = {};
  LOG_HEADERS.forEach(function (h) { row[h] = ''; });
  Object.assign(row, {
    FEEDBACK_ID: feedbackId,
    SPG_CODE: v.spg.code,
    SPG_NAME: v.spg.name,
    PRODUCT_CODE: productCode,
    PRODUCT_NAME: productName,
    DATE: parts.date,
    TIME: parts.time,
    DURATION: duration,
    AUDIO_FILE_ID: audio.id,
    AUDIO_URL: audio.url,
    TRANSCRIPT: transcript,
    STATUS: ST.DRAFT,
    CONSENT: 'TRUE',
    ATTEMPTS: 0
  });
  appendLogRow_(row);

  return {
    ok: true,
    feedbackId: feedbackId,
    transcript: transcript,
    duration: duration,
    audioUrl: audio.url,
    date: parts.date,
    time: parts.time
  };
}

/**
 * Step 2 — the SPG taps "Save Feedback". She is done in three taps; the
 * transcript is produced later by runTranscriptionBatch().
 */
function confirmFeedback(feedbackId, spgCode, note) {
  const found = findLogRow_(feedbackId, spgCode);
  if (!found) return { ok: false, error: 'NOT_FOUND', message: 'That feedback record was not found.' };

  const engine = String(cfg_('TRANSCRIBE_ENGINE')).toUpperCase();
  const typed = String(note || '').trim();
  const patch = { STATUS: engine === 'NONE' ? ST.COMPLETED : ST.PENDING };
  if (typed) patch.TRANSCRIPT = typed;          // optional note the SPG typed herself
  updateLogCells_(found.index, patch);

  return { ok: true, feedbackId: feedbackId, status: patch.STATUS };
}

function discardFeedback(feedbackId, spgCode) {
  const found = findLogRow_(feedbackId, spgCode);
  if (!found) return { ok: false, error: 'NOT_FOUND', message: 'That feedback record was not found.' };
  try { DriveApp.getFileById(found.row.AUDIO_FILE_ID).setTrashed(true); } catch (ignore) {}
  updateLogCells_(found.index, { STATUS: ST.DISCARDED });
  return { ok: true };
}

/* ===========================================================================
 *  8. LOG HELPERS
 * ======================================================================== */
function appendLogRow_(obj) {
  const sh = sheet_(SHEETS.LOG);
  sh.appendRow(LOG_HEADERS.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

function findLogRow_(feedbackId, spgCode) {
  const id = String(feedbackId || '').trim();
  const code = String(spgCode || '').trim().toUpperCase();
  if (!id) return null;
  const sh = sheet_(SHEETS.LOG);
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]).trim() === id) {
      const row = {};
      LOG_HEADERS.forEach(function (h, c) { row[h] = values[i][c]; });
      if (code && String(row.SPG_CODE).trim().toUpperCase() !== code) return null;  // ownership check
      return { index: i + 1, row: row };
    }
  }
  return null;
}

function updateLogCells_(rowIndex, patch) {
  const sh = sheet_(SHEETS.LOG);
  Object.keys(patch).forEach(function (key) {
    const col = LOG_HEADERS.indexOf(key) + 1;
    if (col > 0) sh.getRange(rowIndex, col).setValue(patch[key]);
  });
}

/* ===========================================================================
 *  9. HISTORY
 * ======================================================================== */
function getSPGHistory(spgCode, limit) {
  const code = String(spgCode || '').trim().toUpperCase();
  if (!code) return { ok: false, error: 'EMPTY_CODE', message: 'Missing SPG code.' };
  const max = Math.min(Number(limit) || 30, 100);
  const today = nowParts_().date;

  const all = rows_(SHEETS.LOG).filter(function (r) {
    return String(r.SPG_CODE || '').trim().toUpperCase() === code &&
           String(r.STATUS) !== ST.DISCARDED && String(r.STATUS) !== ST.DRAFT;
  });

  const items = all.slice(-max).reverse().map(function (r) {
    return {
      feedbackId: String(r.FEEDBACK_ID),
      productCode: String(r.PRODUCT_CODE || ''),
      productName: String(r.PRODUCT_NAME || ''),
      date: String(r.DATE), time: String(r.TIME),
      duration: Number(r.DURATION) || 0,
      transcript: String(r.TRANSCRIPT || ''),
      hasAudio: !!String(r.AUDIO_FILE_ID || '').trim(),
      status: String(r.STATUS || '')
    };
  });

  const todayCount = all.filter(function (r) { return String(r.DATE) === today; }).length;
  return { ok: true, items: items, todayCount: todayCount, today: today };
}

function getFeedbackById(feedbackId, spgCode) {
  const found = findLogRow_(feedbackId, spgCode);
  if (!found) return { ok: false, error: 'NOT_FOUND', message: 'That feedback record was not found.' };
  return { ok: true, feedback: found.row };
}

/** Returns the stored audio as base64 so History can replay it privately. */
function getAudio(feedbackId, spgCode) {
  const found = findLogRow_(feedbackId, spgCode);
  if (!found) return { ok: false, error: 'NOT_FOUND', message: 'That recording was not found.' };
  if (!String(found.row.AUDIO_FILE_ID || '').trim()) {
    return { ok: false, error: 'NO_AUDIO', message: 'This feedback was saved as text only.' };
  }
  try {
    const blob = DriveApp.getFileById(found.row.AUDIO_FILE_ID).getBlob();
    return { ok: true, mimeType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { ok: false, error: 'AUDIO_UNAVAILABLE', message: 'We could not load that recording.' };
  }
}

/* ===========================================================================
 *  10. ONE-TIME SETUP  —  run this manually from the Apps Script editor
 * ======================================================================== */
function setupProject() {
  const root = DriveApp.createFolder(CONFIG.APP_NAME);
  const audio = root.createFolder('Audio');
  const transcripts = root.createFolder('Transcripts');

  const ss = SpreadsheetApp.create(CONFIG.APP_NAME + ' — Data');
  DriveApp.getFileById(ss.getId()).moveTo(root);

  const spg = ss.getActiveSheet().setName(SHEETS.SPG);
  spg.getRange('A1:D1').setValues([['SPG_CODE', 'SPG_NAME', 'REGION', 'ACTIVE']]);
  spg.getRange('A2:D4').setValues([
    ['SPG001', 'Demo SPG 01', 'Colombo', 'TRUE'],
    ['SPG002', 'Demo SPG 02', 'Gampaha', 'TRUE'],
    ['SPG003', 'Demo SPG 03', 'Kandy',   'TRUE']
  ]);

  const prod = ss.insertSheet(SHEETS.PRODUCT);
  prod.getRange('A1:C1').setValues([['PRODUCT_CODE', 'PRODUCT_NAME', 'ACTIVE']]);
  prod.getRange('A2:C4').setValues([
    ['P001', 'Product A', 'TRUE'],
    ['P002', 'Product B', 'TRUE'],
    ['P003', 'Product C', 'TRUE']
  ]);

  const log = ss.insertSheet(SHEETS.LOG);
  log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  log.setFrozenRows(1);

  [spg, prod, log].forEach(function (s) { s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold'); });
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  setCfg_('SPREADSHEET_ID', ss.getId());
  setCfg_('ROOT_FOLDER_ID', root.getId());
  setCfg_('AUDIO_FOLDER_ID', audio.getId());
  setCfg_('TRANSCRIPT_FOLDER_ID', transcripts.getId());

  const out = [
    '=== SETUP COMPLETE ===',
    'SPREADSHEET_ID      : ' + ss.getId(),
    'ROOT_FOLDER_ID      : ' + root.getId(),
    'AUDIO_FOLDER_ID     : ' + audio.getId(),
    'TRANSCRIPT_FOLDER_ID: ' + transcripts.getId(),
    'Spreadsheet URL     : ' + ss.getUrl(),
    'Drive folder URL    : ' + root.getUrl(),
    '',
    'NEXT: 1) set TRANSCRIBE_ENGINE + its API key  2) run installTrigger()',
    '      3) Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).'
  ].join('\n');
  console.log(out);
  return out;
}

/** Optional helper: point the script at an existing spreadsheet/folders. */
function setConfigValue(key, value) { setCfg_(key, value); return cfg_(key); }

/* ===========================================================================
 *  11. SELF-TEST  —  run from the editor to verify wiring
 * ======================================================================== */
function testBackend() {
  const report = [];
  function check(label, fn) {
    try { report.push('PASS  ' + label + '  ' + (fn() || '')); }
    catch (e) { report.push('FAIL  ' + label + '  ' + e.message); }
  }

  check('Spreadsheet reachable', function () { return ss_().getName(); });
  check('SPG_MASTER rows',       function () { return rows_(SHEETS.SPG).length + ' rows'; });
  check('PRODUCT_MASTER rows',   function () { return rows_(SHEETS.PRODUCT).length + ' rows'; });
  check('FEEDBACK_LOG headers',  function () { return sheet_(SHEETS.LOG).getLastColumn() + ' cols'; });
  check('Audio folder',          function () { return DriveApp.getFolderById(cfg_('AUDIO_FOLDER_ID')).getName(); });
  check('Transcript folder',     function () { return DriveApp.getFolderById(cfg_('TRANSCRIPT_FOLDER_ID')).getName(); });
  check('validateSPG(SPG001)',   function () { return JSON.stringify(validateSPG('SPG001')); });
  check('Feedback ID',           function () { return createFeedbackId(); });
  check('Log has ATTEMPTS col',  function () {
    const h = sheet_(SHEETS.LOG).getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
    if (String(h[LOG_HEADERS.length - 1]).toUpperCase() !== 'ATTEMPTS') throw new Error('add an ATTEMPTS header in the last column');
    return 'ok';
  });
  check('Transcription engine',  function () { return testEngine().split('\\n')[0]; });
  const out = report.join('\n');
  console.log(out);
  return out;
}
