/*
 * Thread Attachments - Outlook add-in (test build).
 *
 * Everything happens inside this side panel: attachments are read through Outlook's
 * add-in API and combined here with pdf-lib, PDF.js and UTIF.js. The only upload is the
 * optional "attach automatically" step, to the team admin's own PDF Emails Azure function.
 */

const VERSION = '0.8.0';
const SMALL_INLINE_BYTES = 50 * 1024; // logo test for images whose pixel size can't be read
const LOGO_MAX_SIDE = 400;            // images in the email body this small (in pixels) are logos/icons
const LOGO_THIN_SIDE = 120;           // ...and so are thin banners
const MAX_IMAGE_SIDE = 2400;          // photos are scaled to this before going into a PDF
const LOCKED_PDF_DPI = 200;           // locked PDFs are copied in as page images at this quality
const PREVIEW_PAGE_LIMIT = 40;
const THUMB_PX = 96;                  // list pictures are drawn this big and shown at 40 px (sharp on high-DPI screens)
const PENDING_KEY = 'pending-pdf';
const PENDING_MAX_AGE_MS = 2 * 60 * 60 * 1000;
const MB = 1000 * 1000;                // split sizes are in MB as people read them (not MiB)
const AZURE_FUNCTION = 'PdfEmails';    // the function's name in the team admin's Function App
const PDF_TYPES = [{ description: 'PDF file', accept: { 'application/pdf': ['.pdf'] } }];
const TEAM_FILE = 'team.json';
const MY_TEMPLATES_KEY = 'myTemplates'; // each person's own templates (Outlook roaming settings)
const DRAFT_KEY = 'teamDraft';          // the team admin's unpublished edits
const NEW_TEMPLATE = { body: 'Attached are the POD(s) for load {load}.', file: '{load}' };
const STOP_CHOICES = 9;                 // the stop picker offers at least Stop 1-9

const COMBINABLE = new Set(['PDF', 'Image', 'TIFF']);
const IMAGE_EXT = new Set(['jpg', 'jpeg', 'jfif', 'png', 'gif', 'bmp', 'webp']);
const OFFICE_EXT = new Set(['doc', 'docx', 'docm', 'rtf', 'odt', 'xls', 'xlsx', 'xlsm', 'xlsb', 'csv', 'ods', 'ppt', 'pptx']);
const MIME = { jpg: 'image/jpeg', jpeg: 'image/jpeg', jfif: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };
const KIND_LABEL = { PDF: 'PDF', Image: 'Image', TIFF: 'Scan (TIFF)', HEIC: 'iPhone photo', Office: 'Office file', Email: 'Email', Link: 'Cloud link', Other: 'File' };
const WHY_NOT = {
  HEIC: "iPhone .heic photos can't go into the PDF yet",
  Office: "Word and Excel files can't go into the PDF yet",
  Email: "Attached emails can't go into the PDF",
  Link: "Cloud links can't go into the PDF",
  Other: "This file type can't go into the PDF",
};
const EYE_SVG = '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true"><path d="M1.8 10S5 4.5 10 4.5 18.2 10 18.2 10 15 15.5 10 15.5 1.8 10 1.8 10Z" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><circle cx="10" cy="10" r="2.6" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>';

// File-type icons for the list (shown until a file's picture is ready, and for files with no picture)
const ICON_LINE = '#8a8a8a';
const OFFICE_COLOR = { doc: '#185abd', docx: '#185abd', docm: '#185abd', rtf: '#185abd', odt: '#185abd', xls: '#107c41', xlsx: '#107c41', xlsm: '#107c41', xlsb: '#107c41', csv: '#107c41', ods: '#107c41', ppt: '#c43e1c', pptx: '#c43e1c' };
const ENVELOPE_SVG = '<svg viewBox="0 0 28 28" aria-hidden="true"><rect x="3.5" y="7" width="21" height="14.5" rx="1.5" fill="#fff" stroke="#0f6cbd" stroke-width="1.3"/><path d="M4.2 8 14 15.5 23.8 8" fill="none" stroke="#0f6cbd" stroke-width="1.3" stroke-linejoin="round"/></svg>';
const CLOUD_SVG = '<svg viewBox="0 0 28 28" aria-hidden="true"><path d="M8.6 21h11.6a4.4 4.4 0 0 0 .5-8.8 6.2 6.2 0 0 0-11.9 1.7A3.6 3.6 0 0 0 8.6 21Z" fill="#fff" stroke="#0f6cbd" stroke-width="1.3" stroke-linejoin="round"/></svg>';

const $ = (id) => document.getElementById(id);

/** The team key from the install file (?team=… in the add-in's address). Read first thing. */
function readTeamKey() {
  try {
    const key = new URLSearchParams(window.location.search).get('team');
    return key && /^[A-Za-z0-9_-]{22}$/.test(key) ? key : null;
  } catch (error) {
    return null;
  }
}

const team = {
  key: readTeamKey(),
  status: 'unknown',   // nokey | nofile | ok | wrongkey | error
  published: null,     // the decrypted team.json
  draft: null,         // the admin's edits that aren't published yet
  me: '',              // this Outlook account's email address
  error: '',
  justPublished: false,
};

const state = {
  files: [],
  checked: new Set(),
  showLogos: false,
  topic: '',
  fileName: '',          // the name typed in "Name the PDF" (for the current thread)
  loadNumber: '',        // the load number typed for an email template (for the current thread)
  accessorials: emptyAccessorials(), // Detention / Lumper / Other typed for the current thread
  myTemplates: [],
  saveSplitMb: 0,        // "Split into files of up to" for Save PDF (0 = no limit)
  multiStop: false,      // Multi-stop: one PDF per stop, named -1, -2… (for the current thread)
  messages: 0,
  failed: 0,
  loadToken: 0,
  busy: false,
  reloadPending: false,
  built: new Map(),      // PDFs already built for the current selection, so Save then Email doesn't rebuild
  previewIndex: -1,
  previewToken: 0,
  env: [],
  log: [],
};

// --------------------------------------------------------------------------- small helpers

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function userError(message) {
  const error = new Error(message);
  error.userMessage = message;
  return error;
}

/** Office.js callback API -> Promise. */
function call(start) {
  return new Promise((resolve, reject) => {
    try {
      start((result) => {
        if (result && result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error((result.error && result.error.message) || 'Outlook reported an error'));
        } else {
          resolve(result ? result.value : undefined);
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

function base64ToBytes(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function bytesToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

function humanSize(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

function plural(n, word, many) {
  return `${n} ${n === 1 ? word : (many || `${word}s`)}`;
}

function extOf(name) {
  const match = /\.([^.]+)$/.exec(name || '');
  return match ? match[1].toLowerCase() : '';
}

function stripPrefixes(subject) {
  return String(subject || '').replace(/^\s*((re|fw|fwd|aw|wg|sv|vs|tr)\s*(\[\d+\])?\s*:\s*)+/i, '').trim();
}

function safeName(text) {
  const name = String(text || '').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/, '').trim();
  return name.slice(0, 120) || 'Combined attachments';
}

/** Save PDF's file name: the name you typed for this thread, else the load number, else the subject. */
function pdfName() {
  return `${safeName(state.fileName || state.loadNumber || guessLoadNumber(state.topic) || state.topic || 'Combined attachments')}.pdf`;
}

/** What was typed in "Name the PDF", made safe for Windows. Blank means "use the default name". */
function cleanFileName(text) {
  const name = String(text || '').trim().replace(/\.pdf$/i, '').trim();
  return name ? safeName(name) : '';
}

function readSetting(key) {
  try {
    return localStorage.getItem(`thread-attachments.${key}`);
  } catch (error) {
    return null;
  }
}

function writeSetting(key, value) {
  try {
    if (value === null) localStorage.removeItem(`thread-attachments.${key}`);
    else localStorage.setItem(`thread-attachments.${key}`, value);
  } catch (error) {
    // not remembered; harmless
  }
}

function fmtDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  const day = date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
  const time = date.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  return `${day}, ${time}`;
}

function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

function supports(version) {
  try {
    return Office.context.requirements.isSetSupported('Mailbox', version);
  } catch (error) {
    return false;
  }
}

function kindOf(name, contentType, attachmentType) {
  const type = String(attachmentType || '').toLowerCase();
  if (type === 'item') return 'Email';
  if (type === 'cloud') return 'Link';
  const ext = extOf(name);
  const mime = String(contentType || '').toLowerCase();
  if (ext === 'pdf' || mime === 'application/pdf') return 'PDF';
  if (ext === 'tif' || ext === 'tiff' || mime === 'image/tiff') return 'TIFF';
  if (IMAGE_EXT.has(ext) || /^image\/(jpeg|png|gif|bmp|webp)$/.test(mime)) return 'Image';
  if (ext === 'heic' || ext === 'heif' || /^image\/hei[cf]$/.test(mime)) return 'HEIC';
  if (OFFICE_EXT.has(ext)) return 'Office';
  if (ext === 'msg' || ext === 'eml') return 'Email';
  return 'Other';
}

function canCombine(file) {
  return COMBINABLE.has(file.kind) && !!file.bytes;
}

// --------------------------------------------------------------------------- status, report

function setSummary(text) {
  $('summary').textContent = text || '';
}

function setStatus(text, tone) {
  const node = $('status');
  node.textContent = text || '';
  node.className = `status${tone ? ` ${tone}` : ''}`;
  clearParts();
}

/** After a split save: a button per file, in case Outlook held back the later downloads. */
function showParts(files) {
  const box = $('parts');
  box.textContent = '';
  box.append(el('span', 'muted', "Didn't get them all? Click a file to download it:"));
  for (const file of files) {
    const button = el('button', 'link-btn', file.name);
    button.type = 'button';
    button.addEventListener('click', () => {
      try {
        downloadBytes(file.bytes, file.name);
        log(`Download: ${file.name} again`);
      } catch (error) {
        fail(error);
      }
    });
    box.append(button);
  }
  box.hidden = false;
}

function clearParts() {
  const box = $('parts');
  if (!box || box.hidden) return;
  box.hidden = true;
  box.textContent = '';
}

function log(text) {
  const stamp = new Date().toTimeString().slice(0, 8);
  state.log.push(`${stamp}  ${text}`);
  renderReport();
}

function renderReport() {
  $('report').value = [...state.env, '', ...state.log].join('\n');
}

function collectEnvironment() {
  const lines = [`Thread Attachments ${VERSION} (test build)`];
  try {
    const diag = Office.context.diagnostics || {};
    const mailDiag = (Office.context.mailbox && Office.context.mailbox.diagnostics) || {};
    lines.push(`Outlook: ${mailDiag.hostName || diag.host || '?'} ${mailDiag.hostVersion || diag.version || ''} on ${diag.platform || Office.context.platform || '?'}`);
    lines.push(`Mailbox API 1.8 / 1.13 / 1.15: ${['1.8', '1.13', '1.15'].map((v) => (supports(v) ? 'yes' : 'no')).join(' / ')}`);
  } catch (error) {
    lines.push(`Outlook details unavailable: ${error.message}`);
  }
  let framed = 'unknown';
  try {
    framed = window.self !== window.top ? 'yes' : 'no';
  } catch (error) {
    framed = 'yes';
  }
  lines.push(`Inside a frame: ${framed} · Save dialog API: ${typeof window.showSaveFilePicker === 'function' ? 'present' : 'missing'}`);
  lines.push(`Browser: ${navigator.userAgent}`);
  state.env = lines;
  renderReport();
}

async function copyReport() {
  const text = $('report').value;
  try {
    await navigator.clipboard.writeText(text);
    flashCopy('Copied');
    return;
  } catch (error) {
    // fall back to selecting the text
  }
  const box = $('report');
  box.focus();
  box.select();
  let copied = false;
  try {
    copied = document.execCommand('copy');
  } catch (error) {
    copied = false;
  }
  flashCopy(copied ? 'Copied' : 'Press Ctrl+C to copy');
}

function flashCopy(text) {
  const button = $('btn-copy-report');
  button.textContent = text;
  setTimeout(() => { button.textContent = 'Copy report'; }, 2500);
}

function fail(error) {
  const message = (error && (error.userMessage || error.message)) || String(error);
  log(`Error: ${message}`);
  setStatus(message, 'error');
}

// --------------------------------------------------------------------------- views

function showView(name) {
  for (const id of ['read-view', 'compose-view', 'message-view']) $(id).hidden = id !== name;
  $('actions').hidden = name !== 'read-view';
}

function showMessage(text) {
  $('message-text').textContent = text;
  showView('message-view');
}

// --------------------------------------------------------------------------- reading the thread

let loadChain = Promise.resolve();

/**
 * Reloads the list for the current selection. Runs one load at a time (Outlook lets an
 * add-in open only one selected email at a time), and the newest request wins.
 */
function loadThread() {
  if (state.busy) {
    state.reloadPending = true;
    return Promise.resolve();
  }
  const token = ++state.loadToken;
  const run = loadChain.then(() => (token === state.loadToken ? loadThreadNow(token) : undefined));
  loadChain = run.catch(() => {});
  return run;
}

async function loadThreadNow(token) {
  closePreview();
  hideNameForm();
  hideLoadForm();
  closeEmailMenu(false);
  state.fileName = '';
  state.loadNumber = '';
  state.accessorials = emptyAccessorials();
  state.multiStop = false; // a new thread starts as a single-stop load
  state.built.clear();
  setStatus('');
  setSummary('Reading the selected emails…');

  const mailbox = Office.context.mailbox;
  let selected = [];
  if (supports('1.13') && typeof mailbox.getSelectedItemsAsync === 'function') {
    try {
      selected = (await call((done) => mailbox.getSelectedItemsAsync(done))) || [];
    } catch (error) {
      log(`Couldn't read the selection: ${error.message}`);
    }
  }

  const found = [];
  const subjects = [];
  let failed = 0;
  let messages = 0;

  if (selected.length && supports('1.15') && typeof mailbox.loadItemByIdAsync === 'function') {
    for (let i = 0; i < selected.length; i++) {
      if (token !== state.loadToken) return;
      const pick = selected[i];
      messages += 1;
      if (pick.subject) subjects.push(pick.subject);
      if (String(pick.itemMode || '').toLowerCase() === 'compose') {
        log(`Skipped a draft: ${pick.subject || '(no subject)'}`);
        continue;
      }
      // No skipping on hasAttachment: classic Outlook reports false when the only
      // attachments are pictures pasted into the body.
      setSummary(`Reading email ${i + 1} of ${selected.length}…`);
      let item;
      try {
        item = await call((done) => mailbox.loadItemByIdAsync(pick.itemId, done));
      } catch (error) {
        failed += 1;
        log(`Couldn't open "${pick.subject || 'an email'}": ${error.message}`);
        continue;
      }
      try {
        failed += await readAttachments(item, found, token, `Email ${i + 1} (paperclip: ${pick.hasAttachment === false ? 'no' : pick.hasAttachment ? 'yes' : '?'})`);
      } finally {
        try {
          await call((done) => item.unloadAsync(done));
        } catch (error) {
          log(`Couldn't close an email after reading it: ${error.message}`);
        }
      }
    }
    log(`Read ${plural(selected.length, 'selected email')} (multi-select)`);
  } else if (mailbox.item) {
    messages = 1;
    if (mailbox.item.subject) subjects.push(mailbox.item.subject);
    failed += await readAttachments(mailbox.item, found, token, 'Open email');
    log(`Read the open email${selected.length > 1 ? ' (multi-select not available here)' : ''}`);
  } else {
    state.files = [];
    setSummary('');
    showMessage('Select an email, or click a whole conversation, to see its attachments.');
    return;
  }
  if (token !== state.loadToken) return;

  // Oldest first, so a file sent around several times is credited to whoever sent it first.
  found.sort((a, b) => (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0));
  const files = await mergeDuplicates(found);
  if (token !== state.loadToken) return;
  files.forEach((file, index) => {
    file.id = index;
    file.smallLogo = looksLikeLogo(file);
  });

  state.files = files;
  state.messages = messages;
  state.failed = failed;
  state.topic = stripPrefixes(subjects[0] || '');
  state.checked = new Set(); // nothing checked until you pick (or Select all)
  $('title').textContent = state.topic || 'Thread Attachments';
  renderSummary();

  if (!files.length) {
    showMessage(`No attachments in ${messages === 1 ? 'this email' : `these ${messages} emails`}.`);
    return;
  }
  showView('read-view');
  renderList();
  const dupes = files.reduce((sum, f) => sum + f.copies - 1, 0);
  log(`Found ${plural(files.length, 'file')}${dupes ? `, ${plural(dupes, 'duplicate')} merged` : ''}${failed ? `, ${failed} unreadable` : ''}`);
  fillThumbs(token).catch((error) => log(`Pictures: stopped (${error.message})`));
}

// --------------------------------------------------------------------------- picking pages

/** 0-based pages of a PDF that go into the combined file, or null for "every page". */
function includedPages(file) {
  if (file.kind !== 'PDF' || !file.pageChoice) return null;
  return [...file.pageChoice].sort((a, b) => a - b);
}

/** [0, 1, 3] -> "Pages 1–2, 4" */
function pagesLabel(pages) {
  const sorted = [...pages].sort((a, b) => a - b).map((i) => i + 1);
  const parts = [];
  for (let i = 0; i < sorted.length; i++) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    parts.push(j > i ? `${sorted[i]}–${sorted[j]}` : String(sorted[i]));
    i = j;
  }
  return `${sorted.length === 1 ? 'Page' : 'Pages'} ${parts.join(', ')}`;
}

function pageNote(file) {
  if (!file.pageChoice) return null;
  return `${pagesLabel([...file.pageChoice])} of ${file.pages} picked`;
}

// --------------------------------------------------------------------------- Multi-stop: which stop each file (or page) is

/** Every page of a PDF the file contributes, 0-based, or null when its page count isn't known. */
function allIncludedPages(file) {
  const picked = includedPages(file);
  if (picked) return picked;
  return file.pages ? Array.from({ length: file.pages }, (_, i) => i) : null;
}

/** Stop -> pages (null = all picked pages) for one checked file. Pages can go to different stops. */
function stopsOfFile(file) {
  const map = new Map();
  const base = file.stop || 1;
  const pages = file.pageStops && file.pageStops.size ? allIncludedPages(file) : null;
  if (!pages) {
    map.set(base, includedPages(file));
    return map;
  }
  for (const p of pages) {
    const stop = file.pageStops.get(p) || base;
    if (!map.has(stop)) map.set(stop, []);
    map.get(stop).push(p);
  }
  return map;
}

/** [{ stop, entries: [{ file, only }] }] in stop order; files keep their list order within a stop. */
function stopEntries() {
  const byStop = new Map();
  for (const file of checkedFiles()) {
    for (const [stop, only] of stopsOfFile(file)) {
      if (!byStop.has(stop)) byStop.set(stop, []);
      byStop.get(stop).push({ file, only });
    }
  }
  return [...byStop.entries()].sort((a, b) => a[0] - b[0]).map(([stop, entries]) => ({ stop, entries }));
}

/** Stop numbers used by the checked files that have one (leaving out `except`). */
function stopsInUse(except) {
  const used = new Set();
  for (const file of checkedFiles()) {
    if (file === except || !file.stop) continue;
    for (const stop of stopsOfFile(file).keys()) used.add(stop);
  }
  return used;
}

/** Gives a newly checked file the lowest stop number nobody has yet (1, then 2…). */
function assignStop(file) {
  const used = stopsInUse(file);
  let stop = 1;
  while (used.has(stop)) stop += 1;
  file.stop = stop;
  file.pageStops = null;
}

/** The highest stop the pickers should offer. */
function stopChoices() {
  return Math.max(STOP_CHOICES, ...[...stopsInUse()].map((s) => s + 1));
}

/** When every picked page of a file ends up in one stop, it's simply that file's stop again. */
function normalizePageStops(file) {
  if (!file.pageStops || !file.pageStops.size) {
    file.pageStops = null;
    return;
  }
  const stops = [...stopsOfFile(file).keys()];
  if (stops.length <= 1) {
    file.stop = stops[0] || file.stop;
    file.pageStops = null;
    return;
  }
  for (const [page, stop] of [...file.pageStops]) if (stop === file.stop) file.pageStops.delete(page);
}

/** "Stop 1: pages 1–2 · Stop 2: page 3" for a file split across stops, else null. */
function stopNote(file) {
  if (!state.multiStop || !state.checked.has(file.id) || !file.pageStops) return null;
  const map = stopsOfFile(file);
  if (map.size < 2) return null;
  return [...map.entries()].sort((a, b) => a[0] - b[0]).map(([stop, pages]) => `Stop ${stop}: ${pagesLabel(pages).toLowerCase()}`).join(' · ');
}

/** A stop picker (Stop 1…n). `mixed` adds a "By page" choice for a file split across stops. */
function stopPicker(value, label, mixed, onPick) {
  const select = document.createElement('select');
  select.className = 'stop-pick';
  select.setAttribute('aria-label', label);
  if (mixed) {
    const option = new Option('By page', '', true, true);
    option.disabled = true;
    select.append(option);
  }
  for (let stop = 1; stop <= stopChoices(); stop++) select.append(new Option(`Stop ${stop}`, String(stop), false, !mixed && stop === value));
  select.disabled = state.busy;
  select.addEventListener('change', () => onPick(Number(select.value)));
  return select;
}

/** Picking a stop for a whole file (this also undoes any page-by-page stops). */
function setFileStop(file, stop) {
  file.stop = stop;
  file.pageStops = null;
  state.built.clear();
  log(`Multi-stop: ${file.name} -> stop ${stop}`);
  refreshRow(file);
  if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
}

/** Picking a stop for one page in the preview (for a PDF that holds more than one stop's POD). */
function setPageStop(file, page, stop) {
  if (!file.pageStops) file.pageStops = new Map();
  file.pageStops.set(page, stop);
  normalizePageStops(file);
  state.built.clear();
  refreshRow(file);
  syncPreviewPages();
}

function onMultiStopToggle(on) {
  state.multiStop = on;
  state.built.clear();
  if (on) {
    for (const file of checkedFiles()) {
      file.stop = 0;
      file.pageStops = null;
    }
    for (const file of checkedFiles()) assignStop(file); // in list order: 1, 2, 3…
  }
  log(`Multi-stop: ${on ? 'on' : 'off'}`);
  renderList();
  if (!$('preview').hidden) syncPreviewPages();
  if (!$('name-form').hidden) updateNameHint();
  updateLoadHint();
}

/**
 * Signature logos and icons: small pictures in the email body. Judged by pixel size, so a
 * photo of a mostly white page (which can be a very small file) still counts as a photo.
 */
function looksLikeLogo(file) {
  if (!file.inline || (file.kind !== 'Image' && file.kind !== 'TIFF')) return false;
  const dims = file.kind === 'Image' && file.bytes ? imageSize(file.bytes) : null;
  if (dims) {
    file.dims = dims;
    return Math.max(dims[0], dims[1]) <= LOGO_MAX_SIDE || Math.min(dims[0], dims[1]) <= LOGO_THIN_SIDE;
  }
  return file.size < SMALL_INLINE_BYTES;
}

/** [width, height] from a PNG, GIF, JPEG, BMP or WebP header, or null. */
function imageSize(b) {
  if (!b || b.length < 30) return null;
  const be16 = (i) => (b[i] << 8) | b[i + 1];
  const le16 = (i) => b[i] | (b[i + 1] << 8);
  const be32 = (i) => b[i] * 0x1000000 + ((b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]);
  const le32s = (i) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24);
  const ok = (w, h) => (w > 0 && h > 0 ? [w, h] : null);
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return ok(be32(16), be32(20)); // PNG
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return ok(le16(6), le16(8)); // GIF
  if (b[0] === 0x42 && b[1] === 0x4d) return ok(Math.abs(le32s(18)), Math.abs(le32s(22))); // BMP
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45) { // WebP
    const chunk = String.fromCharCode(b[12], b[13], b[14], b[15]);
    if (chunk === 'VP8X') return ok(1 + (b[24] | (b[25] << 8) | (b[26] << 16)), 1 + (b[27] | (b[28] << 8) | (b[29] << 16)));
    if (chunk === 'VP8 ') return ok(le16(26) & 0x3fff, le16(28) & 0x3fff);
    if (chunk === 'VP8L') {
      const bits = le32s(21);
      return ok(1 + (bits & 0x3fff), 1 + ((bits >> 14) & 0x3fff));
    }
    return null;
  }
  if (b[0] === 0xff && b[1] === 0xd8) { // JPEG: find the frame header
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i += 1; continue; }
      const marker = b[i + 1];
      if (marker === 0xff) { i += 1; continue; }
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) { i += 2; continue; }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return ok(be16(i + 7), be16(i + 5));
      }
      i += 2 + be16(i + 2);
    }
  }
  return null;
}

/** Reads one email's attachments into `found`. Returns how many couldn't be read. */
async function readAttachments(item, found, token, label) {
  let failed = 0;
  const from = item.from ? (item.from.displayName || item.from.emailAddress || '') : '';
  const date = item.dateTimeCreated ? new Date(item.dateTimeCreated) : null;
  let attachments = item.attachments;
  if (!attachments && typeof item.getAttachmentsAsync === 'function') {
    try {
      attachments = await call((done) => item.getAttachmentsAsync(done));
    } catch (error) {
      attachments = [];
    }
  }
  const list = attachments || [];
  const inBody = list.filter((att) => att.isInline).length;
  log(`${label || 'Email'}: Outlook listed ${plural(list.length, 'attachment')}${inBody ? ` (${inBody} in the body)` : ''}`);
  for (const att of list) {
    if (token !== state.loadToken) return failed;
    const file = {
      name: att.name || 'attachment',
      contentType: att.contentType || '',
      size: att.size || 0,
      inline: !!att.isInline,
      from,
      date,
      kind: kindOf(att.name, att.contentType, att.attachmentType),
      bytes: null,
      copies: 1,
    };
    if (file.kind === 'Link') {
      found.push(file);
      continue;
    }
    try {
      const content = await call((done) => item.getAttachmentContentAsync(att.id, done));
      const format = String(content.format || '').toLowerCase();
      if (format === 'base64') {
        file.bytes = base64ToBytes(content.content);
      } else if (format === 'eml' || format === 'icalendar') {
        file.bytes = new TextEncoder().encode(content.content);
        if (format === 'eml') {
          file.kind = 'Email';
          if (!/\.(eml|msg)$/i.test(file.name)) file.name += '.eml';
        }
      } else if (format === 'url') {
        file.kind = 'Link';
      }
      if (file.bytes) file.size = file.bytes.length;
      found.push(file);
    } catch (error) {
      failed += 1;
      log(`Couldn't read "${file.name}": ${error.message}`);
    }
  }
  return failed;
}

async function fingerprint(bytes) {
  try {
    const digest = await crypto.subtle.digest('SHA-256', bytes);
    return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('');
  } catch (error) {
    const n = bytes.length;
    return `${n}:${bytes[0]}:${bytes[n >> 1]}:${bytes[n - 1]}`;
  }
}

/** One entry per identical file. Keeps the first copy's sender and date. */
async function mergeDuplicates(found) {
  const unique = [];
  const byKey = new Map();
  for (const file of found) {
    const key = file.bytes ? await fingerprint(file.bytes) : `${file.kind}:${file.name}`;
    const first = byKey.get(key);
    if (!first) {
      byKey.set(key, file);
      unique.push(file);
    } else {
      first.copies += 1;
      first.inline = first.inline && file.inline; // a real attachment anywhere wins
    }
  }
  return unique;
}

// --------------------------------------------------------------------------- the list

function visibleFiles() {
  return state.files.filter((f) => state.showLogos || !f.smallLogo);
}

function checkedFiles() {
  return visibleFiles().filter((f) => canCombine(f) && state.checked.has(f.id));
}

function metaLine(file) {
  const parts = [KIND_LABEL[file.kind] || 'File', humanSize(file.size)];
  if (file.from) parts.push(file.from);
  const when = fmtDate(file.date);
  if (when) parts.push(when);
  return parts.join(' · ');
}

function renderSummary() {
  const parts = [`${plural(state.files.length, 'file')} from ${plural(state.messages, 'email')}`];
  const dupes = state.files.reduce((sum, f) => sum + f.copies - 1, 0);
  if (dupes) parts.push(`${plural(dupes, 'duplicate')} merged`);
  if (state.failed) parts.push(`${state.failed} couldn't be read`);
  setSummary(parts.join(' · '));
}

function renderList() {
  const list = $('file-list');
  list.textContent = '';
  for (const file of visibleFiles()) list.append(fileRow(file));
  const logos = state.files.filter((f) => f.smallLogo).length;
  $('logos-toggle').hidden = logos === 0;
  $('logo-count').textContent = String(logos);
  $('show-logos').checked = state.showLogos;
  updateSelectionUi();
}

function fileRow(file) {
  const row = el('li', `file${canCombine(file) ? '' : ' locked'}`);
  row.dataset.id = String(file.id);
  const label = el('label', 'file-label');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.checked = state.checked.has(file.id);
  box.disabled = !canCombine(file) || state.busy;
  box.setAttribute('aria-label', `Include ${file.name}`);
  box.addEventListener('change', () => onRowToggle(file, box.checked));

  const text = el('span', 'file-text');
  const nameRow = el('span', 'file-name-row');
  const name = el('span', 'file-name', file.name);
  name.title = file.name;
  nameRow.append(name);
  if (file.copies > 1) nameRow.append(el('span', 'chip', `×${file.copies}`));
  text.append(nameRow, el('span', 'file-meta', metaLine(file)));
  if (!canCombine(file)) text.append(el('span', 'file-note', WHY_NOT[file.kind] || "Can't go into the PDF"));
  const note = stopNote(file) || pageNote(file);
  if (note) text.append(el('span', 'file-note picked', note));

  // A small picture of the file, so a POD scan and a photo of the truck look different at a glance
  const thumb = el('span', 'thumb');
  thumb.setAttribute('aria-hidden', 'true'); // the eye button is the accessible way to preview
  fillThumbSlot(thumb, file);
  if (canCombine(file)) {
    thumb.title = 'Preview';
    thumb.addEventListener('click', (event) => {
      event.preventDefault(); // previews instead of ticking the file
      openPreview(file.id);
    });
  }
  label.append(box, thumb, text);

  const preview = document.createElement('button');
  preview.type = 'button';
  preview.className = 'icon-btn';
  preview.title = 'Preview';
  preview.setAttribute('aria-label', `Preview ${file.name}`);
  preview.innerHTML = EYE_SVG;
  preview.addEventListener('click', () => openPreview(file.id));
  row.append(label);
  // Multi-stop: which stop this file is (kept outside the label so picking doesn't tick the box)
  if (state.multiStop && canCombine(file) && state.checked.has(file.id)) {
    const mixed = !!(file.pageStops && file.pageStops.size);
    row.append(stopPicker(file.stop || 1, `Stop for ${file.name}`, mixed, (stop) => setFileStop(file, stop)));
  }
  row.append(preview);
  return row;
}

/** Files "Select all" works on. */
function selectableFiles() {
  return visibleFiles().filter(canCombine);
}

function updateSelectionUi() {
  const selectable = selectableFiles();
  const count = selectable.filter((f) => state.checked.has(f.id)).length;
  const all = $('check-all');
  all.disabled = state.busy || selectable.length === 0;
  all.checked = count > 0 && count === selectable.length;
  all.indeterminate = count > 0 && count < selectable.length;
  const stops = state.multiStop ? stopsInUse().size : 0;
  $('selected-count').textContent = selectable.length
    ? `${count} of ${selectable.length} selected${stops ? ` · ${plural(stops, 'stop')}` : ''}`
    : 'Nothing here can go into a PDF';
  $('multi-stop').checked = state.multiStop;
  $('multi-stop').disabled = state.busy;
  const anyChecked = checkedFiles().length > 0;
  for (const id of ['btn-save', 'name-save', 'load-go']) $(id).disabled = state.busy || !anyChecked;
  $('btn-save').title = anyChecked ? '' : 'Check the files to save first.';
  $('btn-email').disabled = state.busy; // still opens, to reach Manage templates
  for (const box of document.querySelectorAll('#file-list input[type="checkbox"]')) {
    const file = state.files[Number(box.closest('.file').dataset.id)];
    box.disabled = state.busy || !canCombine(file);
    box.checked = state.checked.has(file.id);
  }
  for (const pick of document.querySelectorAll('#file-list .stop-pick, #pv-body .stop-pick')) pick.disabled = state.busy;
  if (!$('preview').hidden) syncPreviewInclude();
  // The open name box or create-email step lists the files it will make
  if (!$('name-form').hidden) updateNameHint();
  if (emailDraft && !$('load-hint').classList.contains('error')) updateLoadHint();
}

function setChecked(id, on) {
  const file = state.files.find((f) => f.id === id);
  if (on) state.checked.add(id);
  else state.checked.delete(id);
  if (on && state.multiStop && file) assignStop(file);
  state.built.clear();
  updateSelectionUi();
}

/** Ticking a file by hand (in the list or the preview). Page picks are kept. */
function onRowToggle(file, on) {
  setChecked(file.id, on);
  refreshRow(file);
  if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
}

function refreshRow(file) {
  const row = document.querySelector(`#file-list .file[data-id="${file.id}"]`);
  if (row) {
    const focused = row.contains(document.activeElement) ? document.activeElement : null;
    const fresh = fileRow(file);
    row.replaceWith(fresh);
    if (focused) (fresh.querySelector(focused.matches('.stop-pick') ? '.stop-pick' : 'input') || fresh.querySelector('input')).focus();
  }
  updateSelectionUi();
}

// --------------------------------------------------------------------------- list pictures

/** A page with a coloured label (PDF, DOCX, TIF…), or a picture frame for photos. */
function typeIcon(file) {
  if (file.kind === 'Email') return ENVELOPE_SVG;
  if (file.kind === 'Link') return CLOUD_SVG;
  const ext = extOf(file.name);
  const badge = (label, color) => (label
    ? `<rect x="2" y="14.5" width="19" height="8" rx="1.5" fill="${color}"/><text x="11.5" y="20.6" fill="#fff" font-size="${label.length > 3 ? 5.2 : 6.2}" font-weight="700" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif">${label}</text>`
    : '');
  if (file.kind === 'Image' || file.kind === 'HEIC') {
    const label = file.kind === 'HEIC' ? 'HEIC' : '';
    return `<svg viewBox="0 0 28 28" aria-hidden="true"><rect x="3.5" y="5.5" width="21" height="17" rx="2" fill="#fff" stroke="${ICON_LINE}" stroke-width="1.2"/><circle cx="10" cy="10.5" r="2" fill="${ICON_LINE}"/><path d="M4.5 21 11 14.5l4 4 3-3 5.5 5.5" fill="none" stroke="${ICON_LINE}" stroke-width="1.2" stroke-linejoin="round"/>${badge(label, '#5c2e91')}</svg>`;
  }
  let label = ext.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 4);
  let color = '#616161';
  if (file.kind === 'PDF') { label = 'PDF'; color = '#c4314b'; }
  else if (file.kind === 'TIFF') { label = 'TIF'; color = '#5c2e91'; }
  else if (file.kind === 'Office') color = OFFICE_COLOR[ext] || '#616161';
  return `<svg viewBox="0 0 28 28" aria-hidden="true"><path d="M7.5 2.5h9.2l5.3 5.3v17.7H7.5Z" fill="#fff" stroke="${ICON_LINE}" stroke-width="1.2" stroke-linejoin="round"/><path d="M16.5 2.7v5.3h5.3" fill="none" stroke="${ICON_LINE}" stroke-width="1.2" stroke-linejoin="round"/>${badge(label, color)}</svg>`;
}

function fillThumbSlot(slot, file) {
  slot.textContent = '';
  slot.classList.toggle('pic', !!file.thumb);
  if (file.thumb) {
    const img = document.createElement('img');
    img.alt = '';
    img.src = file.thumb;
    slot.append(img);
  } else {
    slot.innerHTML = typeIcon(file);
  }
}

/** Swaps a row's type icon for the file's picture once it's ready. */
function placeThumb(file) {
  const slot = document.querySelector(`#file-list .file[data-id="${file.id}"] .thumb`);
  if (slot) fillThumbSlot(slot, file);
}

/** Draws a picture into a thumbnail, keeping its shape, and returns it as a data URL. */
function thumbFrom(source, w, h) {
  const scale = THUMB_PX / Math.max(w, h, 1);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/jpeg', 0.85);
}

/** Page 1 of a PDF as a thumbnail. */
async function pdfThumb(file) {
  const doc = await openPdf(file.bytes);
  try {
    file.pages = doc.numPages;
    const page = await doc.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: THUMB_PX / Math.max(base.width, base.height) });
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.floor(viewport.width));
    canvas.height = Math.max(1, Math.floor(viewport.height));
    await page.render({ canvas, viewport }).promise;
    page.cleanup();
    return canvas.toDataURL('image/jpeg', 0.85);
  } finally {
    closePdf(doc);
  }
}

/** The file's thumbnail (page 1 of a PDF or fax, or the photo, turned the right way up), or null. */
async function makeThumb(file) {
  if (!file.bytes) return null;
  if (file.kind === 'PDF') return pdfThumb(file);
  if (file.kind === 'Image') {
    const image = await decodeImage(file);
    try {
      const [w, h] = sizeOf(image);
      return thumbFrom(image, w, h);
    } finally {
      if (typeof image.close === 'function') image.close(); // free the full-size photo
    }
  }
  if (file.kind === 'TIFF') {
    const page = decodeTiffPages(file.bytes, 1)[0];
    if (!page) return null;
    // Fax pages are often 204 x 98 dpi: draw them at their real shape, not their pixel shape
    const stretch = page.xres > 0 && page.yres > 0 ? page.yres / page.xres : 1;
    return thumbFrom(page.canvas, page.w * stretch, page.h);
  }
  return null;
}

/**
 * Makes the list pictures in the background, one file at a time, files on show first. Stops
 * when another thread is opened. Files with no picture keep their type icon.
 */
async function fillThumbs(token) {
  const shown = visibleFiles();
  const order = [...shown, ...state.files.filter((f) => !shown.includes(f))];
  let made = 0;
  for (const file of order) {
    if (token !== state.loadToken) return;
    if (file.thumb !== undefined) continue;
    let thumb = null;
    try {
      thumb = await makeThumb(file);
    } catch (error) {
      thumb = null;
    }
    if (token !== state.loadToken) return;
    file.thumb = thumb;
    if (thumb) made += 1;
    placeThumb(file);
    await new Promise((resolve) => setTimeout(resolve, 0)); // let clicks and typing in between files
  }
  log(`Pictures: ${made} of ${plural(order.length, 'file')}`);
}

// --------------------------------------------------------------------------- preview

async function openPreview(id) {
  const file = state.files.find((f) => f.id === id);
  if (!file) return;
  state.previewIndex = state.files.indexOf(file);
  const token = ++state.previewToken;
  $('preview').hidden = false;
  $('pv-name').textContent = file.name;
  $('pv-name').title = file.name;
  $('pv-meta').textContent = metaLine(file);
  $('pv-pages').textContent = '';
  syncPreviewInclude();
  updatePreviewNav();

  const body = $('pv-body');
  body.textContent = '';
  body.scrollTop = 0;
  const waiting = el('p', 'notice', 'Loading preview…');
  body.append(waiting);
  try {
    if (!file.bytes) throw userError(file.kind === 'Link' ? 'This is a link to a file in the cloud, not a file in the email.' : "This file couldn't be read from the email.");
    if (file.kind === 'PDF') await previewPdf(file, body, token, waiting);
    else if (file.kind === 'Image') await previewImage(file, body, token, waiting);
    else if (file.kind === 'TIFF') await previewTiff(file, body, token, waiting);
    else throw userError(`There's no preview for this kind of file yet. ${WHY_NOT[file.kind] || ''}`.trim());
  } catch (error) {
    if (token !== state.previewToken) return;
    const reason = error.userMessage;
    waiting.textContent = reason
      ? (/^[a-z]/.test(reason) ? `This file ${reason}.` : reason)
      : `Couldn't preview this file (${error.message}).`;
    if (!reason) log(`Preview failed for ${file.name}: ${error.message}`);
  }
}

function closePreview() {
  state.previewToken += 1;
  $('preview').hidden = true;
  $('pv-body').textContent = '';
}

function stepPreview(step) {
  const shown = visibleFiles();
  const at = shown.findIndex((f) => f === state.files[state.previewIndex]);
  const next = shown[at + step];
  if (next) openPreview(next.id);
}

function updatePreviewNav() {
  const shown = visibleFiles();
  const at = shown.findIndex((f) => f === state.files[state.previewIndex]);
  $('pv-prev').disabled = at <= 0;
  $('pv-next').disabled = at < 0 || at >= shown.length - 1;
}

function syncPreviewInclude() {
  const file = state.files[state.previewIndex];
  if (!file) return;
  const box = $('pv-include');
  box.disabled = state.busy || !canCombine(file);
  box.checked = state.checked.has(file.id);
  $('pv-include-label').textContent = canCombine(file) ? 'Include in PDF' : (WHY_NOT[file.kind] || "Can't go into the PDF");
}

async function previewPdf(file, body, token, waiting) {
  const doc = await openPdf(file.bytes);
  try {
    if (token !== state.previewToken) return;
    waiting.remove();
    const total = doc.numPages;
    file.pages = total;
    $('pv-pages').textContent = plural(total, 'page');
    const width = Math.max(200, body.clientWidth - 20);
    const ratio = Math.min(window.devicePixelRatio || 1, 2);
    const count = Math.min(total, PREVIEW_PAGE_LIMIT);
    for (let p = 1; p <= count; p++) {
      if (token !== state.previewToken) return;
      const page = await doc.getPage(p);
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale: (width * ratio) / base.width });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      canvas.setAttribute('aria-label', `Page ${p}`);
      body.append(pageHolder(file, p - 1, canvas));
      await page.render({ canvas, viewport }).promise;
      page.cleanup();
    }
    if (total > count) body.append(el('p', 'notice', `Showing the first ${count} of ${total} pages.`));
    syncPreviewPages();
  } finally {
    closePdf(doc);
  }
}

/** A previewed PDF page with its own "include this page" checkbox (and, with Multi-stop, its stop). */
function pageHolder(file, index, canvas) {
  const holder = el('div', 'pv-page');
  holder.dataset.page = String(index);
  const top = el('div', 'pv-page-top');
  const bar = el('label', 'pv-page-bar');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.setAttribute('aria-label', `Include page ${index + 1}`);
  box.addEventListener('change', () => onPreviewPageToggle(file));
  bar.append(box, el('span', '', `Page ${index + 1}`));
  const pick = document.createElement('select');
  pick.className = 'stop-pick';
  pick.setAttribute('aria-label', `Stop for page ${index + 1}`);
  pick.addEventListener('change', () => setPageStop(file, index, Number(pick.value)));
  top.append(bar, pick);
  holder.append(top, canvas);
  syncPageHolder(holder, file); // right from the start, while later pages still render
  return holder;
}

function pageStopOf(file, index) {
  return (file.pageStops && file.pageStops.get(index)) || file.stop || 1;
}

function syncPageHolder(holder, file) {
  const index = Number(holder.dataset.page);
  const box = holder.querySelector('input');
  box.checked = isPageIncluded(file, index);
  box.disabled = state.busy || !canCombine(file);
  holder.classList.toggle('excluded', !box.checked);
  const pick = holder.querySelector('.stop-pick');
  pick.hidden = !(state.multiStop && box.checked);
  pick.disabled = state.busy;
  if (!pick.hidden) {
    const choices = stopChoices();
    if (pick.options.length !== choices) {
      pick.textContent = '';
      for (let stop = 1; stop <= choices; stop++) pick.append(new Option(`Stop ${stop}`, String(stop)));
    }
    pick.value = String(pageStopOf(file, index));
  }
}

function isPageIncluded(file, index) {
  if (!state.checked.has(file.id)) return false;
  const pages = includedPages(file);
  return pages === null || pages.includes(index);
}

/** Brings the preview's page checkboxes (and stop pickers) in line with the file. */
function syncPreviewPages() {
  const file = state.files[state.previewIndex];
  if (!file) return;
  for (const holder of document.querySelectorAll('#pv-body .pv-page')) syncPageHolder(holder, file);
}

/** Ticking pages in the preview: those exact pages go in (the file goes in if any are ticked). */
function onPreviewPageToggle(file) {
  const holders = [...document.querySelectorAll('#pv-body .pv-page')];
  const shown = new Set(holders.map((h) => Number(h.dataset.page)));
  const pages = new Set();
  for (const holder of holders) {
    if (holder.querySelector('input').checked) pages.add(Number(holder.dataset.page));
  }
  // Pages past the preview limit keep whatever they had.
  const total = file.pages || holders.length;
  for (let i = 0; i < total; i++) {
    if (!shown.has(i) && isPageIncluded(file, i)) pages.add(i);
  }
  if (pages.size === 0) {
    file.pageChoice = null; // nothing ticked: the file is simply left out
    setChecked(file.id, false);
  } else {
    file.pageChoice = pages.size === total ? null : pages; // every page ticked = no special pick
    if (!state.checked.has(file.id)) setChecked(file.id, true);
  }
  normalizePageStops(file);
  state.built.clear();
  refreshRow(file);
  syncPreviewPages();
}

async function previewImage(file, body, token, waiting) {
  const image = await decodeImage(file);
  if (token !== state.previewToken) return;
  const canvas = rasterize(image, MAX_IMAGE_SIDE);
  waiting.remove();
  body.append(canvas);
  const [w, h] = sizeOf(image);
  $('pv-pages').textContent = `${w} × ${h}`;
}

async function previewTiff(file, body, token, waiting) {
  const pages = decodeTiffPages(file.bytes);
  if (token !== state.previewToken) return;
  if (!pages.length) throw userError('This TIFF file has no pages.');
  waiting.remove();
  $('pv-pages').textContent = plural(pages.length, 'page');
  for (const page of pages.slice(0, PREVIEW_PAGE_LIMIT)) {
    page.canvas.style.aspectRatio = `${page.w / page.xresOr1} / ${page.h / page.yresOr1}`;
    body.append(page.canvas);
  }
}

// --------------------------------------------------------------------------- decoding

let pdfjsPromise = null;

function loadPdfJs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import('./pdf.min.js')
      .then((pdfjs) => {
        pdfjs.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.js', import.meta.url).href;
        return pdfjs;
      })
      .catch((error) => {
        pdfjsPromise = null;
        log(`The PDF reader didn't load: ${error.message}`);
        throw userError("The PDF reader didn't load. Close and reopen the panel.");
      });
  }
  return pdfjsPromise;
}

async function openPdf(bytes) {
  const pdfjs = await loadPdfJs();
  // PDF.js takes ownership of the buffer it's given, so hand it a copy.
  const task = pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false });
  try {
    return await task.promise;
  } catch (error) {
    try { task.destroy(); } catch (ignored) { /* already gone */ }
    if (error && error.name === 'PasswordException') throw userError('needs a password to open');
    throw userError("couldn't be read - the file may be damaged");
  }
}

function closePdf(doc) {
  try {
    doc.loadingTask.destroy();
  } catch (error) {
    // nothing left to free
  }
}

async function decodeImage(file) {
  const type = MIME[extOf(file.name)] || file.contentType || 'application/octet-stream';
  const blob = new Blob([file.bytes], { type });
  if (typeof createImageBitmap === 'function') {
    try {
      return await createImageBitmap(blob, { imageOrientation: 'from-image' }); // phone photos: honour rotation
    } catch (error) {
      // fall back to an <img> below
    }
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    return img;
  } catch (error) {
    throw userError("couldn't be opened as an image");
  } finally {
    URL.revokeObjectURL(url);
  }
}

function sizeOf(image) {
  return [image.naturalWidth || image.width, image.naturalHeight || image.height];
}

/** Draws an image on a white canvas, scaled down to maxSide if needed. */
function rasterize(image, maxSide) {
  const [w0, h0] = sizeOf(image);
  const scale = Math.min(1, maxSide / Math.max(w0, h0));
  const w = Math.max(1, Math.round(w0 * scale));
  const h = Math.max(1, Math.round(h0 * scale));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(image, 0, 0, w, h);
  return canvas;
}

function firstNumber(value) {
  const n = Array.isArray(value) ? value[0] : value;
  return typeof n === 'number' && Number.isFinite(n) ? n : 0;
}

/** Every page of a TIFF (fax scans are usually multi-page), or the first `maxPages`, as canvases plus resolution. */
function decodeTiffPages(bytes, maxPages = Infinity) {
  if (typeof UTIF === 'undefined') throw userError("The TIFF reader didn't load.");
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  let ifds;
  try {
    ifds = UTIF.decode(buffer);
  } catch (error) {
    throw userError("couldn't be read - the file may be damaged");
  }
  const pages = [];
  for (const ifd of ifds) {
    if (pages.length >= maxPages) break;
    if (!ifd.t256 || !ifd.t257) continue;
    try {
      UTIF.decodeImage(buffer, ifd);
    } catch (error) {
      continue;
    }
    const w = ifd.width;
    const h = ifd.height;
    if (!w || !h) continue;
    const rgba = UTIF.toRGBA8(ifd);
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    const data = ctx.createImageData(w, h);
    data.data.set(rgba);
    ctx.putImageData(data, 0, 0);
    const perCm = ifd.t296 && ifd.t296[0] === 3;
    const xres = firstNumber(ifd.t282) * (perCm ? 2.54 : 1);
    const yres = firstNumber(ifd.t283) * (perCm ? 2.54 : 1);
    pages.push({
      canvas, w, h, xres, yres,
      xresOr1: xres || 1,
      yresOr1: yres || 1,
      bilevel: !!(ifd.t258 && ifd.t258[0] === 1),
    });
  }
  return pages;
}

// --------------------------------------------------------------------------- building the PDF

function canvasToBlob(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('image encoding failed'))), type, quality);
  });
}

async function canvasToBytes(canvas, type, quality) {
  const blob = await canvasToBlob(canvas, type, quality);
  return new Uint8Array(await blob.arrayBuffer());
}

/** Page size in points for an image with no trustworthy DPI: fit it on a letter page. */
function fitPage(w, h) {
  const [pageW, pageH] = w > h ? [11, 8.5] : [8.5, 11];
  const dpi = Math.max(w / pageW, h / pageH, 72);
  return [(w * 72) / dpi, (h * 72) / dpi];
}

/** Adds a PDF's pages (only `only`, 0-based, when given) to `out`. */
async function addPdf(out, file, only) {
  let source = null;
  try {
    source = await PDFLib.PDFDocument.load(file.bytes, { ignoreEncryption: true, updateMetadata: false });
  } catch (error) {
    source = null; // let PDF.js have a go below; it's more forgiving
  }
  if (source && !source.isEncrypted) {
    let pages = null;
    try {
      const wanted = only ? only.filter((i) => i < source.getPageCount()) : source.getPageIndices();
      pages = await out.copyPages(source, wanted);
    } catch (error) {
      pages = null; // the file's structure is broken; PDF.js below may still read it
    }
    if (pages) {
      pages.forEach((page) => out.addPage(page));
      return 'copied';
    }
  }
  await addPdfAsImages(out, file, only);
  return source && source.isEncrypted ? 'locked PDF, added as page images' : 'repaired, added as page images';
}

/** For locked (owner-password) or slightly broken PDFs: render each page and add it as an image. */
async function addPdfAsImages(out, file, only) {
  const doc = await openPdf(file.bytes);
  const rendered = [];
  try {
    const numbers = only ? only.map((i) => i + 1).filter((p) => p <= doc.numPages) : Array.from({ length: doc.numPages }, (_, i) => i + 1);
    for (const p of numbers) {
      const page = await doc.getPage(p);
      const base = page.getViewport({ scale: 1 });
      const scale = Math.min(LOCKED_PDF_DPI / 72, MAX_IMAGE_SIDE / Math.max(base.width, base.height));
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      await page.render({ canvas, viewport }).promise;
      page.cleanup();
      rendered.push({ jpg: await canvasToBytes(canvas, 'image/jpeg', 0.9), width: base.width, height: base.height });
    }
  } catch (error) {
    throw userError("couldn't be read - the file may be damaged");
  } finally {
    closePdf(doc);
  }
  if (!rendered.length) throw userError('has no pages');
  // Only add pages once every page rendered, so a file is either all in or all out.
  for (const page of rendered) {
    const image = await out.embedJpg(page.jpg);
    out.addPage([page.width, page.height]).drawImage(image, { x: 0, y: 0, width: page.width, height: page.height });
  }
}

async function addImage(out, file) {
  const canvas = rasterize(await decodeImage(file), MAX_IMAGE_SIDE);
  const image = await out.embedJpg(await canvasToBytes(canvas, 'image/jpeg', 0.9));
  const [w, h] = fitPage(canvas.width, canvas.height);
  out.addPage([w, h]).drawImage(image, { x: 0, y: 0, width: w, height: h });
}

async function addTiff(out, file) {
  const pages = decodeTiffPages(file.bytes);
  if (!pages.length) throw userError('has no pages');
  for (const page of pages) {
    let canvas = page.canvas;
    let trustDpi = page.xres >= 50 && page.yres >= 50 && page.w / page.xres <= 17 && page.h / page.yres <= 17;
    let image;
    if (page.bilevel) {
      image = await out.embedPng(await canvasToBytes(canvas, 'image/png')); // crisp black and white
    } else {
      if (Math.max(page.w, page.h) > MAX_IMAGE_SIDE) trustDpi = false;
      canvas = rasterize(canvas, MAX_IMAGE_SIDE);
      image = await out.embedJpg(await canvasToBytes(canvas, 'image/jpeg', 0.9));
    }
    const [w, h] = trustDpi ? [(page.w / page.xres) * 72, (page.h / page.yres) * 72] : fitPage(canvas.width, canvas.height);
    out.addPage([w, h]).drawImage(image, { x: 0, y: 0, width: w, height: h });
  }
}

/**
 * Builds one PDF from [{ file, only }] in that order (only: the 0-based pages of a PDF to use,
 * or null for its picked pages). Cached until the selection changes.
 */
async function buildPdf(entries, label) {
  const key = entries.map(({ file, only }) => `${file.id}:${(only || includedPages(file) || ['all']).join('.')}`).join(',');
  if (state.built.has(key)) return state.built.get(key);
  const out = await PDFLib.PDFDocument.create();
  const skipped = [];
  for (let i = 0; i < entries.length; i++) {
    const { file } = entries[i];
    const only = entries[i].only || includedPages(file);
    setStatus(`${label ? `${label}: ` : ''}Adding ${i + 1} of ${entries.length}: ${file.name}`);
    try {
      if (file.kind === 'PDF') {
        if (only && !only.length) continue;
        const how = await addPdf(out, file, only);
        if (how !== 'copied') log(`${file.name}: ${how}`);
      } else if (file.kind === 'Image') {
        await addImage(out, file);
      } else if (file.kind === 'TIFF') {
        await addTiff(out, file);
      }
    } catch (error) {
      const reason = error.userMessage || `couldn't be converted (${error.message})`;
      skipped.push({ name: file.name, reason });
      log(`Left out ${file.name}: ${reason}`);
    }
  }
  const pages = out.getPageCount();
  let bytes = null;
  if (pages) {
    if (state.topic) out.setTitle(state.topic);
    out.setCreator('Thread Attachments');
    out.setProducer(`Thread Attachments ${VERSION}`);
    bytes = await out.save();
  }
  const result = { bytes, pages, skipped };
  state.built.set(key, result);
  return result;
}

function leftOutText(result) {
  const seen = new Map();
  for (const s of result.skipped) seen.set(`${s.name}\n${s.reason}`, s);
  if (!seen.size) return '';
  return ` Left out: ${[...seen.values()].map((s) => `${s.name} (${s.reason})`).join('; ')}.`;
}

async function withBusy(task) {
  state.busy = true;
  updateSelectionUi();
  try {
    return await task();
  } catch (error) {
    fail(error);
    return null;
  } finally {
    state.busy = false;
    updateSelectionUi();
    if (state.reloadPending) {
      state.reloadPending = false;
      loadThread().catch(fail);
    }
  }
}

async function combineChecked() {
  const files = checkedFiles();
  if (!files.length) return null;
  const result = await buildPdf(files.map((file) => ({ file, only: null })));
  if (!result.pages) {
    setStatus(`None of the checked files could go into a PDF.${leftOutText(result)}`, 'error');
    return null;
  }
  return result;
}

/**
 * The files to email or save, named after `base`: one per stop with Multi-stop on (base-1.pdf,
 * base-2.pdf… by stop number, whatever their size), otherwise one PDF split by `splitMb`.
 * Returns { files, skipped, pages } or null (after saying why) when nothing could be built.
 */
async function makeFiles(base, splitMb) {
  if (!state.multiStop) {
    const result = await combineChecked();
    if (!result) return null;
    return { files: await splitIntoFiles(result.bytes, base, splitMb), skipped: result.skipped, pages: result.pages };
  }
  const files = [];
  const skipped = [];
  for (const { stop, entries } of stopEntries()) {
    const result = await buildPdf(entries, `Stop ${stop}`);
    skipped.push(...result.skipped);
    if (result.pages) files.push({ name: `${base}-${stop}.pdf`, bytes: result.bytes, pages: result.pages, stop, over: false });
  }
  if (!files.length) {
    setStatus(`None of the checked files could go into a PDF.${leftOutText({ skipped })}`, 'error');
    return null;
  }
  const limit = splitMb > 0 ? Math.floor(splitMb * MB) : 0;
  for (const file of files) file.over = !!limit && file.bytes.length > limit;
  log(`Stops: ${files.map((f) => `${f.name} ${mbText(f.bytes.length)} (${plural(f.pages, 'page')})`).join(', ')}`);
  return { files, skipped, pages: files.reduce((sum, f) => sum + f.pages, 0) };
}

// --------------------------------------------------------------------------- save, new email

/**
 * Save PDF. Where Outlook allows Windows' own Save As window, that opens straight away (you can
 * name the file and pick the folder there). Everywhere else, you name the file in the panel first.
 */
async function onSave() {
  if (!checkedFiles().length) return;
  if (typeof window.showSaveFilePicker === 'function' && readSetting('saveDialog') === 'works' && !state.saveSplitMb && !state.multiStop) {
    const picked = await pickSaveLocation(pdfName());
    if (picked.cancelled) return;
    if (picked.handle) {
      await finishSave(picked.handle, pdfName());
      return;
    }
    writeSetting('saveDialog', null); // it stopped working here; name it in the panel instead
  }
  showNameForm();
}

function showNameForm() {
  $('save-actions').hidden = true;
  $('name-form').hidden = false;
  setStatus('');
  updateSelectionUi();
  const input = $('file-name');
  input.value = pdfName().replace(/\.pdf$/i, '');
  $('save-split').value = state.saveSplitMb ? String(state.saveSplitMb) : '';
  updateNameHint();
  input.focus();
  input.select();
}

/** With Multi-stop on, the name box says which files it will save (no size split then). */
function updateNameHint() {
  const hint = $('name-hint');
  $('save-split-row').hidden = state.multiStop;
  hint.hidden = !state.multiStop;
  if (!state.multiStop) return;
  const base = cleanFileName($('file-name').value) || pdfName().replace(/\.pdf$/i, '');
  const stops = [...stopsInUse()].sort((a, b) => a - b);
  hint.textContent = stops.length ? 'One file per stop: ' : 'One file per stop, named -1, -2…';
  if (stops.length) hint.append(namesFragment(stops.map((stop) => `${base}-${stop}.pdf`)));
}

function hideNameForm() {
  $('name-form').hidden = true;
  $('save-actions').hidden = false;
}

function onNameSubmit(event) {
  event.preventDefault();
  if (state.busy || !checkedFiles().length) return;
  state.fileName = cleanFileName($('file-name').value);
  const splitMb = parseSplitMb($('save-split').value);
  if (splitMb !== state.saveSplitMb) {
    state.saveSplitMb = splitMb;
    writeRoaming('saveSplitMb', splitMb || null).catch(() => {});
  }
  hideNameForm();
  const name = pdfName();
  log(`Save: named ${name}${state.multiStop ? ', one file per stop' : splitMb ? `, split over ${splitMb} MB` : ''}`);
  if (state.multiStop || splitMb) {
    saveFiles(name, state.multiStop ? 0 : splitMb).catch(fail);
    return;
  }
  // Straight from the click, so the browser lets Windows' Save As window open if it can.
  pickSaveLocation(name)
    .then((picked) => (picked.cancelled ? null : finishSave(picked.handle || null, name)))
    .catch(fail);
}

/**
 * Save PDF as downloads: one file per stop with Multi-stop on; otherwise one file, or name-1.pdf,
 * name-2.pdf… when it's over the size limit.
 */
async function saveFiles(name, splitMb) {
  const result = await withBusy(() => makeFiles(name.replace(/\.pdf$/i, ''), splitMb));
  if (!result) return;
  const { files } = result;
  try {
    for (let i = 0; i < files.length; i++) {
      downloadBytes(files[i].bytes, files[i].name);
      if (i < files.length - 1) await new Promise((resolve) => setTimeout(resolve, 700));
    }
  } catch (error) {
    log(`Download: failed (${error.name}: ${error.message})`);
    setStatus("Saving isn't possible here. Use Email instead.", 'error');
    return;
  }
  log(`Download: started ${files.map((f) => f.name).join(', ')}`);
  const what = files.length === 1
    ? `Downloading ${files[0].name} (${plural(files[0].pages || result.pages, 'page')}) to your Downloads folder.`
    : `Downloading ${files.length} files to your Downloads folder: ${files.map((f) => f.name).join(', ')}. If Outlook asks about downloading multiple files, choose Allow.`;
  setStatus(`${what}${oversizeNote(files, splitMb)}${leftOutText(result)}`, 'ok');
  if (files.length > 1) showParts(files);
}

function onNameCancel() {
  hideNameForm();
  setStatus('Save cancelled.');
  $('btn-save').focus();
}

/** Opens Windows' Save As window when Outlook allows it. Must be called straight from a click. */
async function pickSaveLocation(name) {
  if (typeof window.showSaveFilePicker !== 'function') {
    log('Save dialog: not available in this Outlook');
    return {};
  }
  try {
    const handle = await window.showSaveFilePicker({ suggestedName: name, types: PDF_TYPES });
    log('Save dialog: opened');
    return { handle };
  } catch (error) {
    if (error && error.name === 'AbortError') {
      log('Save dialog: opened, then cancelled');
      setStatus('Save cancelled.');
      return { cancelled: true };
    }
    log(`Save dialog: blocked (${error && error.name}: ${error && error.message})`);
    return {};
  }
}

/** Builds the PDF and writes it: into the file picked in Save As, or as a download. */
async function finishSave(handle, name) {
  const result = await withBusy(combineChecked);
  if (!result) return;

  if (handle) {
    try {
      const writable = await handle.createWritable();
      await writable.write(result.bytes);
      await writable.close();
      writeSetting('saveDialog', 'works');
      const saved = handle.name || name;
      state.fileName = cleanFileName(saved); // reuse the name chosen in Save As
      log(`Save dialog: saved ${saved} (${humanSize(result.bytes.length)}) ✔`);
      setStatus(`Saved ${saved} (${plural(result.pages, 'page')}).${leftOutText(result)}`, 'ok');
      return;
    } catch (error) {
      writeSetting('saveDialog', null);
      log(`Save dialog: couldn't write the file (${error.name}: ${error.message})`);
    }
  }
  try {
    downloadBytes(result.bytes, name);
    log(`Download: started as ${name}. Check whether it showed up in Downloads.`);
    setStatus(`Downloading ${name} (${plural(result.pages, 'page')}) to your Downloads folder.${leftOutText(result)}`, 'ok');
  } catch (error) {
    log(`Download: failed (${error.name}: ${error.message})`);
    setStatus("Saving isn't possible here. Use Email instead.", 'error');
  }
}

function downloadBytes(bytes, name) {
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.rel = 'noopener';
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

// --------------------------------------------------------------------------- Email ▾ and templates

/** Replaces {load} with the load number. */
function fill(text, load) {
  return String(text || '').replace(/\{load\}/gi, load);
}

function parseAddresses(text) {
  return String(text || '').split(/[;,\s]+/).map((part) => part.trim()).filter(Boolean);
}

function recipientsText(t, load = '') {
  return parseAddresses(fill(t.to, load)).join(', ');
}

/** Every email's subject, whatever the template: "3645493 POD(s)". */
function emailSubject(load) {
  return `${String(load || '').trim()} POD(s)`.slice(0, 255);
}

/** The PDF's name for an email, without .pdf: the template's PDF name, or just the load number. */
function emailFileBase(t, load) {
  const base = fill((t && t.file) || '{load}', load).trim().replace(/\.pdf$/i, '');
  return safeName(base || load || state.topic || 'Combined attachments');
}

/** The files an email will get: base.pdf (maybe split by size), or base-1.pdf… per stop. */
function plannedFiles(base, splitMb) {
  if (!state.multiStop) return { names: [`${base}.pdf`], note: splitMb ? ` (split over ${splitMb} MB)` : '' };
  const stops = [...stopsInUse()].sort((a, b) => a - b);
  return stops.length ? { names: stops.map((stop) => `${base}-${stop}.pdf`), note: '' } : { names: [`${base}-1.pdf`], note: '…' };
}

/** File names for a hint, each kept on one line ("5550123-2.pdf" doesn't break at its dash). */
function namesFragment(names) {
  const fragment = document.createDocumentFragment();
  names.forEach((name, i) => {
    if (i) fragment.append(', ');
    fragment.append(el('span', 'nowrap', name));
  });
  return fragment;
}

/** A likely load number in a subject or file name, e.g. "Load 4471023 Knoxville" -> 4471023. */
function guessLoadNumber(text) {
  const s = String(text || '');
  const tagged = /\b(?:load|ld|order|pro)\b\s*(?:#|no\.?|number)?\s*[:-]?\s*([a-z0-9-]*\d[a-z0-9-]{3,})/i.exec(s);
  if (tagged) return tagged[1];
  const digits = /\b\d{5,}\b/.exec(s);
  return digits ? digits[0] : '';
}

function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** The new email's message: the template's text, then any accessorials. '' when there's neither. */
function bodyHtml(text, acc) {
  const parts = [];
  if (String(text || '').trim()) parts.push(escapeHtml(String(text).slice(0, 30000)).replace(/\r?\n/g, '<br>'));
  const extra = accessorialHtml(acc);
  if (extra) parts.push(extra);
  if (!parts.length) return '';
  return `<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 11pt;">${parts.join('<br><br>')}</div>`;
}

/** The new email's To, CC, subject and message from a template (or a blank email when t is null). */
function emailFields(t, load, acc) {
  const form = { subject: emailSubject(load) };
  if (t) {
    const to = parseAddresses(fill(t.to, load));
    const cc = parseAddresses(fill(t.cc, load));
    if (to.length) form.toRecipients = to.slice(0, 100);
    if (cc.length) form.ccRecipients = cc.slice(0, 100);
  }
  const html = bodyHtml(t ? fill(t.body, load) : '', acc);
  if (html) form.htmlBody = html;
  return form;
}

// ---- accessorials (Detention, Lumper, Other) for the email body

const ACC_FIELDS = {
  det: ['at', 'in', 'appt', 'out', 'hours'],
  lum: ['at', 'total'],
  oth: ['at', 'type', 'total'],
};

function emptyAccessorials() {
  return {
    det: { on: false, at: '', in: '', appt: '', out: '', hours: '' },
    lum: { on: false, at: '', total: '' },
    oth: { on: false, at: '', type: '', total: '' },
  };
}

function readAccessorials() {
  const acc = {};
  for (const [kind, fields] of Object.entries(ACC_FIELDS)) {
    acc[kind] = { on: $(`acc-${kind}`).checked };
    for (const field of fields) acc[kind][field] = $(`${kind}-${field}`).value.trim();
  }
  return acc;
}

function fillAccessorials(acc) {
  for (const [kind, fields] of Object.entries(ACC_FIELDS)) {
    $(`acc-${kind}`).checked = !!acc[kind].on;
    for (const field of fields) $(`${kind}-${field}`).value = acc[kind][field] || '';
  }
  syncAccessorialBoxes();
}

/** Shows the boxes for the ticked accessorials. */
function syncAccessorialBoxes() {
  for (const kind of Object.keys(ACC_FIELDS)) {
    $(`acc-${kind}-box`).hidden = !$(`acc-${kind}`).checked;
    $(`${kind}-at`).classList.toggle('empty', !$(`${kind}-at`).value);
  }
}

/** "350" -> "$350.00", "1,234.5" -> "$1,234.50", "" -> "". Null when it isn't an amount. */
function money(text) {
  const raw = String(text || '').replace(/[$,\s]/g, '');
  if (!raw) return '';
  if (!/^(\d+(\.\d{0,2})?|\.\d{1,2})$/.test(raw)) return null;
  return `$${Number(raw).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

/** [field id, message] for the first thing to fix, or null. */
function accessorialProblem(acc) {
  if (acc.det.on && !acc.det.at) return ['det-at', 'Pick Shipper or Receiver for Detention.'];
  if (acc.lum.on && !acc.lum.at) return ['lum-at', 'Pick Shipper or Receiver for Lumper.'];
  if (acc.lum.on && money(acc.lum.total) === null) return ['lum-total', 'The lumper total should be an amount, like 350.00.'];
  if (acc.oth.on && !acc.oth.at) return ['oth-at', 'Pick Shipper or Receiver for Other.'];
  if (acc.oth.on && !acc.oth.type) return ['oth-type', 'Type what the other accessorial is.'];
  if (acc.oth.on && money(acc.oth.total) === null) return ['oth-total', 'The other total should be an amount, like 75.00.'];
  return null;
}

/** The ticked accessorials as lines for the email body (blank entries are left out). */
function accessorialHtml(acc) {
  if (!acc) return '';
  const blocks = [];
  const add = (title, at, entries) => {
    const lines = [`<b>${escapeHtml(title)}</b> – ${escapeHtml(at)}`];
    for (const [name, value] of entries) if (value) lines.push(`${escapeHtml(name)}: ${escapeHtml(value)}`);
    blocks.push(lines.join('<br>'));
  };
  if (acc.det.on) add('Detention', acc.det.at, [['IN', acc.det.in], ['APPT', acc.det.appt], ['OUT', acc.det.out], ['Hours requested', acc.det.hours]]);
  if (acc.lum.on) add('Lumper', acc.lum.at, [['Total', money(acc.lum.total)]]);
  if (acc.oth.on) add('Other', acc.oth.at, [['Type', acc.oth.type], ['Total', money(acc.oth.total)]]);
  return blocks.join('<br><br>');
}

/** "Detention at Receiver, Lumper at Receiver $350.00" for the status line and the report. */
function accessorialSummary(acc) {
  if (!acc) return '';
  const parts = [];
  if (acc.det.on) parts.push(`Detention at ${acc.det.at}${acc.det.hours ? ` (${acc.det.hours} h)` : ''}`);
  if (acc.lum.on) parts.push(`Lumper at ${acc.lum.at}${money(acc.lum.total) ? ` ${money(acc.lum.total)}` : ''}`);
  if (acc.oth.on) parts.push(`${acc.oth.type} at ${acc.oth.at}${money(acc.oth.total) ? ` ${money(acc.oth.total)}` : ''}`);
  return parts.join(', ');
}

// ---- the menu

function renderEmailMenu() {
  const menu = $('email-menu');
  menu.textContent = '';
  const nothing = checkedFiles().length === 0;
  const ready = !state.busy && !nothing;
  if (nothing) menu.append(el('div', 'menu-note', 'Check the files to send first.'));
  const teamList = teamTemplates();
  const mine = state.myTemplates;
  if (teamList.length) {
    menu.append(el('div', 'menu-group', 'Team'));
    for (const t of teamList) menu.append(menuItem(t.name, recipientsText(t) || 'No recipients set', () => startEmail(t), !ready));
  }
  if (mine.length) {
    menu.append(el('div', 'menu-group', 'Mine'));
    for (const t of mine) menu.append(menuItem(t.name, recipientsText(t) || 'No recipients set', () => startEmail(t), !ready));
  }
  if (teamList.length || mine.length) menu.append(el('div', 'menu-sep'));
  menu.append(menuItem('Blank email', 'Just the PDF, no recipients', () => startEmail(null), !ready));
  menu.append(menuItem(teamList.length || mine.length ? 'Manage templates…' : 'Create a template…', '', openTemplates, false));
}

function menuItem(title, sub, action, disabled) {
  const button = el('button', 'menu-item');
  button.type = 'button';
  button.setAttribute('role', 'menuitem');
  button.disabled = !!disabled;
  button.append(el('span', 'mi-name', title));
  if (sub) button.append(el('span', 'mi-sub', sub));
  button.addEventListener('click', () => {
    closeEmailMenu(false);
    action();
  });
  return button;
}

function openEmailMenu() {
  renderEmailMenu();
  $('email-menu').hidden = false;
  $('btn-email').setAttribute('aria-expanded', 'true');
  const first = $('email-menu').querySelector('.menu-item:not(:disabled)');
  if (first) first.focus();
}

function closeEmailMenu(returnFocus) {
  if ($('email-menu').hidden) return;
  $('email-menu').hidden = true;
  $('btn-email').setAttribute('aria-expanded', 'false');
  if (returnFocus) $('btn-email').focus();
}

function onMenuKeys(event) {
  if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
  event.preventDefault();
  const items = [...$('email-menu').querySelectorAll('.menu-item:not(:disabled)')];
  if (!items.length) return;
  const at = items.indexOf(document.activeElement);
  const step = event.key === 'ArrowDown' ? 1 : -1;
  items[(at + step + items.length) % items.length].focus();
}

// ---- the create-email step: load number (every email's subject has it) and accessorials

let emailDraft = null; // { template } while the step is open; template is null for a blank email

function startEmail(t) {
  emailDraft = { template: t };
  hideNameForm();
  $('save-actions').hidden = true;
  $('load-form').hidden = false;
  setStatus('');
  $('load-number').value = state.loadNumber || guessLoadNumber(state.fileName) || guessLoadNumber(state.topic) || '';
  fillAccessorials(state.accessorials);
  $('acc-error').textContent = '';
  updateLoadHint();
  updateSelectionUi();
  $('load-number').focus();
  $('load-number').select();
}

function hideLoadForm() {
  if (!$('load-form').hidden) state.accessorials = readAccessorials(); // kept for this thread
  $('load-form').hidden = true;
  if ($('name-form').hidden) $('save-actions').hidden = false;
  emailDraft = null;
}

function updateLoadHint() {
  if (!emailDraft) return;
  const t = emailDraft.template;
  const hint = $('load-hint');
  hint.classList.remove('error');
  const load = $('load-number').value.trim();
  const { names, note } = plannedFiles(emailFileBase(t, load || '…'), t ? t.splitMb : 0);
  const to = t ? recipientsText(t, load) : '';
  hint.textContent = '';
  hint.append(`${t ? t.name : 'Blank email'}: attaches `, namesFragment(names), `${note}${to ? ` · to ${to}` : ''}`);
}

function onLoadSubmit(event) {
  event.preventDefault();
  if (!emailDraft || state.busy || !checkedFiles().length) return;
  const t = emailDraft.template;
  const load = $('load-number').value.trim();
  if (!load) {
    const hint = $('load-hint');
    hint.textContent = 'Type the load number.';
    hint.classList.add('error');
    $('load-number').focus();
    return;
  }
  const acc = readAccessorials();
  const problem = accessorialProblem(acc);
  if (problem) {
    $('acc-error').textContent = problem[1];
    $(problem[0]).focus();
    return;
  }
  state.loadNumber = load;
  hideLoadForm();
  createEmail(t, load, acc).catch(fail);
}

function onAccessorialPick(kind, on) {
  syncAccessorialBoxes();
  $('acc-error').textContent = '';
  if (on) $(`${kind}-at`).focus();
}

/** Tidies a typed amount when you leave the box: 350 -> 350.00 */
function onMoneyChange(event) {
  const value = money(event.target.value);
  if (value) event.target.value = value.slice(1);
}

function onLoadCancel() {
  hideLoadForm();
  setStatus('');
  $('btn-email').focus();
}

// ---- creating the email

/**
 * Builds the PDF (one per stop with Multi-stop on, otherwise split by the template's size limit),
 * then opens the new email. With the team's Azure function set up, the files are uploaded so
 * Outlook attaches them itself; otherwise (or if the upload fails) they wait for the new email's
 * Attach PDF button.
 */
async function createEmail(template, load, acc) {
  const base = emailFileBase(template, load);
  const splitMb = template ? template.splitMb : 0;
  const result = await withBusy(() => makeFiles(base, splitMb));
  if (!result) return;
  const { files } = result;
  const names = files.map((f) => f.name).join(', ');
  const form = emailFields(template, load, acc);
  const extras = accessorialSummary(acc);

  let links = null;
  let azureProblem = '';
  const azure = teamAzure();
  if (azure) {
    links = await withBusy(async () => {
      try {
        const list = [];
        for (let i = 0; i < files.length; i++) {
          setStatus(`Uploading ${files.length > 1 ? `${i + 1} of ${files.length}: ` : ''}${files[i].name}…`);
          list.push({ type: 'file', name: files[i].name, url: await uploadToAzure(azure, files[i].bytes, files[i].name) });
        }
        log(`Azure: uploaded ${names}`);
        return list;
      } catch (error) {
        azureProblem = error.userMessage || error.message;
        log(`Azure: upload failed (${azureProblem})`);
        return null;
      }
    });
  }

  try {
    await idbSet(PENDING_KEY, {
      files: files.map((f) => ({ name: f.name, bytes: f.bytes, pages: f.pages })),
      created: Date.now(),
      subject: form.subject || '',
      viaLink: !!links,
    });
  } catch (error) {
    log(`New email: couldn't store the PDF (${error.name}: ${error.message})`);
    if (!links) {
      setStatus("Couldn't hand the PDF over to a new email here.", 'error');
      return;
    }
  }
  if (links) form.attachments = links;

  const mailbox = Office.context.mailbox;
  try {
    if (typeof mailbox.displayNewMessageFormAsync === 'function') {
      await call((done) => mailbox.displayNewMessageFormAsync(form, done));
    } else {
      mailbox.displayNewMessageForm(form);
    }
  } catch (error) {
    log(`New email: couldn't open one (${error.message})`);
    setStatus(`Couldn't open a new email: ${error.message}`, 'error');
    return;
  }
  const to = (form.toRecipients || []).join(', ');
  log(`New email: opened (${template ? `template "${template.name}"` : 'blank'}; ${names}${to ? `; to ${to}` : ''}${extras ? `; ${extras}` : ''}${links ? '; attached through Azure' : ''})`);
  const accNote = extras ? ` In the message: ${extras}.` : '';
  const extra = `${accNote}${oversizeNote(files, splitMb)}${leftOutText(result)}`;
  if (links) {
    setStatus(`New email opened${to ? ` to ${to}` : ''} with ${names} attached.${extra}`, 'ok');
  } else {
    const why = azureProblem ? ` (Couldn't attach it automatically: ${azureProblem}.)` : '';
    setStatus(`New email opened${to ? ` to ${to}` : ''}. Click Attach PDF on its ribbon to add ${names}.${why}${extra}`, azureProblem ? 'error' : 'ok');
  }
}

// ---- splitting by size

function parseSplitMb(value) {
  const n = parseFloat(String(value === undefined || value === null ? '' : value).replace(',', '.'));
  return Number.isFinite(n) && n > 0 ? Math.round(Math.min(Math.max(n, 0.1), 150) * 10) / 10 : 0;
}

function mbText(bytes) {
  return `${(bytes / MB).toFixed(1)} MB`;
}

/** Copies pages [from, to) of a loaded PDF into a new PDF and returns its bytes. */
async function pagesToPdf(src, from, to) {
  const out = await PDFLib.PDFDocument.create();
  const pages = await out.copyPages(src, Array.from({ length: to - from }, (_, i) => from + i));
  pages.forEach((page) => out.addPage(page));
  if (state.topic) out.setTitle(state.topic);
  out.setCreator('Thread Attachments');
  out.setProducer(`Thread Attachments ${VERSION}`);
  return out.save();
}

/**
 * Splits a PDF into as few parts as possible, each at most `limit` bytes, keeping page order.
 * Grows each part by doubling, then narrows it down, so it only saves a handful of times.
 * A single page bigger than the limit becomes a part of its own (marked `over`).
 */
async function splitPdfBySize(bytes, limit) {
  const src = await PDFLib.PDFDocument.load(bytes, { updateMetadata: false });
  const total = src.getPageCount();
  const parts = [];
  let from = 0;
  while (from < total) {
    setStatus(`Splitting the PDF: part ${parts.length + 1}…`);
    let good = from + 1;
    let goodBytes = await pagesToPdf(src, from, good);
    if (goodBytes.length > limit) {
      parts.push({ bytes: goodBytes, pages: 1, over: true });
      from = good;
      continue;
    }
    let bad = null;
    let step = 1;
    while (good < total) {
      const next = Math.min(total, good + step);
      const candidate = await pagesToPdf(src, from, next);
      if (candidate.length > limit) {
        bad = next;
        break;
      }
      good = next;
      goodBytes = candidate;
      step *= 2;
    }
    while (bad !== null && bad - good > 1) {
      const mid = Math.floor((good + bad) / 2);
      const candidate = await pagesToPdf(src, from, mid);
      if (candidate.length > limit) bad = mid;
      else {
        good = mid;
        goodBytes = candidate;
      }
    }
    parts.push({ bytes: goodBytes, pages: good - from, over: false });
    from = good;
  }
  return parts;
}

/** The files to send or save: base.pdf, or base-1.pdf, base-2.pdf… when it's over the limit. */
async function splitIntoFiles(bytes, base, splitMb) {
  const limit = splitMb > 0 ? Math.floor(splitMb * MB) : 0;
  if (!limit || bytes.length <= limit) return [{ name: `${base}.pdf`, bytes, pages: null, over: false }];
  const parts = await splitPdfBySize(bytes, limit);
  if (parts.length === 1) return [{ name: `${base}.pdf`, bytes: parts[0].bytes, pages: parts[0].pages, over: parts[0].over }];
  const files = parts.map((part, i) => ({ name: `${base}-${i + 1}.pdf`, bytes: part.bytes, pages: part.pages, over: part.over }));
  log(`Split over ${splitMb} MB: ${files.map((f) => `${f.name} ${mbText(f.bytes.length)} (${plural(f.pages, 'page')})`).join(', ')}`);
  return files;
}

function oversizeNote(files, splitMb) {
  const over = files.filter((f) => f.over);
  if (!over.length) return '';
  const why = over[0].stop ? ' (with Multi-stop, each stop stays one file)' : ' because a single page is that big';
  return ` ${over.map((f) => `${f.name} is ${mbText(f.bytes.length)}`).join('; ')}: over the ${splitMb} MB limit${why}.`;
}

// ---- Azure (the team admin's function app)

function randomB64url(bytes) {
  return bytesToBase64(crypto.getRandomValues(new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** https only (plain http just for testing on this computer). */
function okAzureUrl(url) {
  return /^https:\/\/[^\s/]+\/\S+$/i.test(url) || /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\/\S+$/i.test(url);
}

/**
 * The function address as typed or pasted: the Function App's domain alone (as the portal's
 * Overview shows it), .../api, or the function's own URL all become
 * https://<app>.azurewebsites.net/api/PdfEmails.
 */
function azureBaseUrl(text) {
  let url = String(text || '').trim().replace(/[?#].*$/, '').replace(/\/+$/, '');
  if (url && !/^[a-z][a-z0-9+.-]*:\/\//i.test(url) && /^[^\s/:]+\.[^\s/:]+(:\d+)?(\/\S*)?$/.test(url)) url = `https://${url}`;
  if (/^https?:\/\/[^/\s]+$/i.test(url)) url += `/api/${AZURE_FUNCTION}`;
  else if (/^https?:\/\/[^/\s]+\/api$/i.test(url)) url += `/${AZURE_FUNCTION}`;
  return url;
}

function cleanAzure(value) {
  if (!value || typeof value !== 'object') return null;
  const url = String(value.url || '').trim().replace(/\/+$/, '');
  const key = String(value.key || '').trim();
  return okAzureUrl(url) && key.length >= 16 ? { url, key } : null;
}

function teamAzure() {
  const source = (isAdmin() && team.draft) || team.published;
  return source ? source.azure || null : null;
}

/** Turns a failed call to the PDF Emails function into a plain-language reason. */
async function azureProblem(response) {
  let detail = '';
  try {
    const body = await response.json();
    detail = body && body.error ? String(body.error).replace(/\.$/, '') : '';
  } catch (error) {
    detail = '';
  }
  if (response.status === 401) return userError("the upload key doesn't match the Function App's THREAD_ATTACHMENTS_UPLOAD_KEY setting");
  if (response.status === 503) return userError("THREAD_ATTACHMENTS_UPLOAD_KEY isn't set in the Function App yet");
  if (response.status === 404) return userError(`there's no ${AZURE_FUNCTION} function at that address`);
  return userError(`Azure answered ${response.status}${detail ? `: ${detail}` : ''}`);
}

function unreachable() {
  return userError(`couldn't reach the Azure function (check its address, and that its CORS settings list ${window.location.origin})`);
}

/** Uploads one PDF to the team's Azure function; returns the link Outlook downloads it from. */
async function uploadToAzure(azure, bytes, name) {
  let response;
  try {
    response = await fetch(`${azure.url}?name=${encodeURIComponent(name)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream', 'x-upload-key': azure.key },
      body: bytes,
    });
  } catch (error) {
    throw unreachable();
  }
  if (!response.ok) throw await azureProblem(response);
  let link = '';
  try {
    link = String((await response.json()).url || '');
  } catch (error) {
    link = '';
  }
  if (!okAzureUrl(link)) throw userError("the Azure function didn't send back a download link");
  return link;
}

function downloadText(text, name, type) {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.rel = 'noopener';
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

// --------------------------------------------------------------------------- template storage

function newId() {
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}

function cleanTemplates(list) {
  if (!Array.isArray(list)) return [];
  return list
    .filter((t) => t && typeof t === 'object' && String(t.name || '').trim())
    .map((t) => ({
      id: String(t.id || newId()),
      name: String(t.name).trim().slice(0, 60),
      to: String(t.to || ''),
      cc: String(t.cc || ''),
      body: String(t.body || ''),
      file: String(t.file || ''),
      splitMb: parseSplitMb(t.splitMb),
    }));
}

/** Outlook's roaming settings follow the person's mailbox (classic, new Outlook and the web). */
function roamingSettings() {
  try {
    const rs = Office.context.roamingSettings;
    return rs && typeof rs.get === 'function' ? rs : null;
  } catch (error) {
    return null;
  }
}

function readRoaming(name) {
  const rs = roamingSettings();
  try {
    if (rs) return rs.get(name) ?? null;
    const raw = localStorage.getItem(`thread-attachments.${name}`);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    return null;
  }
}

async function writeRoaming(name, value) {
  const rs = roamingSettings();
  if (rs && typeof rs.saveAsync === 'function') {
    if (value === null) rs.remove(name);
    else rs.set(name, value);
    await call((done) => rs.saveAsync(done));
    return;
  }
  if (value === null) localStorage.removeItem(`thread-attachments.${name}`);
  else localStorage.setItem(`thread-attachments.${name}`, JSON.stringify(value));
}

// --------------------------------------------------------------------------- team: install-file key + encrypted team.json

function b64urlToBytes(text) {
  let s = String(text).replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  return base64ToBytes(s);
}

function teamCryptoKey(key) {
  return crypto.subtle.importKey('raw', b64urlToBytes(key), 'AES-GCM', false, ['encrypt', 'decrypt']);
}

async function encryptTeam(data, key) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const plain = new TextEncoder().encode(JSON.stringify(data));
  const sealed = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await teamCryptoKey(key), plain));
  return JSON.stringify({ format: 'thread-attachments-team', v: 1, iv: bytesToBase64(iv), data: bytesToBase64(sealed) });
}

async function decryptTeam(text, key) {
  const box = JSON.parse(text);
  if (!box || box.format !== 'thread-attachments-team') throw new Error('not a team file');
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: base64ToBytes(box.iv) }, await teamCryptoKey(key), base64ToBytes(box.data));
  return JSON.parse(new TextDecoder().decode(plain));
}

function cleanTeam(data) {
  return {
    v: 1,
    updated: Number(data && data.updated) || 0,
    by: String((data && data.by) || ''),
    admins: Array.isArray(data && data.admins) ? data.admins.map((a) => String(a).toLowerCase().trim()).filter(Boolean) : [],
    templates: cleanTemplates(data && data.templates),
    azure: cleanAzure(data && data.azure),
  };
}

async function loadTeam() {
  try {
    const profile = Office.context.mailbox.userProfile;
    team.me = String((profile && profile.emailAddress) || '').toLowerCase().trim();
  } catch (error) {
    team.me = '';
  }
  if (!team.key) {
    team.status = 'nokey';
  } else {
    let response = null;
    try {
      response = await fetch(`${TEAM_FILE}?t=${Date.now()}`, { cache: 'no-store' });
    } catch (error) {
      team.status = 'error';
      team.error = error.message;
    }
    if (response && response.status === 404) {
      team.status = 'nofile';
    } else if (response && !response.ok) {
      team.status = 'error';
      team.error = `the site answered ${response.status}`;
    } else if (response) {
      try {
        team.published = cleanTeam(await decryptTeam(await response.text(), team.key));
        team.status = 'ok';
      } catch (error) {
        team.status = 'wrongkey';
      }
    }
  }
  const draft = readRoaming(DRAFT_KEY);
  team.draft = draft && Array.isArray(draft.templates) ? cleanTeam(draft) : null;
  if (team.draft && team.published && team.published.updated >= team.draft.updated) {
    team.draft = null; // already published
    writeRoaming(DRAFT_KEY, null).catch(() => {});
  }
  const what = {
    nokey: 'no team key in this install',
    nofile: 'no team templates published yet',
    ok: `${plural(team.published ? team.published.templates.length : 0, 'team template')}`,
    wrongkey: "this install's key doesn't open team.json",
    error: `couldn't load team.json (${team.error})`,
  }[team.status];
  log(`Team: ${what}${isAdmin() ? ' · you manage them' : ''}${team.draft && isAdmin() ? ' · unpublished changes' : ''}`);
}

/** The first person to set up a team manages it; after that, whoever team.json lists. */
function isAdmin() {
  if (team.status === 'nofile') return true;
  return team.status === 'ok' && !!team.me && team.published.admins.includes(team.me);
}

function teamTemplates() {
  if (isAdmin() && team.draft) return team.draft.templates;
  return team.published ? team.published.templates : [];
}

function showLocked() {
  const site = new URL('.', window.location.href).href;
  const [title, body] = {
    nokey: ["This copy of Thread Attachments isn't set up for a team.",
      'Install it from the team install file you were sent. If you manage the add-in, make a team install file on its page:'],
    wrongkey: ["This install file doesn't match your team.",
      'Ask whoever sent it to you for the current install file.'],
    error: ["Couldn't load your team's settings.",
      `Check your internet connection, then close and reopen this panel. (${team.error})`],
  }[team.status];
  const card = el('div', 'lock-card');
  card.append(el('p', '', title), el('p', 'muted', body));
  if (team.status === 'nokey') card.append(el('code', '', site));
  const view = $('message-view');
  view.textContent = '';
  view.append(card);
  setSummary('');
  showView('message-view');
}

// --------------------------------------------------------------------------- managing templates

let editing = null; // { scope: 'team' | 'mine', id }
let deleteArmed = 0;

function openTemplates() {
  closeEmailMenu(false);
  editing = null;
  $('templates').hidden = false;
  renderTemplateSheet();
  $('tpl-back').focus();
}

function closeTemplates() {
  $('templates').hidden = true;
  editing = null;
  $('btn-email').focus();
}

function cancelEdit() {
  editing = null;
  renderTemplateSheet();
}

function renderTemplateSheet() {
  const formOpen = !!editing;
  $('tpl-list-view').hidden = formOpen;
  $('tpl-form').hidden = !formOpen;
  if (formOpen) return;
  $('tpl-title').textContent = 'Email templates';
  const admin = isAdmin();
  const teamList = teamTemplates();
  $('team-note').textContent = admin ? 'You manage these' : 'Set by your team admin';
  fillTemplateList($('team-list'), teamList, admin ? 'team' : null);
  $('team-empty').hidden = teamList.length > 0;
  $('team-empty').textContent = admin
    ? 'Add the templates your reps should use, then publish them.'
    : "Your team admin hasn't added any templates yet.";
  $('team-admin').hidden = !admin;
  if (admin) renderPublishBox();
  $('azure-section').hidden = !admin;
  if (admin) renderAzureBox();
  fillTemplateList($('mine-list'), state.myTemplates, 'mine');
  $('mine-empty').hidden = state.myTemplates.length > 0;
}

function fillTemplateList(list, templates, scope) {
  list.textContent = '';
  for (const t of templates) {
    const row = el('li', 'tpl-row');
    const text = el('div', 'tpl-text');
    text.append(el('div', 'tpl-name', t.name), el('div', 'tpl-sub', recipientsText(t) || 'No recipients set'));
    row.append(text);
    if (scope) {
      const edit = el('button', 'btn', 'Edit');
      edit.type = 'button';
      edit.setAttribute('aria-label', `Edit ${t.name}`);
      edit.addEventListener('click', () => editTemplate(scope, t.id));
      row.append(edit);
    }
    list.append(row);
  }
}

function renderPublishBox(message) {
  const pending = !!team.draft;
  const box = $('team-publish').closest('.publish-box');
  box.classList.toggle('pending', pending && !team.justPublished && !message);
  $('team-publish').disabled = !pending;
  const text = $('publish-text');
  if (message) text.textContent = message;
  else if (team.justPublished) text.textContent = 'team.json downloaded. Upload it to your GitHub repository (Add file → Upload files → Commit changes). If it saved as "team (1).json", rename it to team.json first. Reps get the changes the next time they open the panel.';
  else if (pending) text.textContent = "You have changes your team can't see yet. Publish, then upload team.json to GitHub.";
  else if (team.status === 'nofile') text.textContent = 'Add templates, then publish them to your team.';
  else text.textContent = 'Your team has the latest templates.';
}

function editTemplate(scope, id) {
  const list = scope === 'team' ? teamTemplates() : state.myTemplates;
  const t = id ? list.find((x) => x.id === id) : null;
  editing = { scope, id: t ? t.id : null };
  deleteArmed = 0;
  $('tpl-name').value = t ? t.name : '';
  $('tpl-to').value = t ? t.to : '';
  $('tpl-cc').value = t ? t.cc : '';
  $('tpl-message').value = t ? t.body : NEW_TEMPLATE.body;
  $('tpl-file').value = t ? (t.file || '{load}') : NEW_TEMPLATE.file;
  $('tpl-split').value = t && t.splitMb ? String(t.splitMb) : '';
  $('tpl-delete').hidden = !t;
  $('tpl-delete').textContent = 'Delete this template';
  $('tpl-error').textContent = '';
  renderTemplateSheet();
  $('tpl-title').textContent = `${t ? 'Edit' : 'New'} ${scope === 'team' ? 'team template' : 'template'}`;
  $('tpl-name').focus();
}

function templateProblem(t) {
  if (!t.name) return 'Give the template a name.';
  for (const [label, text] of [['To', t.to], ['CC', t.cc]]) {
    for (const address of parseAddresses(text)) {
      if (/\{load\}/i.test(address)) continue;
      if (!/^[^\s@<>()]+@[^\s@<>()]+\.[^\s@<>()]+$/.test(address)) return `${label}: "${address}" doesn't look like an email address.`;
    }
  }
  return '';
}

function upsert(list, item) {
  const at = list.findIndex((x) => x.id === item.id);
  if (at >= 0) list[at] = item;
  else list.push(item);
}

/** Saves the admin's unpublished team changes: { templates } and/or { azure }. */
async function saveTeamDraft(changes) {
  const base = team.draft || team.published;
  const admins = base && base.admins.length ? base.admins : [team.me].filter(Boolean);
  const templates = changes.templates || (base ? base.templates : []);
  const azure = 'azure' in changes ? changes.azure : (base ? base.azure || null : null);
  const draft = { v: 1, updated: Date.now(), by: team.me, admins, templates, azure };
  await writeRoaming(DRAFT_KEY, draft);
  team.draft = draft;
  team.justPublished = false;
}

async function onTemplateSubmit(event) {
  event.preventDefault();
  if (!editing) return;
  const data = {
    id: editing.id || newId(),
    name: $('tpl-name').value.trim(),
    to: $('tpl-to').value.trim(),
    cc: $('tpl-cc').value.trim(),
    body: $('tpl-message').value.replace(/\s+$/, ''),
    file: $('tpl-file').value.trim().replace(/\.pdf$/i, '') || '{load}',
    splitMb: parseSplitMb($('tpl-split').value),
  };
  if ($('tpl-split').value.trim() && !data.splitMb) {
    $('tpl-error').textContent = 'Split size should be a number of MB, like 10.';
    return;
  }
  const problem = templateProblem(data);
  if (problem) {
    $('tpl-error').textContent = problem;
    return;
  }
  try {
    if (editing.scope === 'mine') {
      const list = state.myTemplates.slice();
      upsert(list, data);
      await writeRoaming(MY_TEMPLATES_KEY, list);
      state.myTemplates = list;
    } else {
      const list = teamTemplates().slice();
      upsert(list, data);
      await saveTeamDraft({ templates: list });
    }
  } catch (error) {
    $('tpl-error').textContent = `Couldn't save it: ${error.message}`;
    return;
  }
  log(`Templates: saved ${editing.scope === 'team' ? 'team' : 'my'} template "${data.name}"`);
  editing = null;
  renderTemplateSheet();
}

async function onTemplateDelete() {
  if (!editing || !editing.id) return;
  if (Date.now() - deleteArmed > 4000) { // ask twice, since there's no undo
    deleteArmed = Date.now();
    $('tpl-delete').textContent = 'Click again to delete';
    return;
  }
  const { scope, id } = editing;
  try {
    if (scope === 'mine') {
      const list = state.myTemplates.filter((t) => t.id !== id);
      await writeRoaming(MY_TEMPLATES_KEY, list);
      state.myTemplates = list;
    } else {
      await saveTeamDraft({ templates: teamTemplates().filter((t) => t.id !== id) });
    }
  } catch (error) {
    $('tpl-error').textContent = `Couldn't delete it: ${error.message}`;
    return;
  }
  log(`Templates: deleted a ${scope === 'team' ? 'team' : 'personal'} template`);
  editing = null;
  renderTemplateSheet();
}

/** Encrypts the team templates with the team key and downloads team.json to upload to GitHub. */
async function onPublish() {
  const source = team.draft || team.published;
  if (!source || !team.key) return;
  if (!team.me) {
    renderPublishBox("Couldn't read your email address from Outlook, so this can't be published.");
    return;
  }
  const admins = source.admins.length ? source.admins : [team.me];
  const payload = { v: 1, updated: source.updated || Date.now(), by: team.me, admins, templates: source.templates, azure: source.azure || null };
  downloadText(await encryptTeam(payload, team.key), TEAM_FILE, 'application/json');
  team.justPublished = true;
  log(`Team: published ${plural(payload.templates.length, 'template')} (team.json downloaded)`);
  renderPublishBox();
}

// ---- Azure settings (team admin)

function renderAzureBox(message, tone) {
  const azure = teamAzure();
  $('az-url').value = azure ? azure.url : '';
  $('az-key').value = azure ? azure.key : '';
  $('azure-note').textContent = azure ? 'On' : 'Off';
  setAzureStatus(message || (azure
    ? 'Create email attaches PDFs through your Azure function.'
    : 'Not set up. Create email uses the Attach PDF button instead.'), tone);
}

function setAzureStatus(text, tone) {
  const node = $('az-status');
  node.textContent = text;
  node.className = `hint${tone ? ` ${tone}` : ''}`;
}

function azureFromFields() {
  const typed = $('az-url').value.trim();
  const key = $('az-key').value.trim();
  if (!typed && !key) return { off: true };
  const url = azureBaseUrl(typed);
  if (!okAzureUrl(url)) {
    return { problem: "The function address should be your PDF Emails Function App's Default domain, like yourapp.azurewebsites.net." };
  }
  if (key.length < 16) return { problem: 'The upload key should be at least 16 characters. Click Generate to make one.' };
  return { azure: { url, key } };
}

function onAzureGenerate() {
  $('az-key').value = randomB64url(24);
  $('az-key').select();
  setAzureStatus("New key made. Click Copy, paste it into your Function App as the THREAD_ATTACHMENTS_UPLOAD_KEY setting, then come back and click Test connection.");
}

async function onAzureCopy() {
  const key = $('az-key').value.trim();
  if (!key) return;
  try {
    await navigator.clipboard.writeText(key);
    setAzureStatus('Key copied.');
  } catch (error) {
    $('az-key').select();
    let copied = false;
    try {
      copied = document.execCommand('copy');
    } catch (ignored) {
      copied = false;
    }
    setAzureStatus(copied ? 'Key copied.' : 'Press Ctrl+C to copy the key.');
  }
}

async function onAzureSave() {
  const { azure, problem, off } = azureFromFields();
  if (problem) {
    setAzureStatus(problem, 'error');
    return;
  }
  try {
    await saveTeamDraft({ azure: off ? null : azure });
  } catch (error) {
    setAzureStatus(`Couldn't save it: ${error.message}`, 'error');
    return;
  }
  log(`Azure: settings ${off ? 'cleared' : 'saved'} (not published yet)`);
  renderPublishBox();
  renderAzureBox(off
    ? 'Turned off. Publish to update your team.'
    : 'Saved. Click Publish to team above to share it with your reps.', 'ok');
}

/** Checks the function is there and set up, then uploads a one-page PDF and reads it back. */
async function onAzureTest() {
  const { azure, problem } = azureFromFields();
  if (problem || !azure) {
    setAzureStatus(problem || 'Fill in the function address and upload key first.', 'error');
    return;
  }
  $('az-test').disabled = true;
  setAzureStatus('Testing… (the first call can take a few seconds while Azure starts the function)');
  try {
    let health;
    try {
      health = await fetch(azure.url, { cache: 'no-store' });
    } catch (error) {
      throw unreachable();
    }
    if (!health.ok) throw await azureProblem(health);
    let info = {};
    try {
      info = await health.json();
    } catch (error) {
      info = {};
    }
    if (info.service !== 'thread-attachments-pdf-emails') throw userError(`that address answers, but it isn't the ${AZURE_FUNCTION} function`);
    if (!info.uploadKeySet) throw userError("THREAD_ATTACHMENTS_UPLOAD_KEY isn't set in the Function App yet");

    const doc = await PDFLib.PDFDocument.create();
    doc.addPage([200, 200]).drawText('Thread Attachments test', { x: 20, y: 100, size: 12 });
    const bytes = await doc.save();
    const link = await uploadToAzure(azure, bytes, 'test.pdf');
    let back;
    try {
      back = await fetch(link, { cache: 'no-store' });
    } catch (error) {
      throw userError(`the upload worked, but the file couldn't be read back (check that CORS lists ${window.location.origin})`);
    }
    if (!back.ok) throw userError(`the upload worked, but reading it back answered ${back.status}`);
    const got = new Uint8Array(await back.arrayBuffer());
    if (got.length !== bytes.length) throw userError('the file came back different from what was uploaded');
    log(`Azure: test connection worked (${info.version || 'version unknown'})`);
    setAzureStatus('Connection works. Click Save, then Publish to share it with your team.', 'ok');
  } catch (error) {
    const reason = error.userMessage || error.message;
    log(`Azure: test failed (${reason})`);
    setAzureStatus(`Didn't work: ${reason}.`, 'error');
  } finally {
    $('az-test').disabled = false;
  }
}

/** The install file for reps: the add-in's manifest with this team's key in it. */
async function onShareInstall() {
  let xml;
  try {
    const response = await fetch('manifest.template.xml', { cache: 'no-store' });
    if (!response.ok) throw new Error(`the site answered ${response.status}`);
    const base = new URL('.', window.location.href).href;
    xml = (await response.text()).split('{{BASE}}').join(base).split('{{QUERY}}').join(`?team=${team.key}`);
  } catch (error) {
    renderPublishBox(`Couldn't make the install file: ${error.message}`);
    return;
  }
  downloadText(xml, 'thread-attachments-team.xml', 'application/xml');
  log('Team: install file downloaded');
  renderPublishBox('thread-attachments-team.xml downloaded. Email it to a rep; they install it at aka.ms/olksideload (My add-ins → Add a custom add-in → Add from File). Anyone with this file can use the add-in and see the team templates, so keep it within your team.');
}

// --------------------------------------------------------------------------- hand-off storage (IndexedDB)

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('thread-attachments', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('kv');
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function idbRequest(mode, act) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction('kv', mode);
      const request = act(tx.objectStore('kv'));
      tx.oncomplete = () => resolve(request.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally {
    db.close();
  }
}

const idbSet = (key, value) => idbRequest('readwrite', (store) => store.put(value, key));
const idbGet = (key) => idbRequest('readonly', (store) => store.get(key));
const idbDelete = (key) => idbRequest('readwrite', (store) => store.delete(key));

// --------------------------------------------------------------------------- writing an email: attach the PDF

async function initCompose() {
  showView('compose-view');
  setSummary('Attach a combined PDF to this email');
  let pending = null;
  try {
    pending = await idbGet(PENDING_KEY);
  } catch (error) {
    log(`Compose: couldn't check for a waiting PDF (${error.name}: ${error.message})`);
  }
  const text = $('compose-text');
  const button = $('btn-attach');
  const files = pending
    ? (Array.isArray(pending.files) ? pending.files : [{ name: pending.name, bytes: pending.bytes, pages: pending.pages }])
        .filter((f) => f && f.bytes && f.name)
        .map((f) => ({ ...f, bytes: f.bytes instanceof Uint8Array ? f.bytes : new Uint8Array(f.bytes) }))
    : [];
  if (!files.length || Date.now() - pending.created > PENDING_MAX_AGE_MS) {
    text.textContent = 'No combined PDF is waiting. Open Thread Attachments on a thread, check the files you want, then choose Email.';
    button.hidden = true;
    log('Compose: no PDF waiting');
    return;
  }
  const names = files.map((f) => f.name).join(', ');
  const size = humanSize(files.reduce((sum, f) => sum + f.bytes.length, 0));

  const attach = async () => {
    button.disabled = true;
    let done = 0;
    try {
      for (const file of files) {
        text.textContent = `Attaching ${file.name}…`;
        await call((cb) => Office.context.mailbox.item.addFileAttachmentFromBase64Async(bytesToBase64(file.bytes), file.name, { isInline: false }, cb));
        done += 1;
      }
      try {
        await idbDelete(PENDING_KEY);
      } catch (error) {
        // harmless: it expires on its own
      }
      log(`Compose: attached ${names} ✔`);
      text.textContent = `Attached ${names}.`;
      button.hidden = true;
    } catch (error) {
      log(`Compose: attaching failed after ${done} of ${files.length} (${error.message})`);
      text.textContent = `Couldn't attach ${files[done] ? files[done].name : 'the PDF'}: ${error.message}`;
      button.hidden = false;
      button.disabled = false;
    }
  };
  button.addEventListener('click', attach);

  if (pending.viaLink) {
    // Outlook attached these from the Azure link already; offer the button only as a backup.
    text.textContent = `This email should already have ${names} attached. If ${files.length > 1 ? "they're" : "it's"} missing, click Attach PDF.`;
    button.hidden = false;
    log('Compose: already attached through Azure; button offered as a backup');
    return;
  }

  // The email the Email button just opened: attach without another click.
  let subject = '';
  try {
    const field = Office.context.mailbox.item.subject;
    if (field && typeof field.getAsync === 'function') subject = String((await call((cb) => field.getAsync(cb))) || '');
  } catch (error) {
    subject = '';
  }
  if (pending.subject && subject.trim() === String(pending.subject).trim()) {
    log('Compose: this is the email Thread Attachments opened, attaching right away');
    await attach();
    return;
  }
  text.textContent = `Ready to attach: ${names} (${size}).`;
  button.hidden = false;
  log('Compose: PDF waiting');
}

// --------------------------------------------------------------------------- start-up

function wireReadUi() {
  $('check-all').addEventListener('change', (event) => {
    const on = event.target.checked;
    for (const file of selectableFiles()) {
      const was = state.checked.has(file.id);
      if (on) state.checked.add(file.id);
      else state.checked.delete(file.id);
      if (on && !was && state.multiStop) assignStop(file); // in list order: the next free stop each
    }
    state.built.clear();
    if (state.multiStop) renderList(); // stop pickers come and go with the ticks
    else updateSelectionUi();
    if (!$('preview').hidden) syncPreviewPages();
  });
  $('multi-stop').addEventListener('change', (event) => onMultiStopToggle(event.target.checked));
  $('show-logos').addEventListener('change', (event) => {
    state.showLogos = event.target.checked;
    state.built.clear();
    renderList();
  });
  $('btn-save').addEventListener('click', () => onSave().catch(fail));
  $('btn-email').addEventListener('click', () => ($('email-menu').hidden ? openEmailMenu() : closeEmailMenu(true)));
  $('email-menu').addEventListener('keydown', onMenuKeys);
  document.addEventListener('pointerdown', (event) => {
    if (!$('email-menu').hidden && !event.target.closest('.menu-wrap')) closeEmailMenu(false);
  });
  $('load-form').addEventListener('submit', onLoadSubmit);
  $('load-cancel').addEventListener('click', onLoadCancel);
  $('load-number').addEventListener('input', updateLoadHint);
  for (const kind of Object.keys(ACC_FIELDS)) {
    $(`acc-${kind}`).addEventListener('change', (event) => onAccessorialPick(kind, event.target.checked));
    $(`${kind}-at`).addEventListener('change', syncAccessorialBoxes);
  }
  for (const id of ['lum-total', 'oth-total']) $(id).addEventListener('change', onMoneyChange);
  $('load-form').addEventListener('input', () => { $('acc-error').textContent = ''; });
  $('tpl-back').addEventListener('click', () => (editing ? cancelEdit() : closeTemplates()));
  $('tpl-form').addEventListener('submit', (event) => onTemplateSubmit(event).catch(fail));
  $('tpl-cancel').addEventListener('click', cancelEdit);
  $('tpl-delete').addEventListener('click', () => onTemplateDelete().catch(fail));
  $('team-new').addEventListener('click', () => editTemplate('team', null));
  $('mine-new').addEventListener('click', () => editTemplate('mine', null));
  $('team-publish').addEventListener('click', () => onPublish().catch(fail));
  $('team-share').addEventListener('click', () => onShareInstall().catch(fail));
  $('az-gen').addEventListener('click', onAzureGenerate);
  $('az-copy').addEventListener('click', () => onAzureCopy().catch(fail));
  $('az-save').addEventListener('click', () => onAzureSave().catch(fail));
  $('az-test').addEventListener('click', () => onAzureTest().catch(fail));
  $('name-form').addEventListener('submit', onNameSubmit);
  $('file-name').addEventListener('input', updateNameHint);
  $('name-cancel').addEventListener('click', onNameCancel);
  $('pv-close').addEventListener('click', closePreview);
  $('pv-prev').addEventListener('click', () => stepPreview(-1));
  $('pv-next').addEventListener('click', () => stepPreview(1));
  $('pv-include').addEventListener('change', (event) => {
    const file = state.files[state.previewIndex];
    if (file) onRowToggle(file, event.target.checked);
  });
  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    if (!$('preview').hidden) closePreview();
    else if (!$('templates').hidden) {
      if (editing) cancelEdit();
      else closeTemplates();
    } else if (!$('email-menu').hidden) closeEmailMenu(true);
    else if (!$('name-form').hidden) onNameCancel();
    else if (!$('load-form').hidden) onLoadCancel();
  });
}

function registerHandlers() {
  const mailbox = Office.context.mailbox;
  const reload = debounce(() => loadThread().catch(fail), 350);
  const note = (what) => (result) => {
    if (result && result.status === Office.AsyncResultStatus.Failed) log(`Couldn't watch ${what}: ${result.error && result.error.message}`);
  };
  if (Office.EventType.SelectedItemsChanged && supports('1.13')) {
    mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, reload, note('selection changes'));
  }
  if (Office.EventType.ItemChanged) {
    mailbox.addHandlerAsync(Office.EventType.ItemChanged, reload, note('item changes'));
  }
}

function start() {
  $('btn-copy-report').addEventListener('click', () => copyReport());
  window.addEventListener('error', (event) => log(`Script error: ${event.message}`));
  window.addEventListener('unhandledrejection', (event) => log(`Unhandled: ${event.reason && (event.reason.message || event.reason)}`));
  if (typeof Office === 'undefined') {
    showMessage("Outlook's add-in library didn't load. Check the internet connection, then close and reopen this panel.");
    setSummary('');
    return;
  }
  Office.onReady(async (info) => {
    try {
      collectEnvironment();
      if (!info || info.host !== Office.HostType.Outlook) {
        setSummary('');
        showMessage('Open this from Outlook: select an email, then click Thread Attachments on the ribbon.');
        return;
      }
      const item = Office.context.mailbox.item;
      if (item && typeof item.addFileAttachmentFromBase64Async === 'function') {
        initCompose().catch(fail);
        return;
      }
      setSummary('Checking your team settings…');
      await loadTeam();
      if (team.status === 'nokey' || team.status === 'wrongkey' || team.status === 'error') {
        showLocked();
        return;
      }
      state.myTemplates = cleanTemplates(readRoaming(MY_TEMPLATES_KEY));
      state.saveSplitMb = parseSplitMb(readRoaming('saveSplitMb'));
      wireReadUi();
      registerHandlers();
      loadThread().catch(fail);
    } catch (error) {
      fail(error);
    }
  });
}

start();
