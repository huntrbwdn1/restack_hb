/*
 * Thread Attachments - Outlook add-in (test build).
 *
 * Everything happens inside this side panel: attachments are read through Outlook's
 * add-in API and combined here with pdf-lib, PDF.js and UTIF.js. Nothing is uploaded.
 */

const VERSION = '0.4.0';
const SMALL_INLINE_BYTES = 50 * 1024; // logo test for images whose pixel size can't be read
const LOGO_MAX_SIDE = 400;            // images in the email body this small (in pixels) are logos/icons
const LOGO_THIN_SIDE = 120;           // ...and so are thin banners
const MAX_IMAGE_SIDE = 2400;          // photos are scaled to this before going into a PDF
const LOCKED_PDF_DPI = 200;           // locked PDFs are copied in as page images at this quality
const PRINT_DPI = 150;
const PREVIEW_PAGE_LIMIT = 40;
const PENDING_KEY = 'pending-pdf';
const PENDING_MAX_AGE_MS = 2 * 60 * 60 * 1000;
const PDF_TYPES = [{ description: 'PDF file', accept: { 'application/pdf': ['.pdf'] } }];

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

const $ = (id) => document.getElementById(id);

const state = {
  files: [],
  checked: new Set(),
  showLogos: false,
  topic: '',
  fileName: '',          // the name typed in "Name the PDF" (for the current thread)
  messages: 0,
  failed: 0,
  loadToken: 0,
  busy: false,
  reloadPending: false,
  combined: null,
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

/** The PDF's file name: the name you typed for this thread, else the email subject. */
function pdfName() {
  return `${safeName(state.fileName || state.topic || 'Combined attachments')}.pdf`;
}

/** What was typed in "Name the PDF", made safe for Windows. Blank means "use the subject". */
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
  clearPrint();
  hideNameForm();
  state.fileName = '';
  state.combined = null;
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
  state.checked = new Set(files.filter((f) => canCombine(f) && !f.smallLogo).map((f) => f.id));
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
  const note = pageNote(file);
  if (note) text.append(el('span', 'file-note picked', note));
  label.append(box, text);

  const preview = document.createElement('button');
  preview.type = 'button';
  preview.className = 'icon-btn';
  preview.title = 'Preview';
  preview.setAttribute('aria-label', `Preview ${file.name}`);
  preview.innerHTML = EYE_SVG;
  preview.addEventListener('click', () => openPreview(file.id));
  row.append(label, preview);
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
  $('selected-count').textContent = selectable.length
    ? `${count} of ${selectable.length} selected`
    : 'Nothing here can go into a PDF';
  const anyChecked = checkedFiles().length > 0;
  for (const id of ['btn-save', 'btn-print', 'btn-email', 'name-save']) $(id).disabled = state.busy || !anyChecked;
  for (const box of document.querySelectorAll('#file-list input[type="checkbox"]')) {
    const file = state.files[Number(box.closest('.file').dataset.id)];
    box.disabled = state.busy || !canCombine(file);
    box.checked = state.checked.has(file.id);
  }
  if (!$('preview').hidden) syncPreviewInclude();
}

function setChecked(id, on) {
  if (on) state.checked.add(id);
  else state.checked.delete(id);
  state.combined = null;
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
    const hadFocus = row.contains(document.activeElement);
    const fresh = fileRow(file);
    row.replaceWith(fresh);
    if (hadFocus) fresh.querySelector('input').focus();
  }
  updateSelectionUi();
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

/** A previewed PDF page with its own "include this page" checkbox. */
function pageHolder(file, index, canvas) {
  const holder = el('div', 'pv-page');
  holder.dataset.page = String(index);
  const bar = el('label', 'pv-page-bar');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.setAttribute('aria-label', `Include page ${index + 1}`);
  box.checked = isPageIncluded(file, index); // right from the start, while later pages still render
  box.disabled = state.busy || !canCombine(file);
  holder.classList.toggle('excluded', !box.checked);
  box.addEventListener('change', () => onPreviewPageToggle(file));
  bar.append(box, el('span', '', `Page ${index + 1}`));
  holder.append(bar, canvas);
  return holder;
}

function isPageIncluded(file, index) {
  if (!state.checked.has(file.id)) return false;
  const pages = includedPages(file);
  return pages === null || pages.includes(index);
}

/** Brings the preview's page checkboxes in line with the file. */
function syncPreviewPages() {
  const file = state.files[state.previewIndex];
  if (!file) return;
  for (const holder of document.querySelectorAll('#pv-body .pv-page')) {
    const index = Number(holder.dataset.page);
    const box = holder.querySelector('input');
    box.checked = isPageIncluded(file, index);
    box.disabled = state.busy || !canCombine(file);
    holder.classList.toggle('excluded', !box.checked);
  }
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
  state.combined = null;
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

/** Every page of a TIFF (fax scans are usually multi-page), as canvases plus resolution. */
function decodeTiffPages(bytes) {
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

/** Combines the checked files, in list order. Cached until the selection changes. */
async function buildCombined(files) {
  const key = files.map((f) => `${f.id}:${(includedPages(f) || ['all']).join('.')}`).join(',');
  if (state.combined && state.combined.key === key) return state.combined;
  const out = await PDFLib.PDFDocument.create();
  const skipped = [];
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    setStatus(`Adding ${i + 1} of ${files.length}: ${file.name}`);
    try {
      if (file.kind === 'PDF') {
        const only = includedPages(file);
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
  state.combined = { key, bytes, pages, skipped, files: files.length };
  return state.combined;
}

function leftOutText(result) {
  let text = '';
  if (result.skipped.length) text += ` Left out: ${result.skipped.map((s) => `${s.name} (${s.reason})`).join('; ')}.`;
  return text;
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
  const result = await buildCombined(files);
  if (!result.pages) {
    setStatus(`None of the checked files could go into a PDF.${leftOutText(result)}`, 'error');
    return null;
  }
  return result;
}

// --------------------------------------------------------------------------- save, print, new email

/**
 * Save PDF. Where Outlook allows Windows' own Save As window, that opens straight away (you can
 * name the file and pick the folder there). Everywhere else, you name the file in the panel first.
 */
async function onSave() {
  if (!checkedFiles().length) return;
  if (typeof window.showSaveFilePicker === 'function' && readSetting('saveDialog') === 'works') {
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
  input.value = state.fileName || safeName(state.topic || 'Combined attachments');
  input.focus();
  input.select();
}

function hideNameForm() {
  $('name-form').hidden = true;
  $('save-actions').hidden = false;
}

function onNameSubmit(event) {
  event.preventDefault();
  if (state.busy || !checkedFiles().length) return;
  state.fileName = cleanFileName($('file-name').value);
  hideNameForm();
  const name = pdfName();
  log(`Save: named ${name}`);
  // Straight from the click, so the browser lets Windows' Save As window open if it can.
  pickSaveLocation(name)
    .then((picked) => (picked.cancelled ? null : finishSave(picked.handle || null, name)))
    .catch(fail);
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
    setStatus(`Downloading ${name} (${plural(result.pages, 'page')}) to your Downloads folder. If it doesn't show up, use Print… and choose "Save as PDF".${leftOutText(result)}`, 'ok');
  } catch (error) {
    log(`Download: failed (${error.name}: ${error.message})`);
    setStatus('Saving isn\'t possible here. Use Print… and choose "Save as PDF", or Put in new email.', 'error');
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

async function preparePrint(bytes) {
  clearPrint();
  const doc = await openPdf(bytes);
  const root = $('print-root');
  try {
    for (let p = 1; p <= doc.numPages; p++) {
      setStatus(`Getting page ${p} of ${doc.numPages} ready to print…`);
      const page = await doc.getPage(p);
      const viewport = page.getViewport({ scale: PRINT_DPI / 72 });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      await page.render({ canvas, viewport }).promise;
      page.cleanup();
      const img = new Image();
      img.alt = `Page ${p}`;
      img.src = canvas.toDataURL('image/jpeg', 0.92);
      const holder = el('div', 'pp');
      holder.append(img);
      root.append(holder);
      try {
        await img.decode();
      } catch (error) {
        // printing still works; the image just finishes loading later
      }
    }
  } finally {
    closePdf(doc);
  }
}

function clearPrint() {
  $('print-root').textContent = '';
}

async function onPrint() {
  const result = await withBusy(async () => {
    const combined = await combineChecked();
    if (!combined) return null;
    await preparePrint(combined.bytes);
    return combined;
  });
  if (!result) return;

  let before = false;
  let after = false;
  const pageTitle = document.title;
  const restoreTitle = () => { document.title = pageTitle; };
  const onBefore = () => { before = true; };
  const onAfter = () => {
    after = true;
    clearPrint();
    restoreTitle();
  };
  window.addEventListener('beforeprint', onBefore);
  window.addEventListener('afterprint', onAfter, { once: true });
  const started = performance.now();
  document.title = pdfName().replace(/\.pdf$/i, ''); // the name "Save as PDF" suggests
  try {
    window.print();
  } catch (error) {
    log(`Print: failed (${error.name}: ${error.message})`);
  }
  const ms = Math.round(performance.now() - started);
  setTimeout(() => {
    window.removeEventListener('beforeprint', onBefore);
    if (!before || after) restoreTitle();
    log(`Print: print() returned after ${ms} ms; print window ${before ? 'opened' : 'did not open'}${after ? ' and closed' : ''}`);
    if (!before && ms < 150) {
      setStatus('The print window didn\'t open here. Try Save PDF… or Put in new email.', 'error');
    } else {
      setStatus(`Sent ${plural(result.pages, 'page')} to the print window. To save a file, choose "Save as PDF" or "Microsoft Print to PDF" as the printer.${leftOutText(result)}`, 'ok');
    }
  }, 400);
}

async function onEmail() {
  const result = await withBusy(combineChecked);
  if (!result) return;
  const name = pdfName();
  try {
    await idbSet(PENDING_KEY, { name, bytes: result.bytes, pages: result.pages, created: Date.now() });
    log('New email: PDF stored for the Attach PDF button');
  } catch (error) {
    log(`New email: couldn't store the PDF (${error.name}: ${error.message})`);
    setStatus("Couldn't hand the PDF over to a new email here.", 'error');
    return;
  }
  const mailbox = Office.context.mailbox;
  const form = { subject: state.topic ? `${state.topic} - attachments` : 'Attachments' };
  try {
    if (typeof mailbox.displayNewMessageFormAsync === 'function') {
      await call((done) => mailbox.displayNewMessageFormAsync(form, done));
    } else {
      mailbox.displayNewMessageForm(form);
    }
    log('New email: opened');
    setStatus(`A new email opened. In it, click Attach PDF on the ribbon.${leftOutText(result)}`, 'ok');
  } catch (error) {
    log(`New email: couldn't open one (${error.message})`);
    setStatus(`Couldn't open a new email: ${error.message}`, 'error');
  }
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
  if (!pending || !pending.bytes || Date.now() - pending.created > PENDING_MAX_AGE_MS) {
    text.textContent = 'No combined PDF is waiting. Open Thread Attachments on a thread, check the files you want, then choose Put in new email.';
    button.hidden = true;
    log('Compose: no PDF waiting');
    return;
  }
  const bytes = pending.bytes instanceof Uint8Array ? pending.bytes : new Uint8Array(pending.bytes);
  text.textContent = `Ready to attach: ${pending.name} (${plural(pending.pages || 0, 'page')}, ${humanSize(bytes.length)}).`;
  button.hidden = false;
  log('Compose: PDF waiting');
  button.addEventListener('click', async () => {
    button.disabled = true;
    try {
      await call((done) => Office.context.mailbox.item.addFileAttachmentFromBase64Async(bytesToBase64(bytes), pending.name, { isInline: false }, done));
      try {
        await idbDelete(PENDING_KEY);
      } catch (error) {
        // harmless: it expires on its own
      }
      log('Compose: PDF attached ✔');
      text.textContent = `Attached ${pending.name}.`;
      button.hidden = true;
    } catch (error) {
      log(`Compose: attaching failed (${error.message})`);
      text.textContent = `Couldn't attach the PDF: ${error.message}`;
      button.disabled = false;
    }
  });
}

// --------------------------------------------------------------------------- start-up

function wireReadUi() {
  $('check-all').addEventListener('change', (event) => {
    const on = event.target.checked;
    for (const file of selectableFiles()) {
      if (on) state.checked.add(file.id);
      else state.checked.delete(file.id);
    }
    state.combined = null;
    updateSelectionUi();
  });
  $('show-logos').addEventListener('change', (event) => {
    state.showLogos = event.target.checked;
    state.combined = null;
    renderList();
  });
  $('btn-save').addEventListener('click', () => onSave().catch(fail));
  $('btn-print').addEventListener('click', () => onPrint().catch(fail));
  $('btn-email').addEventListener('click', () => onEmail().catch(fail));
  $('name-form').addEventListener('submit', onNameSubmit);
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
    else if (!$('name-form').hidden) onNameCancel();
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
  Office.onReady((info) => {
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
      wireReadUi();
      registerHandlers();
      loadThread().catch(fail);
    } catch (error) {
      fail(error);
    }
  });
}

start();
