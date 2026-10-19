/*
 * Restack - Outlook add-in (test build).
 *
 * Everything happens inside this side panel: attachments are read through Outlook's
 * add-in API and combined here with pdf-lib, PDF.js and UTIF.js. The only upload is the
 * optional "attach automatically" step, to the team admin's own PDF Emails Azure function.
 */

const VERSION = '0.20.0';
const SMALL_INLINE_BYTES = 50 * 1024; // logo test for images whose pixel size can't be read
const LOGO_MAX_SIDE = 400;            // images in the email body this small (in pixels) are logos/icons
const LOGO_THIN_SIDE = 120;           // ...and so are thin banners
const MAX_IMAGE_SIDE = 2400;          // fax scans and locked PDF pages are drawn at most this big
const PHOTO_MAX_SIDE = 2000;          // photos are scaled to this (long side) before going into a PDF: the whole photo, smaller
const LOCKED_PDF_DPI = 200;           // locked PDFs are copied in as page images at this quality
const PREVIEW_PAGE_LIMIT = 40;
const THUMB_PX = 240;                 // list pictures are drawn this big: sharp at every View and text size
const PENDING_KEY = 'pending-pdf';
const PENDING_MAX_AGE_MS = 2 * 60 * 60 * 1000;
const MB = 1000 * 1000;                // sizes are in MB as people read them (not MiB)
const SPLIT_MB = 10;                   // every PDF stays under this; bigger ones split into -1, -2…
const SPLIT_LIMIT = SPLIT_MB * MB - 1;
const AZURE_FUNCTION = 'PdfEmails';    // the function's name in the team admin's Function App
const PDF_TYPES = [{ description: 'PDF file', accept: { 'application/pdf': ['.pdf'] } }];
const TEAM_FILE = 'team.json';
// The "sent with restack" line under the message of every email Restack creates (a team setting, off to
// start): one small picture from this site, attached inside the email so it shows without a "download
// pictures" prompt and looks the same in every mail app. Drawn at 3x for sharp screens, shown at this size.
const SENT_WITH = { file: 'sent-with-restack.png', width: 136, height: 23, alt: 'sent with restack' };
const MAX_ADMINS = 20;                  // team admins (Settings → Team → Team admins)
// The beta: the site's beta folder, installed as "Restack Beta" next to Restack, for trying a new
// version before reps get it. It reads the main site's team.json and keeps its own saved settings.
const BETA = /\/beta\/[^/]*$/.test(window.location.pathname);
const SITE_ROOT = new URL(BETA ? '../' : './', window.location.href).href; // the main site
const BETA_ID = '8f3a2c71-5d4e-4b6a-9c2f-71e0d4b85a93';                      // Restack Beta's add-in ID (never changes)
// The admin console: this page opened in a browser (console.html), for the team's templates, settings and usage
const CONSOLE = new URLSearchParams(window.location.search).has('console');
// This browser's saved bits: thread-attachments.* (the beta's under .beta., the console's under .console.)
const STORE_PREFIX = `thread-attachments.${BETA ? 'beta.' : ''}${CONSOLE ? 'console.' : ''}`;
const MY_TEMPLATES_KEY = 'myTemplates'; // each person's own templates (Outlook roaming settings)
const DRAFT_KEY = 'teamDraft';          // the team admin's unpublished edits
const NEW_TEMPLATE = { body: 'Attached are the POD(s) for load {load}.', file: '{load}', lumperName: '{load} Receipt' };
const OLD_RECEIPT_NAME = '{load} Lumper'; // the receipt file's name before 0.16; templates still on it get the new one
const STOP_CHOICES = 9;                 // the stop picker offers at least Stop 1-9
const QUIZ = {                          // a template's questions, asked before the Create email step
  maxQuestions: 12,
  maxAnswers: 8,
  intro: 'This is a {company} load. Please answer the following before continuing:',
};
const RC = {                            // the rate confirmation check
  minWords: 15,                         // digital pages with fewer words than this are too sparse to be one
  scanWords: 3,                         // pages with fewer words than this are scans, compared by layout
  textMatch: 0.45,                      // share of the sample's words found on a page
  layoutMatch: 0.7,                     // layout likeness for scanned pages and photos (-1 to 1)
  boxSlack: 0.2,                        // a scan's printed area may differ this much (share of the page) from the sample's
  maxPages: 30,                         // pages checked per file
  samplePages: 3,                       // pages kept from the sample rate confirmation
  sampleWords: 250,                     // words kept per sample page
  gridW: 24,                            // a page's layout is ink density in a 24 x 32 grid
  gridH: 32,
  renderW: 220,                         // pages are drawn this many pixels wide to measure their layout
};

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
const X_SVG = '<svg viewBox="0 0 20 20" width="14" height="14" aria-hidden="true"><path d="m5 5 10 10M15 5 5 15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>';
const GRIP_SVG = '<svg viewBox="0 0 8 14" width="7" height="12" aria-hidden="true"><g fill="currentColor"><circle cx="2" cy="2" r="1.2"/><circle cx="6" cy="2" r="1.2"/><circle cx="2" cy="7" r="1.2"/><circle cx="6" cy="7" r="1.2"/><circle cx="2" cy="12" r="1.2"/><circle cx="6" cy="12" r="1.2"/></g></svg>';
const UP_SVG = '<svg viewBox="0 0 20 20" width="14" height="14" aria-hidden="true"><path d="M5 12.5 10 7.5l5 5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const DOWN_SVG = '<svg viewBox="0 0 20 20" width="14" height="14" aria-hidden="true"><path d="M5 7.5 10 12.5l5-5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const STOP_SVG = '<svg viewBox="0 0 20 20" width="20" height="20" aria-hidden="true"><path d="M7 2.5h6L17.5 7v6L13 17.5H7L2.5 13V7Z" fill="#b10e1c"/><path d="M6.5 10h7" stroke="#fff" stroke-width="2" stroke-linecap="round"/></svg>';
// The how-to animation's little pictures (a mouse pointer, Outlook's expand arrows, a paperclip)
const CURSOR_SVG = '<svg viewBox="0 0 16 22" width="16" height="22" aria-hidden="true"><path d="M1 1v17.5l4.4-4.2 2.9 6.6 2.7-1.2-2.9-6.5h6.1Z" fill="#fff" stroke="#242424" stroke-width="1.3" stroke-linejoin="round"/></svg>';
const CHEVRON_SVG = '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="m3.5 1.5 3.5 3.5-3.5 3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const TRIANGLE_SVG = '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M3 1.5v7L8 5Z" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round"/></svg>';
// A pushpin: Axle's orange on a pinned template, an outline on the others
const PIN_SVG = '<svg class="pin-icon" viewBox="0 0 20 20" width="16" height="16" aria-hidden="true"><g transform="rotate(45 10 10)"><rect x="6.2" y="1.8" width="7.6" height="2.8" rx="1.3"/><path d="M7.3 4.6h5.4l-.7 4.9 2.7 2.5v1.5H5.3V12l2.7-2.5z"/><path class="pin-needle" d="M10 13.6v4.6"/></g></svg>';
const CLIP_SVG = '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M8.2 3.6 4.6 7.2a.9.9 0 0 0 1.3 1.3l3.9-3.9a1.9 1.9 0 0 0-2.7-2.7L3.2 5.8a2.9 2.9 0 0 0 4.1 4.1L9.6 7.6" fill="none" stroke="currentColor" stroke-width=".9" stroke-linecap="round"/></svg>';

const $ = (id) => document.getElementById(id);

// =============================================================================================
// WORDING: most of what reps read on the panel, in one place. To reword something, change the
// text between the double quotes. Keep the double quotes, the comma at the end, and anything in
// {curly brackets}. An apostrophe is fine; a double quote inside the text isn't. Then upload
// taskpane.js: the label changes everywhere it's used, within a few minutes (or close and reopen
// Outlook). Messages and warnings, and the template screens, aren't in this list.
// =============================================================================================
const WORDS = {
  // The list of files
  selectAll: "Select all",                         // the checkbox above the list
  selected: "{n} of {total} selected",             // next to Select all
  multiStop: "Multi-stop",                         // the checkbox above the list
  multiStopNote: "one PDF per stop: -1, -2…",      // the grey text after Multi-stop
  stop: "Stop {n}",                                // the stop picker on each file (and page), with Multi-stop on
  byPage: "By page",                               // a file's stop picker when its pages are at different stops
  receipt: "Receipt",                              // the tag on each checked file, and on pages in the preview
  receiptPages: "Receipt: {pages}",                // under a file that's only partly a receipt, like "Receipt: page 2"
  oneEmailTip: "Only 1 email came through. To include the whole thread, select every email in it.",
  keepAddingTip: "Or keep adding as you click each one",  // the button under that tip
  keepAddingNote: "Click each email in the thread to add its files.", // in the Keep adding bar
  showLogos: "Show logos & icons",                 // under the list, when emails have logos in them

  // The buttons under the list, and the Email menu
  email: "Email",
  savePdf: "Save PDF…",
  blankEmail: "Blank email",
  blankEmailNote: "Just the PDF, no recipients",   // under Blank email
  team: "Team",                                    // the menu's headings over the templates
  mine: "Mine",
  pinned: "Pinned",                                // the heading over your pinned templates
  otherTemplates: "Other templates",               // under them: opens the rest, A to Z
  pin: "Pin to the top",                           // the pin on each template
  unpin: "Unpin",
  manageTemplates: "Manage templates…",
  createTemplate: "Create a template…",            // in place of Manage templates, before there are any

  // Save PDF
  nameThePdf: "Name the PDF",
  save: "Save",
  cancel: "Cancel",                                // Save PDF and Create email

  // Create email (the accessorials' names and fields are written the same way in the email)
  loadNumber: "Load number",
  accessorials: "Accessorials",
  detention: "Detention",
  layover: "Layover",
  lumper: "Lumper",
  other: "Other",
  timeIn: "IN",                                    // Detention's and Layover's times
  appointment: "APPT",
  timeOut: "OUT",
  hoursRequested: "Hours requested",
  total: "Total",                                  // Lumper's and Other's amount
  otherType: "Type",                               // what the Other accessorial is
  shipper: "Shipper",                              // where a charge was
  receiver: "Receiver",
  receiverAtStop: "Receiver (Stop {n})",           // with Multi-stop on
  shipperOrReceiver: "Shipper or receiver?",       // the drop-down's prompt
  shipperOrStop: "Shipper or which stop?",         // the same, with Multi-stop on
  createEmail: "Create email",
  subject: "{load} POD(s)",                        // every email's subject

  // The separate accessorial email (for templates with "Accessorials in a separate email")
  accessorialSubject: "{load} - {types} Request",  // its subject; {types} is like Detention/Lumper
  accessorialIntro: "Please see the {types} request for load {load} below.", // its first line, unless the template has its own

  // The rate confirmation check
  looksLikeRateCon: "Looks like a rate confirmation", // under a flagged file
  leaveItOut: "Leave it out",
  leaveThemOut: "Leave them out",
  notARateCon: "Not a rate con",
  rateConBadge: "Rate con?",                       // on a flagged page in the preview
  useAsRcSample: "Use as rate con sample",         // the team admin's link at the bottom of the preview

  // The preview
  includeInPdf: "Include in PDF",
  rotate: "Rotate",                                // the button on each page and photo (turns it a quarter clockwise)
  movePageUp: "Move page up",                      // the arrows on each page of a PDF or fax in the preview
  movePageDown: "Move page down",
  pageOrder: "New page order: {pages}",            // under a file whose pages were moved, like "New page order: 2, 1, 3"
  resetOrder: "Reset page order",                  // in the preview, once pages have been moved

  // The bottom of the panel, and Settings (the gear across from Test report)
  testReport: "Test report",
  copyReport: "Copy report",
  settings: "Settings",
  appearance: "Appearance",                        // drops down Light, Dark and Match Outlook
  light: "Light",
  dark: "Dark",
  matchOutlook: "Match Outlook",
  textSize: "Text size",                           // drops down Small, Medium and Large
  textSmall: "Small",
  textMedium: "Medium",
  textLarge: "Large",
  view: "View",                                    // how the list shows files, like File Explorer's View menu
  viewLarge: "Large pictures",
  viewMedium: "Medium pictures",
  viewSmall: "Small pictures",
  viewList: "List",
  viewDetails: "Details",
  viewTiles: "Tiles",
  tagReceipts: "Tag receipts by file name",        // a switch: files named like receipts start out tagged Receipt
  keepAdding: "Keep adding",                       // a switch: each email you click adds its files to the list

  // A new email
  attachPdf: "Attach PDF",                         // in the panel and messages (the ribbon's is in the install file)
};

/** A WORDS entry with its {placeholders} filled in. */
function word(key, values = {}) {
  const text = Object.prototype.hasOwnProperty.call(WORDS, key) ? WORDS[key] : key;
  return String(text).replace(/\{(\w+)\}/g, (match, name) => (Object.prototype.hasOwnProperty.call(values, name) ? String(values[name]) : match));
}

/** Puts WORDS into the page's fixed labels: elements marked data-w (text), data-w-aria and data-w-title. */
function applyWords() {
  for (const node of document.querySelectorAll('[data-w]')) node.textContent = word(node.dataset.w);
  for (const node of document.querySelectorAll('[data-w-aria]')) node.setAttribute('aria-label', word(node.dataset.wAria));
  for (const node of document.querySelectorAll('[data-w-title]')) node.title = word(node.dataset.wTitle);
}

/** The team key from the install file (?team=… in the add-in's address). Read first thing. */
function readTeamKey() {
  try {
    // The console: from the link Outlook opened (#k=…), or kept in this browser from last time
    const key = CONSOLE
      ? new URLSearchParams(window.location.hash.replace(/^#/, '')).get('k') || readSetting('teamKey')
      : new URLSearchParams(window.location.search).get('team');
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
  multiStop: false,      // Multi-stop: one PDF per stop, named -1, -2… (for the current thread)
  order: null,           // file ids in the order you put them (null: as they came, oldest first)
  places: new Map(),     // each checked file's place in its PDF (1, 2, 3…), worked out by renderList
  messages: 0,
  failed: 0,
  seen: new Set(),       // the emails the list was read from (Keep adding skips them)
  keepWaiting: new Map(), // Keep adding: emails clicked that haven't gone in yet (itemId -> pick), so quick clicks all count
  fresh: false,          // the next load starts a new list even with Keep adding on (Start over)
  loadToken: 0,
  busy: false,
  reloadPending: false,
  built: new Map(),      // PDFs already built for the current selection, so Save then Email doesn't rebuild
  rcPending: false,      // the rate confirmation check is still going through the files
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

/** A note under a file, where "pages 1–3" (or "page 2") stays on one line. */
function noteEl(className, text) {
  const node = el('span', className);
  String(text).split(/(pages? \d+(?:–\d+)?|\d+–\d+)/i).forEach((part, i) => {
    if (part) node.append(i % 2 ? el('span', 'nowrap', part) : part);
  });
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

/** Sizes in KB and MB of 1000, the same MB as the 10 MB limit, so a file shown as 9.8 MB isn't split. */
function humanSize(n) {
  if (n < 1000) return `${n} B`;
  if (n < 999500) return `${Math.round(n / 1000)} KB`;
  return `${(n / MB).toFixed(1)} MB`;
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
    return localStorage.getItem(`${STORE_PREFIX}${key}`);
  } catch (error) {
    return null;
  }
}

function writeSetting(key, value) {
  try {
    if (value === null) localStorage.removeItem(`${STORE_PREFIX}${key}`);
    else localStorage.setItem(`${STORE_PREFIX}${key}`, value);
  } catch (error) {
    // not remembered; harmless
  }
}

/** "10/14 3:00 PM": short enough for the Details view's column. */
function shortDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  const day = date.toLocaleDateString(undefined, { month: 'numeric', day: 'numeric' });
  const time = date.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  return `${day} ${time}`;
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
  const lines = [`Restack ${VERSION}${BETA ? ' beta' : ''}${CONSOLE ? ' admin console' : ''} (test build)`];
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
  setTimeout(() => { button.textContent = word('copyReport'); }, 2500);
}

function fail(error) {
  const message = (error && (error.userMessage || error.message)) || String(error);
  log(`Error: ${message}`);
  if (CONSOLE) consoleNote(message, 'error');
  else setStatus(message, 'error');
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

// --------------------------------------------------------------------------- only one email came through: tip and how-to

/**
 * When only one email came through, a tip under the summary. Outlook doesn't say how many emails
 * the conversation has, so it shows whenever there's just one. Its info icon (hover or click)
 * plays a short how-to for selecting every email, for new Outlook or classic Outlook.
 */
function renderOneEmailTip(messages) {
  const show = messages === 1 && !prefs.keepAdding; // with Keep adding on, its bar says what to do
  $('one-tip').hidden = !show;
  if (!show) closeHowto(false);
  else log('Tip: only 1 email came through');
}

const HOWTO = {
  new: {
    label: 'New Outlook',
    steps: [
      'Click › next to the conversation to open it.',
      'Tick the checkbox at the top of the conversation.',
      'Every email is selected, so Restack reads them all.',
    ],
    note: '',
    script: [ // what the pointer does, step by step (wait: ms before the next frame)
      { step: 0, state: 'collapsed', cursor: 'rest', wait: 900 },
      { cursor: 'chev', wait: 850 },
      { click: true, state: 'expanded', wait: 1000 },
      { step: 1, cursor: 'head-check', wait: 850 },
      { click: true, state: 'all', wait: 500 },
      { step: 2, cursor: 'rest', wait: 2800 },
    ],
  },
  classic: {
    label: 'Classic Outlook',
    steps: [
      'Click ▷ next to the conversation to open it.',
      'Click the first email under it.',
      'Hold Shift and click the last email.',
      'Every email is selected, so Restack reads them all.',
    ],
    note: 'To pick only some of them, hold Ctrl and click each one.',
    script: [
      { step: 0, state: 'collapsed', cursor: 'rest', wait: 900 },
      { cursor: 'chev', wait: 850 },
      { click: true, state: 'expanded', wait: 1000 },
      { step: 1, cursor: 'first', wait: 850 },
      { click: true, state: 'first', wait: 700 },
      { step: 2, shift: true, cursor: 'last', wait: 1100 },
      { click: true, state: 'range', wait: 500 },
      { step: 3, shift: false, cursor: 'rest', wait: 2800 },
    ],
  },
  flat: { // mail shown as individual messages, not conversations: Keep adding
    label: 'No conversations',
    steps: [
      'Pin Restack: click the pin at the top of its panel, so it stays open.',
      'Turn on Keep adding (under this tip, or in Settings).',
      'Click each email in the thread. Restack adds its files to the list.',
    ],
    note: 'Or search for the load number and press Ctrl+A to select every result at once. They have to be in the same folder.',
    script: [
      { step: 0, state: 'flat', cursor: 'rest', wait: 1300 },
      { step: 1, wait: 1300 },
      { step: 2, cursor: 'r1', wait: 850 },
      { click: true, mark: 'r1', wait: 900 },
      { cursor: 'r3', wait: 850 },
      { click: true, mark: 'r3', wait: 900 },
      { cursor: 'r5', wait: 850 },
      { click: true, mark: 'r5', wait: 700 },
      { cursor: 'rest', wait: 2600 },
    ],
  },
};
const HOWTO_TABS = [['howto-tab-new', 'new'], ['howto-tab-classic', 'classic'], ['howto-tab-flat', 'flat']];

// The made-up conversation the how-to plays with
const HOWTO_EMAILS = [
  { who: 'Dispatch', what: 'POD stop 1', clip: true },
  { who: 'You', what: 'Got it, thanks!', clip: false },
  { who: 'Driver', what: 'Lumper receipt', clip: true },
  { who: 'Dispatch', what: 'POD stop 2', clip: true },
];
// ...and the same load's emails among others, for mail that isn't shown as conversations
const HOWTO_FLAT = [
  { who: 'Dispatch', what: 'Load 3645493 PODs', files: 2 },
  { who: 'Billing', what: 'Invoice 88123', files: 0 },
  { who: 'Driver', what: 'RE: Load 3645493 PODs', files: 1 },
  { who: 'Broker', what: 'New load offer', files: 0 },
  { who: 'Dispatch', what: 'RE: Load 3645493 PODs', files: 1 },
];

const howto = { client: 'new', detected: '', pinned: false, token: 0, hoverTimer: 0, leaveTimer: 0 };

/** Which Outlook this is: 'classic', 'new' (new Outlook or the web) or '' when it doesn't say. */
function detectOutlook() {
  try {
    const host = String((Office.context.mailbox.diagnostics || {}).hostName || '');
    if (host === 'Outlook') return 'classic';
    if (host === 'newOutlookWindows' || host === 'OutlookWebApp') return 'new';
  } catch (error) {
    // not told
  }
  return '';
}

function hostLabel() {
  try {
    const host = String((Office.context.mailbox.diagnostics || {}).hostName || '');
    if (host === 'OutlookWebApp') return 'Outlook on the web, which works like New Outlook';
  } catch (error) {
    // not told
  }
  return howto.detected ? HOWTO[howto.detected].label : '';
}

function buildHowtoStage(client) {
  const stage = $('howto-stage');
  stage.textContent = '';
  stage.dataset.client = client;
  stage.dataset.state = 'collapsed';
  stage.classList.remove('shift');
  if (client === 'flat') {
    buildFlatStage(stage);
    return;
  }
  const list = el('div', 'ml');
  const head = el('div', 'ml-row ml-head');
  const chev = el('span', 'ml-chev');
  chev.dataset.target = 'chev';
  chev.innerHTML = client === 'classic' ? TRIANGLE_SVG : CHEVRON_SVG;
  head.append(chev);
  if (client === 'new') {
    const box = el('span', 'ml-check');
    box.dataset.target = 'head-check';
    head.append(box);
  }
  const title = el('span', 'ml-lines');
  title.append(el('b', '', 'Load 3645493 PODs'), el('i', '', `${HOWTO_EMAILS.length} emails`));
  head.append(title);
  const kids = el('div', 'ml-kids');
  HOWTO_EMAILS.forEach((email, i) => {
    const row = el('div', 'ml-row ml-kid');
    if (i === 0) row.dataset.target = 'first';
    if (i === HOWTO_EMAILS.length - 1) row.dataset.target = 'last';
    if (client === 'new') row.append(el('span', 'ml-check'));
    const text = el('span', 'ml-lines');
    text.append(el('b', '', email.who), el('i', '', email.what));
    row.append(text);
    if (email.clip) {
      const clip = el('span', 'ml-clip');
      clip.innerHTML = CLIP_SVG;
      row.append(clip);
    }
    kids.append(row);
  });
  list.append(head, kids);
  const ring = el('span', 'ml-ring');
  const cursor = el('span', 'ml-cursor');
  cursor.innerHTML = CURSOR_SVG;
  stage.append(list, ring, cursor, el('span', 'ml-key', 'Shift'));
}

/** The individual-messages picture: the load's emails among others, and what Restack has added so far. */
function buildFlatStage(stage) {
  stage.dataset.state = 'flat';
  const list = el('div', 'ml ml-flat');
  HOWTO_FLAT.forEach((email, i) => {
    const row = el('div', 'ml-row');
    row.dataset.target = `r${i + 1}`;
    row.dataset.files = String(email.files);
    const text = el('span', 'ml-lines');
    text.append(el('b', '', email.who), el('i', '', email.what));
    row.append(text);
    if (email.files) {
      const clip = el('span', 'ml-clip');
      clip.innerHTML = CLIP_SVG;
      row.append(clip, el('span', 'ml-add', `+${email.files}`));
    }
    list.append(row);
  });
  const ring = el('span', 'ml-ring');
  const cursor = el('span', 'ml-cursor');
  cursor.innerHTML = CURSOR_SVG;
  stage.append(list, el('span', 'ml-tally', ''), ring, cursor);
  markFlat(stage, null);
}

/** Marks an email as added in the individual-messages picture (null starts over), and counts its files. */
function markFlat(stage, target) {
  if (!target) for (const row of stage.querySelectorAll('.ml-row.added')) row.classList.remove('added');
  else {
    const row = stage.querySelector(`[data-target="${target}"]`);
    if (row) row.classList.add('added');
  }
  const rows = [...stage.querySelectorAll('.ml-row.added')];
  const files = rows.reduce((sum, row) => sum + Number(row.dataset.files || 0), 0);
  const tally = stage.querySelector('.ml-tally');
  if (tally) tally.textContent = rows.length ? `Restack: ${plural(files, 'file')} from ${plural(rows.length, 'email')}` : 'Restack: Keep adding is on';
}

/** Where the pointer goes for a target: the middle of a small box, or into a row's text. */
function howtoPoint(stage, target) {
  if (target === 'rest') return [stage.clientWidth - 40, stage.clientHeight - 30];
  const node = stage.querySelector(`[data-target="${target}"]`);
  if (!node) return [stage.clientWidth / 2, stage.clientHeight / 2];
  const s = stage.getBoundingClientRect();
  const r = node.getBoundingClientRect();
  const zoom = pageZoom(); // (positions on the page are zoomed with the text size; the stage's own aren't)
  const x = r.width < 40 * zoom ? r.left + r.width / 2 : r.left + Math.min(70 * zoom, r.width / 2);
  return [Math.round((x - s.left) / zoom), Math.round((r.top + r.height / 2 - s.top) / zoom)];
}

function moveHowtoCursor(stage, target, instant) {
  const cursor = stage.querySelector('.ml-cursor');
  const [x, y] = howtoPoint(stage, target);
  cursor.classList.toggle('instant', !!instant);
  cursor.style.transform = `translate(${x}px, ${y}px)`;
  stage.dataset.at = target;
  if (instant) void cursor.offsetWidth; // apply the jump before transitions come back
  cursor.classList.remove('instant');
}

function howtoClick(stage) {
  const ring = stage.querySelector('.ml-ring');
  const [x, y] = howtoPoint(stage, stage.dataset.at || 'rest');
  ring.style.left = `${x}px`;
  ring.style.top = `${y}px`;
  ring.classList.remove('go');
  void ring.offsetWidth; // restart the ripple
  ring.classList.add('go');
}

function markHowtoStep(index) {
  [...$('howto-steps').children].forEach((item, i) => {
    item.classList.toggle('on', i === index);
    item.classList.toggle('done', index >= 0 && i < index);
  });
}

function reducedMotion() {
  try {
    return window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  } catch (error) {
    return false;
  }
}

/** Plays the how-to over and over until the box closes (or the other Outlook is picked). */
async function playHowto(token) {
  const stage = $('howto-stage');
  const plan = HOWTO[howto.client];
  if (reducedMotion()) { // just the end result, with every step listed
    if (howto.client === 'flat') {
      for (const frame of plan.script) if (frame.mark) markFlat(stage, frame.mark);
    } else {
      stage.dataset.state = howto.client === 'classic' ? 'range' : 'all';
    }
    markHowtoStep(-1);
    return;
  }
  moveHowtoCursor(stage, 'rest', true);
  while (token === howto.token) {
    if (howto.client === 'flat') markFlat(stage, null);
    for (const frame of plan.script) {
      if (token !== howto.token) return;
      if (frame.step !== undefined) markHowtoStep(frame.step);
      if (frame.shift !== undefined) stage.classList.toggle('shift', frame.shift);
      if (frame.cursor) moveHowtoCursor(stage, frame.cursor);
      if (frame.click) howtoClick(stage);
      if (frame.state) stage.dataset.state = frame.state;
      if (frame.mark) markFlat(stage, frame.mark);
      await new Promise((resolve) => setTimeout(resolve, frame.wait));
    }
  }
}

function showHowtoFor(client) {
  howto.client = client;
  for (const [id, which] of HOWTO_TABS) {
    $(id).setAttribute('aria-selected', String(which === client));
    $(id).tabIndex = which === client ? 0 : -1;
  }
  const plan = HOWTO[client];
  const steps = $('howto-steps');
  steps.textContent = '';
  for (const text of plan.steps) steps.append(el('li', '', text));
  $('howto-note').textContent = plan.note;
  $('howto-note').hidden = !plan.note;
  const label = client === 'flat' ? '' : hostLabel();
  $('howto-you').textContent = label ? `You're using ${label}.` : '';
  $('howto-you').hidden = !label;
  buildHowtoStage(client);
  const token = ++howto.token;
  requestAnimationFrame(() => {
    if (token === howto.token) playHowto(token).catch(() => {});
  });
}

function positionHowto() {
  const zoom = pageZoom(); // the tip's position is in zoomed pixels; the box's own style isn't
  const tip = $('one-tip').getBoundingClientRect();
  const top = Math.max(8, Math.round(Math.min(tip.bottom + 6, window.innerHeight - 160)));
  $('howto').style.top = `${Math.round(top / zoom)}px`;
  $('howto').style.maxHeight = `${Math.round(Math.max(150, window.innerHeight - top - 8) / zoom)}px`;
}

function openHowto(pin) {
  clearTimeout(howto.leaveTimer);
  clearTimeout(howto.hoverTimer);
  if (pin) howto.pinned = true;
  if (!$('howto').hidden) return;
  if ($('one-tip').hidden) return;
  if (!howto.detected) {
    howto.detected = detectOutlook();
    howto.client = howto.detected || 'new';
  }
  positionHowto();
  $('howto').hidden = false;
  $('tip-info').setAttribute('aria-expanded', 'true');
  showHowtoFor(howto.client);
  log(`Tip: how-to opened (${HOWTO[howto.client].label}${pin ? ', clicked' : ', hovered'})`);
}

function closeHowto(returnFocus) {
  clearTimeout(howto.hoverTimer);
  clearTimeout(howto.leaveTimer);
  howto.token += 1; // stops the animation
  howto.pinned = false;
  if ($('howto').hidden) return;
  $('howto').hidden = true;
  $('tip-info').setAttribute('aria-expanded', 'false');
  if (returnFocus) $('tip-info').focus();
}

/** Hovering opens the how-to; moving away closes it again, unless it was clicked open. */
function laterCloseHowto() {
  if (howto.pinned) return;
  clearTimeout(howto.leaveTimer);
  howto.leaveTimer = setTimeout(() => closeHowto(false), 350);
}

function wireHowto() {
  const info = $('tip-info');
  info.addEventListener('mouseenter', () => {
    clearTimeout(howto.leaveTimer);
    clearTimeout(howto.hoverTimer);
    howto.hoverTimer = setTimeout(() => openHowto(false), 200);
  });
  info.addEventListener('mouseleave', () => {
    clearTimeout(howto.hoverTimer);
    laterCloseHowto();
  });
  info.addEventListener('click', () => {
    if (!$('howto').hidden && howto.pinned) closeHowto(true);
    else openHowto(true);
  });
  $('howto').addEventListener('mouseenter', () => clearTimeout(howto.leaveTimer));
  $('howto').addEventListener('mouseleave', laterCloseHowto);
  $('howto-close').addEventListener('click', () => closeHowto(true));
  HOWTO_TABS.forEach(([id, client], i) => {
    $(id).addEventListener('click', () => {
      howto.pinned = true;
      if (client !== howto.client) showHowtoFor(client);
    });
    $(id).addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      howto.pinned = true;
      const [nextId, next] = HOWTO_TABS[(i + (event.key === 'ArrowRight' ? 1 : HOWTO_TABS.length - 1)) % HOWTO_TABS.length];
      showHowtoFor(next);
      $(nextId).focus();
    });
  });
  document.addEventListener('pointerdown', (event) => {
    if (!$('howto').hidden && !event.target.closest('#howto, #tip-info')) closeHowto(false);
  });
  window.addEventListener('resize', () => {
    if (!$('howto').hidden) positionHowto();
  });
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
  // Keep adding: once there's a list, each email clicked adds its files instead of replacing it
  if (prefs.keepAdding && !state.fresh && state.seen.size) {
    await addFromSelection(token);
    return;
  }
  state.fresh = false;
  state.keepWaiting.clear();
  hideKeepAsk();
  closePreview();
  closeQuiz();
  hideNameForm();
  hideLoadForm();
  closeEmailMenu(false);
  renderOneEmailTip(0);
  state.fileName = '';
  state.loadNumber = '';
  state.accessorials = emptyAccessorials();
  state.multiStop = false; // a new thread starts as a single-stop load
  state.order = null;      // ...in the order its files came in
  state.built.clear();
  setStatus('');
  setSummary('Reading the selected emails…');

  const picks = await selectedEmails();
  if (token !== state.loadToken) return;
  if (!picks) {
    state.files = [];
    state.seen = new Set();
    setSummary('');
    renderKeepBar();
    showMessage('Select an email, or click a whole conversation, to see its attachments.');
    return;
  }
  const read = await readEmails(picks, token);
  if (!read) return;

  // Oldest first, so a file sent around several times is credited to whoever sent it first.
  read.found.sort(byDate);
  const files = await mergeDuplicates(read.found);
  if (token !== state.loadToken) return;
  files.forEach((file, index) => {
    file.id = index;
    file.smallLogo = looksLikeLogo(file);
  });

  state.files = files;
  state.messages = read.messages;
  state.failed = read.failed;
  state.seen = new Set(picks.map((pick) => pick.itemId).filter(Boolean));
  state.topic = stripPrefixes(read.subjects[0] || '');
  state.checked = new Set(); // nothing checked until you pick (or Select all)
  applyReceiptNames(false);
  $('title').textContent = state.topic || 'Restack';
  renderSummary();
  renderOneEmailTip(read.messages);
  renderKeepBar();

  if (!files.length) {
    showMessage(`No attachments in ${read.messages === 1 ? 'this email' : `these ${read.messages} emails`}.`);
    return;
  }
  showView('read-view');
  renderList();
  const dupes = files.reduce((sum, f) => sum + f.copies - 1, 0);
  log(`Found ${plural(files.length, 'file')}${dupes ? `, ${plural(dupes, 'duplicate')} merged` : ''}${read.failed ? `, ${read.failed} unreadable` : ''}`);
  logScans(files);
  fillThumbs(token).catch((error) => log(`Pictures: stopped (${error.message})`));
  checkRateCons(token).catch((error) => log(`Rate con check: stopped (${error.message})`));
}

function byDate(a, b) {
  return (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0);
}

function logScans(files) {
  for (const file of files) {
    const scans = file.kind === 'PDF' && file.bytes ? scanFormats(file.bytes) : '';
    if (scans) log(`${file.name}: has ${scans} images`);
  }
}

/**
 * The selected emails ([{ itemId, subject, itemMode, hasAttachment }]), or the open one where
 * Outlook can't say what's selected ([{ itemId, subject, open }]), or null when nothing is.
 */
async function selectedEmails() {
  const mailbox = Office.context.mailbox;
  let selected = [];
  if (supports('1.13') && typeof mailbox.getSelectedItemsAsync === 'function') {
    try {
      selected = (await call((done) => mailbox.getSelectedItemsAsync(done))) || [];
    } catch (error) {
      log(`Couldn't read the selection: ${error.message}`);
    }
  }
  if (selected.length && supports('1.15') && typeof mailbox.loadItemByIdAsync === 'function') {
    return selected.map((pick) => ({ itemId: pick.itemId, subject: pick.subject, itemMode: pick.itemMode, hasAttachment: pick.hasAttachment, multi: true }));
  }
  if (mailbox.item) {
    if (selected.length > 1) log('Multi-select not available here: reading the open email');
    return [{ itemId: String(mailbox.item.itemId || ''), subject: mailbox.item.subject || '', open: mailbox.item }];
  }
  return null;
}

/** Reads the emails' attachments. Returns { found, subjects, failed, messages }, or null when a newer load took over. */
async function readEmails(picks, token) {
  const mailbox = Office.context.mailbox;
  const found = [];
  const subjects = [];
  let failed = 0;
  let messages = 0;
  for (let i = 0; i < picks.length; i++) {
    if (token !== state.loadToken) return null;
    const pick = picks[i];
    messages += 1;
    if (pick.subject) subjects.push(pick.subject);
    if (pick.open) {
      failed += await readAttachments(pick.open, found, token, 'Open email');
      continue;
    }
    if (String(pick.itemMode || '').toLowerCase() === 'compose') {
      log(`Skipped a draft: ${pick.subject || '(no subject)'}`);
      continue;
    }
    // No skipping on hasAttachment: classic Outlook reports false when the only
    // attachments are pictures pasted into the body.
    setSummary(`Reading email ${i + 1} of ${picks.length}…`);
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
  if (token !== state.loadToken) return null;
  log(picks.some((pick) => pick.multi) ? `Read ${plural(picks.length, 'selected email')} (multi-select)` : 'Read the open email');
  return { found, subjects, failed, messages };
}

// ---- Keep adding: for people who don't show their mail as conversations. With the panel pinned,
// each email clicked adds its files to the list, until Start over (or another load's email).

/** Whether an email goes with the list's: the same load number, or else the same subject. */
function sameThread(subject) {
  if (!state.topic) return true;
  const ours = guessLoadNumber(state.topic);
  const theirs = guessLoadNumber(subject);
  if (ours && theirs) return ours.toLowerCase() === theirs.toLowerCase();
  const plain = (text) => stripPrefixes(text).toLowerCase().replace(/\s+/g, ' ').trim();
  return plain(subject) === plain(state.topic);
}

/** Keep adding: notes the emails selected right now, so a quick run of clicks adds every one of them. */
async function noteSelection() {
  if (!prefs.keepAdding || state.fresh || !state.seen.size) return;
  const picks = await selectedEmails();
  for (const pick of picks || []) {
    if (pick.itemId && !state.seen.has(pick.itemId) && !state.keepWaiting.has(pick.itemId)) state.keepWaiting.set(pick.itemId, pick);
  }
}

/** Starts the pictures and the rate con check again for files still without them (a newer load stopped them). */
function resumeChecks(token) {
  fillThumbs(token).catch((error) => log(`Pictures: stopped (${error.message})`));
  checkRateCons(token).catch((error) => log(`Rate con check: stopped (${error.message})`));
}

/** Keep adding: the clicked emails' files go into the list (asking first about one that looks like another load). */
async function addFromSelection(token, force = false) {
  const picks = (await selectedEmails()) || [];
  if (token !== state.loadToken) return;
  const current = new Set(picks.map((pick) => pick.itemId).filter(Boolean));
  for (const pick of picks) if (pick.itemId && !state.seen.has(pick.itemId)) state.keepWaiting.set(pick.itemId, pick);
  const waiting = [...state.keepWaiting.values(), ...picks.filter((pick) => !pick.itemId)];
  const fits = (pick) => force || sameThread(pick.subject || '');
  const fresh = waiting.filter(fits);
  const other = waiting.filter((pick) => !fits(pick));
  for (const pick of other) if (!current.has(pick.itemId)) state.keepWaiting.delete(pick.itemId); // clicked past it: skipped
  const asking = other.filter((pick) => current.has(pick.itemId));
  if (asking.length) askKeep(asking[0].subject || '');
  else hideKeepAsk();
  if (!fresh.length) {
    if (!asking.length && picks.length) keepFlash(picks.length > 1 ? 'Those emails are in the list already.' : 'That email is in the list already.');
    resumeChecks(token); // this load stopped the ones that were running
    return;
  }
  keepFlash(`Reading ${fresh.length > 1 ? `${fresh.length} emails` : 'the email'}…`);
  const read = await readEmails(fresh, token);
  if (!read) return; // a newer click took over: these stay waiting, and it adds them too
  const added = await mergeInto(read.found, token);
  if (added === null) return;
  for (const pick of fresh) {
    if (!pick.itemId) continue;
    state.seen.add(pick.itemId);
    state.keepWaiting.delete(pick.itemId);
  }
  state.messages += read.messages;
  state.failed += read.failed;
  if (!state.topic) {
    state.topic = stripPrefixes(read.subjects[0] || '');
    $('title').textContent = state.topic || 'Restack';
  }
  applyReceiptNames(false);
  renderSummary();
  renderKeepBar();
  const from = added.length && added[0].from ? ` from ${added[0].from}` : '';
  keepFlash(added.length ? `Added ${plural(added.length, 'file')}${from}.` : `Nothing new in ${read.messages > 1 ? 'those emails' : 'that email'}.`);
  log(`Keep adding: ${plural(read.messages, 'email')}, ${plural(added.length, 'new file')}`);
  if (!state.files.length) {
    showMessage(`No attachments in ${state.messages === 1 ? 'this email' : `these ${state.messages} emails`}.`);
    return;
  }
  showView('read-view');
  renderList();
  updateSelectionUi();
  for (const file of added) flashRow(file);
  logScans(added);
  fillThumbs(token).catch((error) => log(`Pictures: stopped (${error.message})`));
  checkRateCons(token).catch((error) => log(`Rate con check: stopped (${error.message})`));
}

/**
 * Adds newly read files to the list. A copy of a file that's already there just counts as a copy;
 * the rest go in by date, oldest first (or at the end, once you've put the list in your own order).
 * Returns the files added, or null when a newer load took over.
 */
async function mergeInto(found, token) {
  const byKey = new Map();
  for (const file of state.files) {
    if (!file.fp) file.fp = file.bytes ? await fingerprint(file.bytes) : `${file.kind}:${file.name}`;
    byKey.set(file.fp, file);
  }
  found.sort(byDate);
  const added = [];
  for (const file of found) {
    file.fp = file.bytes ? await fingerprint(file.bytes) : `${file.kind}:${file.name}`;
    const first = byKey.get(file.fp);
    if (first) {
      first.copies += 1;
      first.inline = first.inline && file.inline; // a real attachment anywhere wins
      continue;
    }
    byKey.set(file.fp, file);
    file.smallLogo = looksLikeLogo(file);
    added.push(file);
  }
  if (token !== state.loadToken) return null;
  if (!added.length) return added;
  const before = new Map(state.files.map((file) => [file, file.id]));
  const previewed = state.files[state.previewIndex];
  const all = [...state.files, ...added].sort(byDate);
  all.forEach((file, index) => { file.id = index; });
  const moved = new Map([...before].map(([file, old]) => [old, file.id]));
  state.checked = new Set([...state.checked].map((id) => moved.get(id)).filter((id) => id !== undefined));
  if (state.order) state.order = [...state.order.map((id) => moved.get(id)).filter((id) => id !== undefined), ...added.map((file) => file.id)];
  state.files = all;
  state.previewIndex = previewed ? all.indexOf(previewed) : -1;
  state.built.clear();
  return added;
}

/** The Keep adding bar under the header, while it's on. */
function renderKeepBar() {
  const bar = $('keep-bar');
  if (!bar) return;
  const on = prefs.keepAdding && readUiWired;
  bar.hidden = !on;
  if (!on) {
    hideKeepAsk();
    return;
  }
  $('keep-count').textContent = state.seen.size ? `· ${plural(state.seen.size, 'email')}` : '';
}

/** A moment's note in the Keep adding bar (what was just added), then back to the how-to line. */
function keepFlash(text) {
  const line = $('keep-text');
  if (!line) return;
  line.textContent = text;
  clearTimeout(keepFlash.timer);
  keepFlash.timer = setTimeout(() => { line.textContent = word('keepAddingNote'); }, 5000);
}

/** The clicked email looks like another load: add it anyway, or start over with it? */
function askKeep(subject) {
  $('keep-ask-text').textContent = `“${stripPrefixes(subject) || 'This email'}” looks like a different load from “${state.topic}”.`;
  $('keep-ask').hidden = false;
  log('Keep adding: the clicked email looks like another load, asked first');
}

function hideKeepAsk() {
  const box = $('keep-ask');
  if (box) box.hidden = true;
}

/** "Add its files": the asked-about email goes into the list after all. */
function keepAddAnyway() {
  if (state.busy) {
    keepFlash('Wait for the PDF to finish, then click Add its files again.');
    return Promise.resolve();
  }
  hideKeepAsk();
  const token = ++state.loadToken;
  const run = loadChain.then(() => (token === state.loadToken ? addFromSelection(token, true) : undefined));
  loadChain = run.catch(() => {});
  return run;
}

/** "Start over": a fresh list from the selected email(s), as if the panel had just opened on them. */
function keepStartOver() {
  hideKeepAsk();
  state.keepWaiting.clear();
  state.fresh = true;
  log('Keep adding: start over');
  return loadThread();
}

function setKeepAdding(on) {
  savePref('keepAdding', on);
  markSwitch('set-keep-adding', on);
  state.keepWaiting.clear();
  hideKeepAsk();
  renderKeepBar();
  renderOneEmailTip(state.messages);
  if (on) keepFlash(word('keepAddingNote'));
  log(`Keep adding: ${on ? 'on' : 'off'}`);
}

// ---- receipts by file name (a switch in Settings; the team admin sets the words)

/** Whether a file's name has one of the team's receipt words in it. */
function namedLikeReceipt(file) {
  const name = String(file.name || '').toLowerCase();
  const pod = name.replace(/delivery[\s_-]*receipt/g, ' '); // "Delivery Receipt" is a POD: its "receipt" doesn't count
  return receiptWords().some((w) => w && pod.includes(w));
}

/**
 * With Tag receipts by file name on, files named like receipts start out tagged Receipt, and lose the
 * tag again when it's turned off. A tag set by hand stays as it was set.
 */
function applyReceiptNames(render = true) {
  let changed = false;
  for (const file of state.files) {
    if (file.lumperByHand || file.pageLumper || !canCombine(file)) continue;
    const want = prefs.tagReceipts && namedLikeReceipt(file);
    if (!!file.lumperByName === want) continue;
    file.lumperByName = want;
    file.lumper = want;
    changed = true;
  }
  if (!changed) return;
  state.built.clear();
  normalizeOrder();
  if (render && readUiWired && !$('read-view').hidden) renderList();
}

function setTagReceipts(on) {
  savePref('tagReceipts', on);
  markSwitch('set-tag-receipts', on);
  log(`Tag receipts by file name: ${on ? `on (${receiptWords().join(', ') || 'no words set'})` : 'off'}`);
  applyReceiptNames();
}

// --------------------------------------------------------------------------- picking pages

/**
 * 0-based pages of a PDF that go into the combined file, in the order they go in (as moved in the
 * preview), or null for "every page, as it came".
 */
function includedPages(file) {
  if (file.kind !== 'PDF') return null;
  const order = movedPages(file);
  if (!file.pageChoice) return order;
  return order ? order.filter((p) => file.pageChoice.has(p)) : [...file.pageChoice].sort((a, b) => a - b);
}

// ---- page order (moved in the preview, like Rotate: the attachment itself isn't changed)

/** How many pages a PDF or fax (TIFF) has, once it's known. */
function pageTotal(file) {
  return file.kind === 'TIFF' ? file.tiffPages || 0 : file.pages || 0;
}

/** The order a file's pages go into the PDF (0-based page numbers), as moved in the preview. */
function pageSeq(file) {
  const total = pageTotal(file);
  const order = file.pageOrder;
  if (order && order.length === total) return order.slice();
  return Array.from({ length: total }, (_, i) => i);
}

/** The pages in their new order, or null when they haven't been moved. */
function movedPages(file) {
  const order = file.pageOrder;
  return order && order.length === pageTotal(file) && order.some((p, i) => p !== i) ? order.slice() : null;
}

/** The page that goes in first (0-based): the list's picture shows it. */
function firstPage(file) {
  const order = movedPages(file);
  return order ? order[0] : 0;
}

/** Draws the list's picture again (after the first page changed). */
function refreshThumb(file) {
  makeThumb(file)
    .then((thumb) => { file.thumb = thumb; })
    .catch(() => { file.thumb = null; })
    .then(() => placeThumb(file));
}

/** "2, 1, 3" (the pages going in, in their new order), or '' when they haven't been moved. */
function orderText(file) {
  const order = movedPages(file);
  if (!order) return '';
  const going = file.kind === 'PDF' && file.pageChoice ? order.filter((p) => file.pageChoice.has(p)) : order;
  const shown = going.slice(0, 12).map((p) => p + 1).join(', ');
  return going.length > 12 ? `${shown}…` : shown;
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
  const order = orderText(file);
  const moved = order ? word('pageOrder', { pages: order }) : '';
  if (!file.pageChoice) return moved || null;
  const picked = `${pagesLabel([...file.pageChoice])} of ${file.pages} picked`;
  return moved ? `${picked} · ${moved}` : picked;
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

/**
 * How one checked file is divided up: [{ stop, lumper, pages }] (pages null = all its picked
 * pages). Stops count only with Multi-stop on. Lumper receipt pages are a part of their own, so
 * they can go after the PODs (or into a file of their own).
 */
function partsOfFile(file) {
  const stopOf = (p) => (state.multiStop ? ((file.pageStops && file.pageStops.get(p)) || file.stop || 1) : 1);
  const byPage = file.kind === 'PDF' && ((state.multiStop && file.pageStops && file.pageStops.size) || (file.pageLumper && file.pageLumper.size));
  const pages = byPage ? allIncludedPages(file) : null;
  if (!pages) return [{ stop: state.multiStop ? (file.stop || 1) : 1, lumper: !!file.lumper, pages: includedPages(file) }];
  const parts = new Map();
  for (const p of pages) {
    const stop = stopOf(p);
    const lumper = pageLumperOf(file, p);
    const key = `${stop}|${lumper}`;
    if (!parts.has(key)) parts.set(key, { stop, lumper, pages: [] });
    parts.get(key).pages.push(p);
  }
  return [...parts.values()];
}

/**
 * The PDFs to make, before building them: [{ stop, lumper, entries: [{ file, only }] }], PODs
 * first, then lumper receipts. splitStops: one PDF per stop (stop 0 = all stops together, in stop
 * order). splitLumpers: lumper receipts in their own PDF(s); otherwise they go after the PODs in
 * the same file (after every stop, when the stops share a file). Within that, files go in the
 * order of the list, which you can change.
 */
function fileGroups(splitStops, splitLumpers, noReceipts = false) {
  const groups = new Map();
  checkedFiles().forEach((file, index) => {
    for (const part of partsOfFile(file)) {
      if (noReceipts && part.lumper) continue; // the POD email, when the receipts go in the accessorial email
      const stop = splitStops ? part.stop : 0;
      const own = splitLumpers && part.lumper;
      const key = `${own ? 'L' : 'P'}${stop}`;
      if (!groups.has(key)) groups.set(key, { stop, lumper: own, items: [] });
      groups.get(key).items.push({ file, only: part.pages, receipt: part.lumper, order: part.stop, index });
    }
  });
  return [...groups.values()]
    .sort((a, b) => (Number(a.lumper) - Number(b.lumper)) || (a.stop - b.stop))
    .map((group) => ({
      stop: group.stop,
      lumper: group.lumper,
      entries: group.items
        .sort((x, y) => (Number(x.receipt) - Number(y.receipt)) || (x.order - y.order) || (x.index - y.index))
        .map(({ file, only }) => ({ file, only })),
    }));
}

/** "Order in 3645493.pdf: pod1.pdf, pod2.pdf, scan.pdf (page 2)" for the test report. */
function logOrder(name, entries) {
  if (entries.length < 2) return;
  const parts = entries.map(({ file, only }) => (only && file.pages && only.length < file.pages ? `${file.name} (${pagesLabel(only).toLowerCase()})` : file.name));
  log(`Order in ${name}: ${parts.join(', ')}`);
}

/** What an email template does with stops and lumper receipts (t null: Save PDF and Blank email). */
function layoutOf(t) {
  return {
    splitStops: state.multiStop && (!t || t.stopFiles !== false),
    splitLumpers: !!(t && t.lumperSeparate),
  };
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
  return [...map.entries()].sort((a, b) => a[0] - b[0]).map(([stop, pages]) => `${word('stop', { n: stop })}: ${pagesLabel(pages).toLowerCase()}`).join(' · ');
}

/** A stop picker (Stop 1…n). `mixed` adds a "By page" choice for a file split across stops. */
function stopPicker(value, label, mixed, onPick) {
  const select = document.createElement('select');
  select.className = 'stop-pick';
  select.setAttribute('aria-label', label);
  if (mixed) {
    const option = new Option(word('byPage'), '', true, true);
    option.disabled = true;
    select.append(option);
  }
  for (let stop = 1; stop <= stopChoices(); stop++) select.append(new Option(word('stop', { n: stop }), String(stop), false, !mixed && stop === value));
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
  refreshList();
  if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
}

/** Picking a stop for one page in the preview (for a PDF that holds more than one stop's POD). */
function setPageStop(file, page, stop) {
  if (!file.pageStops) file.pageStops = new Map();
  file.pageStops.set(page, stop);
  normalizePageStops(file);
  state.built.clear();
  refreshList();
  syncPreviewPages();
}

// --------------------------------------------------------------------------- lumper receipts: which files (or pages) they are
// Every checked file has a "Receipt" tag (WORDS.receipt). Tagged files and pages go after the PODs, or
// into a file of their own for templates that want that.

function pageLumperOf(file, page) {
  return file.pageLumper && file.pageLumper.has(page) ? file.pageLumper.get(page) : !!file.lumper;
}

/** When every picked page of a file is (or isn't) a lumper receipt, that's simply the file's tag again. */
function normalizePageLumper(file) {
  if (!file.pageLumper || !file.pageLumper.size) {
    file.pageLumper = null;
    return;
  }
  const pages = allIncludedPages(file);
  if (!pages) return;
  const values = new Set(pages.map((p) => pageLumperOf(file, p)));
  if (values.size <= 1) {
    if (values.size) file.lumper = [...values][0];
    file.pageLumper = null;
    return;
  }
  for (const [page, value] of [...file.pageLumper]) if (value === !!file.lumper) file.pageLumper.delete(page);
}

/** "Receipt: page 3" for a file that's partly a lumper receipt, else null. */
function lumperNote(file) {
  if (!state.checked.has(file.id) || !file.pageLumper || !file.pageLumper.size) return null;
  const pages = (allIncludedPages(file) || []).filter((p) => pageLumperOf(file, p));
  return pages.length ? word('receiptPages', { pages: pagesLabel(pages).toLowerCase() }) : null;
}

/** A "Receipt" tag: a checkbox in a small pill. */
function lumperChip(on, mixed, label, onToggle) {
  const chip = el('label', 'tag-chip');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.checked = on && !mixed;
  box.indeterminate = mixed;
  box.disabled = state.busy;
  box.setAttribute('aria-label', label);
  box.addEventListener('change', () => onToggle(box.checked));
  chip.classList.toggle('on', on || mixed);
  chip.append(box, el('span', '', word('receipt')));
  return chip;
}

/** Tagging a whole file as a lumper receipt (or not). This also undoes any page-by-page tags. */
function setFileLumper(file, on) {
  file.lumper = on;
  file.lumperByHand = true; // Tag receipts by file name leaves it as you set it
  file.pageLumper = null;
  state.built.clear();
  log(`Receipt: ${file.name} ${on ? 'yes' : 'no'}`);
  const moved = normalizeOrder(); // a receipt moves below the PODs
  if (moved) glide(file, refreshList); // and slides there, so it's easy to see where it went
  else refreshList();
  if (moved) flashRow(file);
  if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
}

/** Tagging one page in the preview (a PDF with the POD and the lumper receipt in it). */
function setPageLumper(file, page, on) {
  file.lumperByHand = true;
  if (!file.pageLumper) file.pageLumper = new Map();
  file.pageLumper.set(page, on);
  normalizePageLumper(file);
  state.built.clear();
  normalizeOrder(); // every page tagged makes it a whole-file receipt
  refreshList();
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
  if (String(file.name).toLowerCase() === SENT_WITH.file) return true; // Restack's own sent with restack line
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
    file.fp = key;
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

/** Every file, in the list's order: as they came (oldest first) until you move one. */
function orderedFiles() {
  if (!state.order) return state.files;
  const rank = new Map(state.order.map((id, i) => [id, i]));
  return [...state.files].sort((a, b) => (rank.has(a.id) ? rank.get(a.id) : 1e9) - (rank.has(b.id) ? rank.get(b.id) : 1e9));
}

function visibleFiles() {
  return orderedFiles().filter((f) => state.showLogos || !f.smallLogo);
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
  if (drag) return; // redrawn when the drag ends
  const spot = focusedSpot();
  state.places = orderPlaces();
  const list = $('file-list');
  list.dataset.view = prefs.view;
  $('details-head').hidden = prefs.view !== 'details';
  list.textContent = '';
  for (const file of visibleFiles()) list.append(fileRow(file));
  const logos = state.files.filter((f) => f.smallLogo).length;
  $('logos-toggle').hidden = logos === 0;
  $('logo-count').textContent = String(logos);
  $('show-logos').checked = state.showLogos;
  updateSelectionUi();
  restoreFocus(spot);
}

// Redrawing the list keeps focus on the same control of the same file
const ROW_CONTROLS = ['.order-grip', '.order-up', '.order-down', '.stop-pick', '.tag-chip input', '.file-label > input', '.eye'];

function focusedSpot() {
  const node = document.activeElement;
  const row = node && node.closest ? node.closest('#file-list .file') : null;
  if (!row) return null;
  return { id: row.dataset.id, control: ROW_CONTROLS.find((s) => node.matches(s)) || '.file-label > input' };
}

function restoreFocus(spot) {
  if (!spot) return;
  const row = document.querySelector(`#file-list .file[data-id="${spot.id}"]`);
  if (!row) return;
  const target = [spot.control, '.order-grip', '.file-label > input'].map((s) => row.querySelector(s)).find((n) => n && !n.disabled);
  if (target) target.focus(gliding ? { preventScroll: true } : undefined); // a sliding file: glide() scrolls along smoothly
}

// --------------------------------------------------------------------------- the order files go into the PDF

/** 1st, 2nd, 3rd, 4th… */
function ordinal(n) {
  const tens = n % 100;
  const suffix = tens >= 11 && tens <= 13 ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th');
  return `${n}${suffix}`;
}

/**
 * Each checked file's place in its PDF (1, 2, 3…, counted per stop with Multi-stop on), or none
 * when fewer than two files are checked. The list keeps lumper receipts below the PODs, so these
 * are simply the list's order.
 */
function orderPlaces() {
  const places = new Map();
  const checked = checkedFiles();
  if (checked.length < 2) return places;
  const counts = new Map();
  for (const file of checked) {
    const key = state.multiStop ? (file.stop || 1) : 0;
    counts.set(key, (counts.get(key) || 0) + 1);
    places.set(file.id, counts.get(key));
  }
  return places;
}

/** Puts the checked files in this order, in the places checked files have now (the others stay put). */
function setCheckedOrder(sorted) {
  const ids = new Set(sorted.map((f) => f.id));
  let k = 0;
  state.order = orderedFiles().map((f) => (ids.has(f.id) ? sorted[k++].id : f.id));
  state.built.clear();
}

/** Lumper receipts (whole files) go below the other checked files, keeping their own order. */
function normalizeOrder() {
  const checked = checkedFiles();
  const sorted = [...checked.filter((f) => !f.lumper), ...checked.filter((f) => f.lumper)];
  if (sorted.every((f, i) => f === checked[i])) return false;
  setCheckedOrder(sorted);
  return true;
}

/** The checked file a step up (-1) or down (1) that `file` can swap with: same stop, both PODs or both receipts. */
function neighbor(file, step) {
  const checked = checkedFiles();
  const same = (f) => !!f.lumper === !!file.lumper && (!state.multiStop || (f.stop || 1) === (file.stop || 1));
  for (let i = checked.indexOf(file) + step; i >= 0 && i < checked.length; i += step) {
    if (same(checked[i])) return checked[i];
  }
  return null;
}

/** The up and down arrows (and the arrow keys on a file's handle). */
function moveFile(file, step) {
  const target = neighbor(file, step);
  if (!target || state.busy) return;
  const others = checkedFiles().filter((f) => f !== file);
  others.splice(others.indexOf(target) + (step > 0 ? 1 : 0), 0, file);
  setCheckedOrder(others);
  log(`Order: ${file.name} ${step < 0 ? 'up' : 'down'}`);
  glide(file, renderList);
  flashRow(file);
  if (!$('name-form').hidden) updateNameHint();
}

/** A moment's highlight on a file that just moved, so it's easy to see where it went. */
function flashRow(file) {
  const row = document.querySelector(`#file-list .file[data-id="${file.id}"]`);
  if (!row) return;
  row.classList.remove('moved');
  void row.offsetWidth; // restart the highlight
  row.classList.add('moved');
}

// ---- files sliding to their new place (the up / down arrows, a Receipt tag, a drop)
// The list is drawn again from scratch, so each file's spot is measured before and after, and every file
// that changed places slides from its old spot to its new one: the moved file on top and highlighted, the
// files it passes going the other way. Nothing slides for people who've asked their computer for less motion.

let gliding = false; // while a move is being drawn: focus goes back to the file without jumping the list

function motionOk() {
  return typeof Element.prototype.animate === 'function'
    && !(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
}

/** Where each file is in the list (measured from the list's corner, so scrolling doesn't count as moving). */
function fileSpots() {
  const list = $('file-list');
  const base = list.getBoundingClientRect();
  const spots = new Map();
  for (const row of list.querySelectorAll('.file')) {
    const r = row.getBoundingClientRect();
    spots.set(row.dataset.id, { x: r.left - base.left, y: r.top - base.top });
  }
  return spots;
}

/** Draws the list again with `redraw`, sliding the files that changed places. `mover` is the file that was moved. */
function glide(mover, redraw) {
  if (!motionOk() || $('read-view').hidden) {
    redraw();
    return;
  }
  const before = fileSpots(); // (mid-slide, that's where a file is on screen, so a quick second click carries on from there)
  gliding = true;
  try {
    redraw();
  } finally {
    gliding = false;
  }
  const list = $('file-list');
  const base = list.getBoundingClientRect();
  const zoom = pageZoom(); // spots are in the page's zoomed pixels; a slide is in the file's own
  let lead = null;
  let leadBox = null;
  const slides = [];
  for (const row of list.querySelectorAll('.file')) {
    const from = before.get(row.dataset.id);
    if (!from) continue;
    const box = row.getBoundingClientRect();
    const dx = (from.x - (box.left - base.left)) / zoom;
    const dy = (from.y - (box.top - base.top)) / zoom;
    if (Math.abs(dx) < 1 && Math.abs(dy) < 1) continue;
    if (mover && row.dataset.id === String(mover.id)) {
      lead = row;
      leadBox = box;
    }
    slides.push({ row, dx, dy });
  }
  for (const { row, dx, dy } of slides) { // (measured first, then started: a slide doesn't move the files after it)
    // Slow enough to follow: 0.56 s for the next spot, up to 0.9 s for a long way, moving evenly rather than all at the start
    const ms = Math.round(Math.min(900, Math.max(560, 480 + Math.hypot(dx, dy) * 0.5)));
    const run = row.animate([{ transform: `translate(${dx}px, ${dy}px)` }, { transform: 'translate(0, 0)' }],
      { id: 'slide', duration: ms, easing: 'cubic-bezier(.45, 0, .25, 1)' });
    if (row === lead) {
      row.classList.add('gliding'); // rides above the files it passes
      run.addEventListener('finish', () => row.classList.remove('gliding'));
    }
  }
  // The list scrolls along to the moved file's new spot, if that's out of sight. (Not by focus: a click
  // doesn't focus a button in Outlook on a Mac. Every move is made in the list, so not with the preview open.)
  if (lead && $('preview').hidden) {
    const view = $('read-view');
    const top = view.querySelector('.list-tools').getBoundingClientRect().bottom + 6;
    const bottom = view.getBoundingClientRect().bottom - 6;
    let by = 0;
    if (leadBox.bottom > bottom) by = leadBox.bottom - bottom;
    if (leadBox.top - by < top) by = leadBox.top - top; // (taller than the space: its top)
    if (Math.abs(by) >= 1) view.scrollBy({ top: by / zoom, behavior: 'smooth' });
  }
}

/** Puts every file that's still sliding straight into its place (picking one up to drag: it follows the pointer, and the others are where they look). */
function settleSlides() {
  if (typeof Element.prototype.getAnimations !== 'function') return; // (no slides there either)
  for (const row of document.querySelectorAll('#file-list .file')) {
    for (const run of row.getAnimations()) if (run.id === 'slide') run.finish();
  }
}

/** A checked file's place in the PDF, its handle for dragging, and up / down arrows. */
function orderControl(file, place) {
  const wrap = el('span', 'order-ctl');
  const where = state.multiStop ? `in stop ${file.stop || 1}'s PDF` : 'in the PDF';
  const grip = el('button', 'order-grip');
  grip.type = 'button';
  grip.innerHTML = GRIP_SVG;
  grip.append(el('span', 'order-n', String(place)));
  grip.title = `${ordinal(place)} ${where}. Drag to move it.`;
  grip.setAttribute('aria-label', `${file.name}, ${ordinal(place)} ${where}. Drag to move it, or press the up or down arrow key.`);
  grip.disabled = state.busy;
  grip.addEventListener('pointerdown', (event) => startDrag(event, file, grip));
  grip.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    event.preventDefault();
    moveFile(file, event.key === 'ArrowUp' ? -1 : 1);
  });
  const arrow = (cls, svg, step, label) => {
    const button = el('button', `order-step ${cls}`);
    button.type = 'button';
    button.innerHTML = svg;
    button.title = label;
    button.setAttribute('aria-label', `${label}: ${file.name}`);
    button.dataset.edge = neighbor(file, step) ? 'no' : 'yes'; // nothing to swap with that way
    button.disabled = state.busy || button.dataset.edge === 'yes';
    button.addEventListener('click', () => moveFile(file, step));
    return button;
  };
  wrap.append(grip, arrow('order-up', UP_SVG, -1, 'Move up'), arrow('order-down', DOWN_SVG, 1, 'Move down'));
  return wrap;
}

// ---- dragging a file by its handle

let drag = null; // { file, grip, row, rows, startY, scroll0, moved, index } while a handle is held

function startDrag(event, file, grip) {
  if (event.button !== 0 || state.busy || drag) return;
  event.preventDefault(); // no text selection while dragging
  grip.focus();
  settleSlides();
  const row = grip.closest('.file');
  drag = {
    file, grip, row,
    rows: [...document.querySelectorAll('#file-list .file.ordered')].filter((r) => r !== row),
    startX: event.clientX,
    startY: event.clientY,
    scroll0: $('read-view').scrollTop,
    moved: false,
    index: null,
    tiles: $('file-list').dataset.view === 'tiles', // cards side by side: across as well as down
  };
  try { grip.setPointerCapture(event.pointerId); } catch (error) { /* moves still arrive */ }
  grip.addEventListener('pointermove', onDragMove);
  grip.addEventListener('pointerup', endDrag);
  grip.addEventListener('pointercancel', endDrag);
}

function onDragMove(event) {
  if (!drag) return;
  const view = $('read-view');
  const box = view.getBoundingClientRect();
  if (event.clientY < box.top + 30) view.scrollTop -= 12; // keep going past the visible part of a long list
  else if (event.clientY > box.bottom - 30) view.scrollTop += 12;
  // (The pointer and the rows' positions are in the page's zoomed pixels; a row's own position isn't.)
  const zoom = pageZoom();
  const dy = (event.clientY - drag.startY) / zoom + (view.scrollTop - drag.scroll0);
  const dx = drag.tiles ? (event.clientX - drag.startX) / zoom : 0;
  if (!drag.moved && Math.abs(dy) < 4 && Math.abs(dx) < 4) return;
  if (!drag.moved) {
    drag.moved = true;
    drag.row.classList.add('dragging');
  }
  drag.row.style.transform = `translate(${dx}px, ${dy}px)`;
  // It lands before the first file whose middle is below the pointer (in Tiles: past the pointer, reading across)
  let index = drag.rows.length;
  for (let i = 0; i < drag.rows.length; i++) {
    const r = drag.rows[i].getBoundingClientRect();
    const before = drag.tiles
      ? event.clientY < r.top || (event.clientY <= r.bottom && event.clientX < r.left + r.width / 2)
      : event.clientY < r.top + r.height / 2;
    if (before) {
      index = i;
      break;
    }
  }
  drag.index = index;
  const line = $('drop-line');
  const wrap = line.parentElement.getBoundingClientRect();
  if (drag.tiles) { // a line down the side of the card it lands before
    const r = (drag.rows[index] || drag.rows[drag.rows.length - 1]).getBoundingClientRect();
    const x = index < drag.rows.length ? r.left - 6 : r.right + 3;
    line.classList.add('vertical');
    line.style.left = `${Math.round((x - wrap.left) / zoom)}px`;
    line.style.top = `${Math.round((r.top - wrap.top) / zoom)}px`;
    line.style.height = `${Math.round(r.height / zoom)}px`;
  } else {
    line.classList.remove('vertical');
    line.style.left = '';
    line.style.height = '';
    const edge = index < drag.rows.length ? drag.rows[index].getBoundingClientRect().top : drag.rows[drag.rows.length - 1].getBoundingClientRect().bottom;
    line.style.top = `${Math.round((edge - wrap.top) / zoom) - 1}px`;
  }
  line.hidden = false;
}

function endDrag(event) {
  if (!drag) return;
  const { file, grip, row, rows, moved, index } = drag;
  drag = null;
  grip.removeEventListener('pointermove', onDragMove);
  grip.removeEventListener('pointerup', endDrag);
  grip.removeEventListener('pointercancel', endDrag);
  try { grip.releasePointerCapture(event.pointerId); } catch (error) { /* already released */ }
  $('drop-line').hidden = true;
  const redraw = () => {
    row.classList.remove('dragging');
    row.style.transform = '';
    if (moved && index !== null && event.type === 'pointerup') {
      const others = rows.map((r) => state.files.find((f) => f.id === Number(r.dataset.id))).filter(Boolean);
      others.splice(index, 0, file);
      setCheckedOrder(others);
      normalizeOrder(); // receipts stay below the PODs
      log(`Order: moved ${file.name}`);
      if (!$('name-form').hidden) updateNameHint();
    }
    renderList();
    const handle = document.querySelector(`#file-list .file[data-id="${file.id}"] .order-grip`);
    if (handle) handle.focus(gliding ? { preventScroll: true } : undefined);
  };
  if (moved) glide(file, redraw); // it settles from where it was let go (or slides back, if the drag was called off)
  else redraw();
  if (moved) flashRow(file);
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
  if (note) text.append(noteEl('file-note picked', note));
  const lumperLine = lumperNote(file);
  if (lumperLine) text.append(noteEl('file-note lumper', lumperLine));
  const rcLine = rcNote(file);
  if (rcLine) text.append(noteEl(`file-note rc${rcLine.going ? '' : ' out'}`, rcLine.text)); // grey once it's not going in

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
  // The Details view's columns (the other views hide them)
  const cell = (cls, value) => {
    const node = el('span', `cell ${cls}`, value);
    if (value) node.title = value;
    return node;
  };
  label.append(box, thumb, text, cell('c-from', file.from || ''), cell('c-date', shortDate(file.date)), cell('c-size', humanSize(file.size)));

  const preview = document.createElement('button');
  preview.type = 'button';
  preview.className = 'icon-btn eye';
  preview.title = 'Preview';
  preview.setAttribute('aria-label', `Preview ${file.name}`);
  preview.innerHTML = EYE_SVG;
  preview.addEventListener('click', () => openPreview(file.id));
  row.append(label);
  // Its place in the PDF, which stop it is, and whether it's a lumper receipt (outside the label,
  // so using them doesn't tick the file's own box)
  if (canCombine(file) && state.checked.has(file.id)) {
    const tags = el('div', 'row-tags');
    const place = state.places.get(file.id);
    if (place) {
      row.classList.add('ordered');
      tags.append(orderControl(file, place));
    }
    if (state.multiStop) {
      const mixed = !!(file.pageStops && file.pageStops.size);
      tags.append(stopPicker(file.stop || 1, `Stop for ${file.name}`, mixed, (stop) => setFileStop(file, stop)));
    }
    const mixed = !!(file.pageLumper && file.pageLumper.size);
    tags.append(lumperChip(!!file.lumper, mixed, `${file.name}: ${word('receipt')}`, (on) => setFileLumper(file, on)));
    row.append(tags);
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
    ? `${word('selected', { n: count, total: selectable.length })}${stops ? ` · ${plural(stops, 'stop')}` : ''}`
    : 'Nothing here can go into a PDF';
  $('multi-stop').checked = state.multiStop;
  $('multi-stop').disabled = state.busy;
  const anyChecked = checkedFiles().length > 0;
  const problems = teamRateCon() ? rcProblems() : [];
  renderRcBanner(problems);
  for (const id of ['btn-save', 'name-save', 'load-go']) $(id).disabled = state.busy || !anyChecked || problems.length > 0;
  $('btn-save').title = !anyChecked ? 'Check the files to save first.' : problems.length ? 'Leave out the rate confirmation first.' : '';
  $('btn-email').disabled = state.busy; // still opens, to reach Manage templates
  for (const id of ['keep-add-anyway', 'keep-ask-restart']) if ($(id)) $(id).disabled = state.busy;
  for (const box of document.querySelectorAll('#file-list .file-label > input[type="checkbox"]')) {
    const file = state.files[Number(box.closest('.file').dataset.id)];
    box.disabled = state.busy || !canCombine(file);
    box.checked = state.checked.has(file.id);
  }
  for (const pick of document.querySelectorAll('#file-list .stop-pick, #pv-body .stop-pick, #file-list .tag-chip input, #pv-body .tag-chip input, #file-list .order-grip')) pick.disabled = state.busy;
  for (const button of document.querySelectorAll('#file-list .order-step')) button.disabled = state.busy || button.dataset.edge === 'yes';
  if (!$('preview').hidden) {
    syncPreviewInclude();
    syncMoveButtons();
    syncPreviewOrder();
  }
  // The open name box or create-email step lists the files it will make
  if (!$('name-form').hidden) updateNameHint();
  if (emailDraft) syncAtChoices(); // the stops in the Shipper/Receiver drop-downs follow the files
  if (emailDraft && !$('load-hint').classList.contains('error')) updateLoadHint();
}

function setChecked(id, on) {
  const file = state.files.find((f) => f.id === id);
  if (on) state.checked.add(id);
  else state.checked.delete(id);
  if (on && state.multiStop && file) assignStop(file);
  state.built.clear();
  normalizeOrder(); // a receipt that's ticked again goes back below the PODs
  updateSelectionUi();
}

/** Ticking a file by hand (in the list or the preview). Page picks are kept. */
function onRowToggle(file, on) {
  setChecked(file.id, on);
  refreshList();
  if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
}

/** Redraws the list after a file changes: that can move other files or change their places. */
function refreshList() {
  renderList();
}

// --------------------------------------------------------------------------- rate confirmation check

/** The team's sample rate confirmation (set by the team admin), or null when the check is off. */
function teamRateCon() {
  const settings = teamSettings();
  return settings && settings.rateCon ? settings.rateCon : null;
}

/** Words of 4+ letters on a page, lower case. */
function wordSet(text) {
  return new Set(String(text || '').toLowerCase().match(/[a-z]{4,}/g) || []);
}

async function pageText(page) {
  const content = await page.getTextContent();
  return content.items.map((item) => item.str || '').join(' ');
}

/**
 * A page's layout: how much ink is in each cell of a grid laid over its printed area (margins
 * trimmed, so a scan that's shifted or a little smaller still lines up), and where that printed
 * area is on the page: { grid, box: [top, left, bottom, right] as shares of the page }. Null for a
 * blank page.
 */
function layoutGrid(canvas) {
  const w = canvas.width;
  const h = canvas.height;
  const px = canvas.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
  const n = w * h;
  const lum = new Uint8Array(n);
  const hist = new Uint32Array(256);
  for (let i = 0, j = 0; i < n; i++, j += 4) {
    const v = (px[j] * 77 + px[j + 1] * 150 + px[j + 2] * 29) >> 8;
    lum[i] = v;
    hist[v] += 1;
  }
  let seen = 0;
  let paper = 255;
  for (let v = 0; v < 256; v++) {
    seen += hist[v];
    if (seen >= n * 0.9) { paper = v; break; } // the paper's brightness
  }
  const cut = Math.max(40, paper * 0.72);
  const ink = new Uint8Array(n);
  const rows = new Uint32Array(h);
  const cols = new Uint32Array(w);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (lum[i] < cut) {
        ink[i] = 1;
        rows[y] += 1;
        cols[x] += 1;
      }
    }
  }
  const rowMin = Math.max(2, w * 0.01);
  const colMin = Math.max(2, h * 0.01);
  let top = 0;
  while (top < h && rows[top] < rowMin) top++;
  let bottom = h - 1;
  while (bottom > top && rows[bottom] < rowMin) bottom--;
  let left = 0;
  while (left < w && cols[left] < colMin) left++;
  let right = w - 1;
  while (right > left && cols[right] < colMin) right--;
  if (bottom - top < 10 || right - left < 10) return null;
  const { gridW: gw, gridH: gh } = RC;
  const bw = right - left + 1;
  const bh = bottom - top + 1;
  const cells = new Float32Array(gw * gh);
  for (let gy = 0; gy < gh; gy++) {
    const y0 = top + Math.floor((gy * bh) / gh);
    const y1 = Math.max(y0 + 1, top + Math.floor(((gy + 1) * bh) / gh));
    for (let gx = 0; gx < gw; gx++) {
      const x0 = left + Math.floor((gx * bw) / gw);
      const x1 = Math.max(x0 + 1, left + Math.floor(((gx + 1) * bw) / gw));
      let sum = 0;
      for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) sum += ink[y * w + x];
      cells[gy * gw + gx] = sum / ((y1 - y0) * (x1 - x0));
    }
  }
  const out = new Uint8Array(gw * gh); // softened, so neighbouring cells blend a little
  for (let gy = 0; gy < gh; gy++) {
    for (let gx = 0; gx < gw; gx++) {
      let sum = 0;
      let count = 0;
      for (let dy = -1; dy <= 1; dy++) {
        for (let dx = -1; dx <= 1; dx++) {
          const yy = gy + dy;
          const xx = gx + dx;
          if (yy < 0 || xx < 0 || yy >= gh || xx >= gw) continue;
          sum += cells[yy * gw + xx];
          count += 1;
        }
      }
      out[gy * gw + gx] = Math.min(255, Math.round((sum / count) * 510));
    }
  }
  const share = (v, total) => Math.round((v / total) * 100) / 100;
  return { grid: out, box: [share(top, h), share(left, w), share(bottom + 1, h), share(right + 1, w)] };
}

/** Draws a picture (photo, fax page) at the check's size and measures its layout. */
function gridFromSource(source, w, h) {
  const canvas = document.createElement('canvas');
  canvas.width = RC.renderW;
  canvas.height = Math.max(1, Math.round((RC.renderW * h) / Math.max(1, w)));
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  return layoutGrid(canvas);
}

async function pageGrid(page) {
  const base = page.getViewport({ scale: 1 });
  const viewport = page.getViewport({ scale: RC.renderW / base.width });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  await page.render({ canvas, viewport }).promise;
  return layoutGrid(canvas);
}

/** How alike two grids are: correlation, from -1 to 1. */
function likeness(a, b) {
  if (!a || !b || a.length !== b.length) return 0;
  const n = a.length;
  let ma = 0;
  let mb = 0;
  for (let i = 0; i < n; i++) {
    ma += a[i];
    mb += b[i];
  }
  ma /= n;
  mb /= n;
  let num = 0;
  let da = 0;
  let db = 0;
  for (let i = 0; i < n; i++) {
    const x = a[i] - ma;
    const y = b[i] - mb;
    num += x * y;
    da += x * x;
    db += y * y;
  }
  return da && db ? num / Math.sqrt(da * db) : 0;
}

const sampleGrids = new Map(); // base64 -> bytes
function sampleGrid(base64) {
  if (!base64) return null;
  if (!sampleGrids.has(base64)) sampleGrids.set(base64, base64ToBytes(base64));
  return sampleGrids.get(base64);
}

/** Share of a sample page's words found on this page (the best sample page counts). */
function textLikeness(words, sample) {
  let best = 0;
  for (const page of sample.pages) {
    const list = page.words ? page.words.split(' ') : [];
    if (list.length < 10) continue;
    let hits = 0;
    for (const word of list) if (words.has(word)) hits += 1;
    best = Math.max(best, hits / list.length);
  }
  return best;
}

/**
 * How alike a scanned page's layout is to the sample's pages (the best one counts). A page whose
 * printed area is a different size (say, two lines at the top of the page) doesn't count, since
 * its few lines would otherwise be stretched to fill the grid.
 */
function layoutLikeness(layout, sample) {
  if (!layout) return -1;
  let best = -1;
  for (const page of sample.pages) {
    const box = page.box || [0, 0, 1, 1];
    const tall = Math.abs((layout.box[2] - layout.box[0]) - (box[2] - box[0]));
    const wide = Math.abs((layout.box[3] - layout.box[1]) - (box[3] - box[1]));
    if (tall > RC.boxSlack || wide > RC.boxSlack) continue;
    best = Math.max(best, likeness(layout.grid, sampleGrid(page.grid)));
  }
  return best;
}

/** Makes the team's sample from one of your rate confirmations (a PDF): its words and layout. */
async function rateConSample(file) {
  const doc = await openPdf(file.bytes);
  try {
    const pages = [];
    for (let p = 1; p <= Math.min(doc.numPages, RC.samplePages); p++) {
      const page = await doc.getPage(p);
      const words = [...wordSet(await pageText(page))].slice(0, RC.sampleWords);
      const layout = await pageGrid(page);
      page.cleanup();
      pages.push({ words: words.join(' '), grid: layout ? bytesToBase64(layout.grid) : '', box: layout ? layout.box : null });
    }
    return { v: 1, name: file.name.slice(0, 120), saved: Date.now(), pages };
  } finally {
    closePdf(doc);
  }
}

function cleanRateCon(value) {
  if (!value || typeof value !== 'object' || !Array.isArray(value.pages)) return null;
  const pages = value.pages.slice(0, RC.samplePages).map((page) => {
    const grid = /^[A-Za-z0-9+/=]+$/.test(String((page && page.grid) || '')) && base64ToBytes(page.grid).length === RC.gridW * RC.gridH ? page.grid : '';
    const box = page && Array.isArray(page.box) && page.box.length === 4 && page.box.every((v) => typeof v === 'number' && v >= 0 && v <= 1) ? page.box : [0, 0, 1, 1];
    return {
      words: String((page && page.words) || '').split(' ').filter((w) => /^[a-z]{4,}$/.test(w)).slice(0, RC.sampleWords).join(' '),
      grid,
      box: grid ? box : null,
    };
  }).filter((page) => page.words || page.grid);
  return pages.length ? { v: 1, name: String(value.name || 'rate confirmation').slice(0, 120), saved: Number(value.saved) || 0, pages } : null;
}

/** The numbers printed under barcodes on a page, like *AXLL-3668149* (the text between the asterisks). */
function barcodeNumbers(text) {
  const found = new Set();
  for (const match of String(text || '').matchAll(/\*\s?([A-Z0-9][A-Z0-9.\/-]{4,30}[A-Z0-9])\s?\*/gi)) {
    const code = match[1].toUpperCase();
    if (!/\d{4}/.test(code)) continue;
    if (/^\d{1,4}[\/.-]\d{1,2}[\/.-]\d{1,4}$/.test(code)) continue; // a date, like *10/16/2026*
    if (/^\d{3}[.-]\d{3}[.-]\d{4}$/.test(code)) continue;          // a phone number
    found.add(code);
  }
  return [...found];
}

/** Checks one file's pages against the sample. Marks file.rcPages (0-based pages that look like it). */
async function checkFileForRateCon(file, sample, token) {
  const found = new Set();
  const scores = [];
  const judge = (label, score, limit) => {
    const hit = score >= limit;
    scores.push(`${label} ${score.toFixed(2)}${hit ? ' ⚑' : ''}`);
    return hit;
  };
  if (file.kind === 'PDF') {
    let doc;
    try {
      doc = await openPdf(file.bytes);
    } catch (error) {
      return; // password-protected or damaged: it can't be read, so it can't be sent either
    }
    try {
      file.pages = doc.numPages;
      const codes = []; // the numbers printed under barcodes on each page
      for (let p = 1; p <= Math.min(doc.numPages, RC.maxPages); p++) {
        if (token !== state.loadToken) return;
        const page = await doc.getPage(p);
        const text = await pageText(page);
        const words = wordSet(text);
        codes.push(barcodeNumbers(text));
        let hit;
        const named = /\b(rate|load)\s+confirmation\b/i.exec(text); // not "Download confirmation"
        if (named) {
          hit = judge(`p${p} says "${named[0].toLowerCase().replace(/\s+/, ' ')}"`, 1, 0.5);
        } else {
          // Words and layout both count, whichever is closer: a sample made from pictures of pages has
          // no words to compare, and the words a scanner app reads off a page often come out garbled
          const byText = words.size >= RC.minWords ? textLikeness(words, sample) : 0;
          let byLayout = -1;
          try {
            byLayout = layoutLikeness(await pageGrid(page), sample);
          } catch (error) {
            scores.push(`p${p} couldn't be drawn (${error.message})`); // its words still count
          }
          hit = byText >= RC.textMatch || byLayout >= RC.layoutMatch;
          scores.push(`p${p} text ${byText.toFixed(2)} layout ${byLayout.toFixed(2)}${hit ? ' ⚑' : ''}`);
        }
        page.cleanup();
        if (hit) found.add(p - 1);
      }
      // A page with the same barcode number as a flagged page is part of the same rate con
      const flagged = new Set();
      for (const i of found) for (const code of codes[i]) flagged.add(code);
      codes.forEach((list, i) => {
        if (found.has(i)) return;
        const shared = list.find((code) => flagged.has(code));
        if (!shared) return;
        found.add(i);
        scores.push(`p${i + 1} has the same barcode as a flagged page (${shared}) ⚑`);
      });
    } finally {
      closePdf(doc);
    }
  } else if (file.kind === 'Image') {
    const image = await decodeImage(file);
    try {
      const [w, h] = sizeOf(image);
      if (judge('layout', layoutLikeness(gridFromSource(image, w, h), sample), RC.layoutMatch)) found.add(0);
    } finally {
      if (typeof image.close === 'function') image.close();
    }
  } else if (file.kind === 'TIFF') {
    decodeTiffPages(file.bytes, 10).forEach((page, i) => {
      const stretch = page.xres > 0 && page.yres > 0 ? page.yres / page.xres : 1;
      if (judge(`p${i + 1} layout`, layoutLikeness(gridFromSource(page.canvas, page.w * stretch, page.h), sample), RC.layoutMatch)) found.add(0);
    });
  }
  if (token !== state.loadToken) return;
  file.rcPages = found;
  if (scores.length) log(`Rate con check: ${file.name}: ${scores.join(', ')}`);
}

/** Runs the check over the thread's files in the background (only when the team has a sample). */
async function checkRateCons(token) {
  const sample = teamRateCon();
  if (!sample) return;
  state.rcPending = true;
  try {
    for (const file of state.files) {
      if (token !== state.loadToken) return;
      if (file.rcPages || !canCombine(file)) continue;
      try {
        await checkFileForRateCon(file, sample, token);
      } catch (error) {
        log(`Rate con check: couldn't check ${file.name} (${error.message})`);
      }
      if (token !== state.loadToken) return;
      if (file.rcPages && file.rcPages.size) refreshList();
      else updateSelectionUi();
      if (!$('preview').hidden && state.files[state.previewIndex] === file) syncPreviewPages();
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  } finally {
    if (token === state.loadToken) {
      state.rcPending = false;
      updateSelectionUi();
    }
  }
}

/** Starts the check over (after the sample changes). */
function recheckRateCons() {
  for (const file of state.files) {
    file.rcPages = null;
    file.rcOk = false;
  }
  state.rcPending = false;
  renderList();
  checkRateCons(state.loadToken).catch((error) => log(`Rate con check: stopped (${error.message})`));
}

/** Pages that look like a rate confirmation, for the list and the preview (null when none). */
function rcFlags(file) {
  return teamRateCon() && file.rcPages && file.rcPages.size && !file.rcOk ? file.rcPages : null;
}

/** Checked files with pages that look like a rate confirmation and are still going in: [{ file, pages }]. */
function rcProblems() {
  const out = [];
  for (const file of checkedFiles()) {
    const flags = rcFlags(file);
    if (!flags) continue;
    if (file.kind !== 'PDF') {
      out.push({ file, pages: null });
      continue;
    }
    const pages = (allIncludedPages(file) || [...flags]).filter((p) => flags.has(p));
    if (pages.length) out.push({ file, pages });
  }
  return out;
}

/** "(pages 2–3)" when only some of a PDF's pages are flagged, else '' (the whole file is). */
function rcPagesText(file, pages) {
  return pages && file.kind === 'PDF' && pages.length < file.pages ? pagesLabel(pages).toLowerCase() : '';
}

/** The list's note for a file with rate-con-like pages: { text, going } (going: those pages are checked to go in). */
function rcNote(file) {
  const flags = rcFlags(file);
  if (!flags) return null;
  const which = rcPagesText(file, [...flags]);
  const going = state.checked.has(file.id) && (file.kind !== 'PDF' || (allIncludedPages(file) || [...flags]).some((p) => flags.has(p)));
  return { text: `${word('looksLikeRateCon')}${which ? `: ${which}` : ''}`, going };
}

/** Shows (or hides) the warning above the buttons. */
function renderRcBanner(problems) {
  const banner = $('rc-banner');
  banner.hidden = !problems.length;
  if (!problems.length) return;
  const list = problems.slice(0, 3).map(({ file, pages }) => {
    const which = rcPagesText(file, pages);
    return which ? `${file.name} (${which})` : file.name;
  });
  if (problems.length > 3) list.push(`${problems.length - 3} more`);
  const many = problems.length > 1;
  $('rc-banner-text').textContent = many
    ? `These look like rate confirmations: ${list.join(', ')}. They have to come out before you email or save.`
    : `This looks like a rate confirmation: ${list[0]}. It has to come out before you email or save.`;
  $('rc-leave-out').textContent = word(many ? 'leaveThemOut' : 'leaveItOut');
}

/** "Leave it out": unticks the flagged pages (or the whole file). */
function onRcLeaveOut() {
  const problems = rcProblems();
  for (const { file, pages } of problems) {
    if (!pages) {
      state.checked.delete(file.id);
      continue;
    }
    const keep = (allIncludedPages(file) || []).filter((p) => !file.rcPages.has(p));
    if (!keep.length) {
      file.pageChoice = null;
      state.checked.delete(file.id);
    } else {
      file.pageChoice = new Set(keep);
      normalizePageStops(file);
      normalizePageLumper(file);
    }
  }
  state.built.clear();
  log(`Rate con check: left out ${problems.map(({ file, pages }) => (pages ? `${file.name} ${pagesLabel(pages).toLowerCase()}` : file.name)).join(', ')}`);
  countUsage('rateCon', { pages: problems.reduce((sum, { file, pages }) => sum + (pages ? pages.length : (file.pages || 1)), 0), files: problems.length });
  renderList();
  if (!$('preview').hidden) syncPreviewPages();
  setStatus('Rate confirmation left out.', 'ok');
}

/** "Not a rate con": stops warning about these files for this thread. */
function onRcKeep() {
  const problems = rcProblems();
  for (const { file } of problems) file.rcOk = true;
  log(`Rate con check: kept ${problems.map(({ file }) => file.name).join(', ')} (not a rate con)`);
  renderList();
  if (!$('preview').hidden) syncPreviewPages();
}

/** Why sending or saving has to wait, or ''. */
function rcBlock() {
  if (!teamRateCon()) return '';
  if (rcProblems().length) return 'Leave out the rate confirmation first, or tell the add-in it isn\'t one.';
  if (state.rcPending) return 'Still checking the files for rate confirmations. Try again in a moment.';
  return '';
}

/** The admin's "Use as rate con sample" in the preview. */
async function onUseRateConSample() {
  const file = state.files[state.previewIndex];
  if (!file || file.kind !== 'PDF' || !isAdmin()) return;
  const button = $('pv-rc-sample');
  button.disabled = true;
  try {
    const sample = await rateConSample(file);
    if (!sample.pages.some((page) => page.words || page.grid)) throw userError("nothing could be read from this file");
    await saveTeamDraft({ rateCon: sample });
    const readable = sample.pages.filter((page) => page.words.split(' ').filter(Boolean).length >= 10).length;
    log(`Rate con check: sample set from ${file.name} (${plural(sample.pages.length, 'page')}, readable text on ${readable}), not published yet`);
    button.textContent = `Saved as the sample (${plural(sample.pages.length, 'page')}${readable ? '' : ', pictures of pages, so it\'s matched by layout'}). Publish it (Settings → Team).`;
    recheckRateCons();
  } catch (error) {
    button.textContent = `Couldn't use it: ${error.userMessage || error.message}`;
  } finally {
    button.disabled = false;
  }
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
    const turn = rotationOf(file, firstPage(file));
    if (turn) img.style.transform = `rotate(${turn}deg)`; // as turned in the preview
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
    const page = await doc.getPage(Math.min(firstPage(file) + 1, doc.numPages)); // the page that goes in first
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
    const first = firstPage(file);
    const page = decodeTiffPages(file.bytes, first + 1)[first] || decodeTiffPages(file.bytes, 1)[0];
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
  const sampleButton = $('pv-rc-sample');
  sampleButton.hidden = !(isAdmin() && file.kind === 'PDF' && file.bytes);
  sampleButton.textContent = word('useAsRcSample');
  syncPreviewInclude();
  updatePreviewNav();
  $('pv-order-reset').hidden = true;

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
  $('pv-include-label').textContent = canCombine(file) ? word('includeInPdf') : (WHY_NOT[file.kind] || "Can't go into the PDF");
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
    const seq = pageSeq(file).slice(0, count); // in the order they go into the PDF
    for (const index of seq) {
      const p = index + 1;
      if (token !== state.previewToken) return;
      const page = await doc.getPage(p);
      const rotation = (((page.rotate || 0) + rotationOf(file, p - 1)) % 360 + 360) % 360; // as turned in the preview
      const base = page.getViewport({ scale: 1, rotation });
      const viewport = page.getViewport({ scale: (width * ratio) / base.width, rotation });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      canvas.setAttribute('aria-label', `Page ${p}`);
      body.append(pageHolder(file, p - 1, canvas));
      syncMoveButtons();
      await page.render({ canvas, viewport }).promise;
      page.cleanup();
    }
    if (total > count) body.append(el('p', 'notice', `Showing the first ${count} of ${total} pages.`));
    syncPreviewPages();
    syncPreviewOrder();
  } finally {
    closePdf(doc);
  }
}

// ---- turning pages and photos (a quarter turn clockwise per click, for the PDF; the attachment stays as it is)

const ROTATE_SVG = '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 4v6h-6"/><path d="M19.4 15a8.5 8.5 0 1 1-2-8.8L21 10"/></svg>';

/** How far a page (or the photo, index 0) has been turned in the preview: 0, 90, 180 or 270 degrees clockwise. */
function rotationOf(file, index) {
  return (file.rotations && file.rotations.get(index)) || 0;
}

/** One more quarter turn for a page; returns its new rotation. */
function turnPage(file, index) {
  if (!file.rotations) file.rotations = new Map();
  const next = (rotationOf(file, index) + 90) % 360;
  if (next) file.rotations.set(index, next);
  else file.rotations.delete(index);
  state.built.clear();
  log(`Rotate: ${file.name}${file.kind === 'Image' ? '' : ` page ${index + 1}`} now turned ${next}°`);
  if (index === firstPage(file)) placeThumb(file);
  return next;
}

/** A copy of a canvas turned clockwise by 90, 180 or 270 degrees (the same canvas for 0). */
function turnCanvas(source, deg) {
  if (!deg) return source;
  const canvas = document.createElement('canvas');
  const quarter = deg % 180 !== 0;
  canvas.width = quarter ? source.height : source.width;
  canvas.height = quarter ? source.width : source.height;
  const ctx = canvas.getContext('2d');
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate((deg * Math.PI) / 180);
  ctx.drawImage(source, -source.width / 2, -source.height / 2);
  const label = source.getAttribute('aria-label');
  if (label) canvas.setAttribute('aria-label', label);
  if (source.style.aspectRatio) {
    const [w, h] = source.style.aspectRatio.split('/').map((v) => v.trim());
    canvas.style.aspectRatio = quarter ? `${h} / ${w}` : `${w} / ${h}`;
  }
  return canvas;
}

/** The Rotate button for a page or photo in the preview. It turns the picture there straight away. */
function rotateButton(file, index, holder) {
  const button = el('button', 'icon-btn pv-rotate');
  button.type = 'button';
  const what = file.kind === 'Image' ? 'the photo' : `page ${index + 1}`;
  button.setAttribute('aria-label', `${word('rotate')} ${what}`);
  button.title = word('rotate');
  button.innerHTML = ROTATE_SVG;
  button.addEventListener('click', () => {
    turnPage(file, index);
    const canvas = holder.querySelector('canvas');
    if (canvas) canvas.replaceWith(turnCanvas(canvas, 90));
  });
  return button;
}

/** A photo or fax page in the preview: its label and Rotate button above the picture. */
function turnHolder(file, index, canvas, label) {
  const holder = el('div', 'pv-page');
  holder.dataset.page = String(index);
  const top = el('div', 'pv-page-top');
  top.append(el('span', 'pv-page-label', label));
  if (file.kind === 'TIFF') top.append(moveButtons(file, index, holder));
  top.append(rotateButton(file, index, holder));
  holder.append(top, turnCanvas(canvas, rotationOf(file, index)));
  return holder;
}

/** Turns the pages just added to `out` for a file the way they were turned in the preview. */
function applyTurns(out, file, from, only) {
  if (!file.rotations || !file.rotations.size) return;
  const pages = out.getPages().slice(from);
  const sources = file.kind === 'PDF' && only ? only : file.kind === 'TIFF' ? pageSeq(file) : pages.map((_, i) => i);
  pages.forEach((page, i) => {
    const turn = rotationOf(file, sources[i]);
    if (!turn) return;
    const now = ((page.getRotation().angle % 360) + 360) % 360;
    page.setRotation(PDFLib.degrees((now + turn) % 360));
  });
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
  const chip = lumperChip(false, false, `Page ${index + 1}: ${word('receipt')}`, (on) => setPageLumper(file, index, on));
  const tags = el('div', 'pv-page-tags');
  const badge = el('span', 'rc-badge', word('rateConBadge'));
  badge.title = 'This page looks like a rate confirmation';
  tags.append(badge, pick, chip, rotateButton(file, index, holder));
  top.append(bar, moveButtons(file, index, holder), tags);
  holder.append(top, canvas);
  syncPageHolder(holder, file); // right from the start, while later pages still render
  return holder;
}

/** Up and down arrows that move a page of a PDF or fax within the file (shown when it has 2 pages or more). */
function moveButtons(file, index, holder) {
  const box = el('span', 'pv-move');
  const arrow = (svg, step, key) => {
    const button = el('button', 'icon-btn pv-move-btn');
    button.type = 'button';
    button.innerHTML = svg;
    button.dataset.step = String(step);
    button.title = word(key);
    button.setAttribute('aria-label', `${word(key)}: page ${index + 1}`);
    button.addEventListener('click', () => movePage(file, holder, step, button));
    return button;
  };
  box.append(arrow(UP_SVG, -1, 'movePageUp'), arrow(DOWN_SVG, 1, 'movePageDown'));
  box.hidden = pageTotal(file) < 2;
  return box;
}

/** Turns each page's arrows off at the top and bottom (and while Restack is busy). */
function syncMoveButtons() {
  const holders = [...document.querySelectorAll('#pv-body .pv-page')];
  holders.forEach((holder, i) => {
    for (const button of holder.querySelectorAll('.pv-move-btn')) {
      const step = Number(button.dataset.step);
      button.disabled = state.busy || (step < 0 ? i === 0 : i === holders.length - 1);
    }
  });
}

/** Moves a page up (-1) or down (1) in the preview, and in the PDF. */
function movePage(file, holder, step, button) {
  const holders = [...document.querySelectorAll('#pv-body .pv-page')];
  const at = holders.indexOf(holder);
  const other = holders[at + step];
  if (at < 0 || !other) return;
  const seq = pageSeq(file);
  const a = seq.indexOf(Number(holder.dataset.page));
  const b = seq.indexOf(Number(other.dataset.page));
  if (a < 0 || b < 0) return;
  const first = firstPage(file);
  [seq[a], seq[b]] = [seq[b], seq[a]];
  file.pageOrder = seq.some((p, i) => p !== i) ? seq : null;
  if (firstPage(file) !== first) refreshThumb(file);
  if (step < 0) other.before(holder);
  else other.after(holder);
  state.built.clear();
  log(`Page order: ${file.name} ${file.pageOrder ? `now ${seq.map((p) => p + 1).join(', ')}` : 'back as it came'}`);
  syncMoveButtons();
  syncPreviewOrder();
  refreshList();
  (button && !button.disabled ? button : holder.querySelector('.pv-move-btn:not(:disabled)') || holder.querySelector('.pv-rotate')).focus();
  holder.scrollIntoView({ block: 'nearest' });
}

/** Puts a file's pages back as they came. */
function resetPageOrder() {
  const file = state.files[state.previewIndex];
  if (!file || !file.pageOrder) return;
  const first = firstPage(file);
  file.pageOrder = null;
  if (first !== 0) refreshThumb(file);
  state.built.clear();
  log(`Page order: ${file.name} back as it came`);
  const body = $('pv-body');
  const holders = [...body.querySelectorAll('.pv-page')].sort((x, y) => Number(x.dataset.page) - Number(y.dataset.page));
  const notice = body.querySelector(':scope > .notice');
  for (const holder of holders) body.insertBefore(holder, notice);
  syncMoveButtons();
  syncPreviewOrder();
  refreshList();
  $('pv-include').focus();
}

/** The preview's "Reset page order" link: there once a page has been moved. */
function syncPreviewOrder() {
  const file = state.files[state.previewIndex];
  $('pv-order-reset').hidden = !(file && movedPages(file));
  $('pv-order-reset').disabled = state.busy;
}

function pageStopOf(file, index) {
  return (file.pageStops && file.pageStops.get(index)) || file.stop || 1;
}

function syncPageHolder(holder, file) {
  const index = Number(holder.dataset.page);
  const box = holder.querySelector('.pv-page-bar input');
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
      for (let stop = 1; stop <= choices; stop++) pick.append(new Option(word('stop', { n: stop }), String(stop)));
    }
    pick.value = String(pageStopOf(file, index));
  }
  const flags = rcFlags(file);
  holder.querySelector('.rc-badge').hidden = !(flags && flags.has(index));
  const chip = holder.querySelector('.tag-chip');
  const lumper = pageLumperOf(file, index);
  chip.hidden = !box.checked;
  chip.classList.toggle('on', lumper);
  chip.querySelector('input').checked = lumper;
  chip.querySelector('input').disabled = state.busy;
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
    if (holder.querySelector('.pv-page-bar input').checked) pages.add(Number(holder.dataset.page));
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
  normalizePageLumper(file);
  state.built.clear();
  refreshList();
  syncPreviewPages();
}

async function previewImage(file, body, token, waiting) {
  const image = await decodeImage(file);
  if (token !== state.previewToken) return;
  const canvas = rasterize(image, PHOTO_MAX_SIDE); // as it goes into the PDF
  waiting.remove();
  body.append(turnHolder(file, 0, canvas, 'Photo'));
  const [w, h] = sizeOf(image);
  $('pv-pages').textContent = `${w} × ${h}`;
}

async function previewTiff(file, body, token, waiting) {
  const pages = decodeTiffPages(file.bytes);
  if (token !== state.previewToken) return;
  if (!pages.length) throw userError('This TIFF file has no pages.');
  waiting.remove();
  file.tiffPages = pages.length;
  $('pv-pages').textContent = plural(pages.length, 'page');
  for (const i of pageSeq(file).slice(0, PREVIEW_PAGE_LIMIT)) { // in the order they go into the PDF
    const page = pages[i];
    page.canvas.style.aspectRatio = `${page.w / page.xresOr1} / ${page.h / page.yresOr1}`;
    body.append(turnHolder(file, i, page.canvas, `Page ${i + 1}`));
  }
  if (pages.length > PREVIEW_PAGE_LIMIT) body.append(el('p', 'notice', `Showing the first ${PREVIEW_PAGE_LIMIT} of ${pages.length} pages.`));
  syncMoveButtons();
  syncPreviewOrder();
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
  // PDF.js takes ownership of the buffer it's given, so hand it a copy. wasmUrl is where its
  // decoders for fax/JBIG2 and JPEG 2000 scans are: without them those pages draw blank.
  const task = pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, wasmUrl: new URL('./wasm/', import.meta.url).href });
  try {
    return await task.promise;
  } catch (error) {
    try { task.destroy(); } catch (ignored) { /* already gone */ }
    if (error && error.name === 'PasswordException') throw userError('needs a password to open');
    throw userError("couldn't be read - the file may be damaged");
  }
}

/**
 * How a PDF's page images are compressed, for the test report: "fax (CCITT), JPEG 2000"…, or ''
 * for none of the scanner formats. Image dictionaries are never inside compressed object
 * streams, so their /Filter names can be found in the raw bytes.
 */
function scanFormats(bytes) {
  const text = new TextDecoder('latin1').decode(bytes.length > 40 * MB ? bytes.subarray(0, 40 * MB) : bytes);
  const found = [];
  if (/\/CCITTFaxDecode\b/.test(text)) found.push('fax (CCITT)');
  if (/\/JBIG2Decode\b/.test(text)) found.push('JBIG2');
  if (/\/JPXDecode\b/.test(text)) found.push('JPEG 2000');
  if (/\/DCTDecode\b/.test(text)) found.push('JPEG');
  return found.join(', ');
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
  const canvas = rasterize(await decodeImage(file), PHOTO_MAX_SIDE); // the whole photo, scaled (never cropped)
  const image = await out.embedJpg(await canvasToBytes(canvas, 'image/jpeg', 0.9));
  const [w, h] = fitPage(canvas.width, canvas.height);
  out.addPage([w, h]).drawImage(image, { x: 0, y: 0, width: w, height: h });
}

async function addTiff(out, file) {
  const pages = decodeTiffPages(file.bytes);
  if (!pages.length) throw userError('has no pages');
  file.tiffPages = pages.length;
  for (const index of pageSeq(file)) { // as moved in the preview
    const page = pages[index];
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
  const turns = (file) => (file.rotations ? [...file.rotations].map(([i, d]) => `${i}r${d}`).join('.') : '');
  const key = entries.map(({ file, only }) => `${file.fp || file.id}:${(only || includedPages(file) || ['all']).join('.')}:${turns(file)}:${file.pageOrder ? file.pageOrder.join('.') : ''}`).join(',');
  if (state.built.has(key)) return state.built.get(key);
  const out = await PDFLib.PDFDocument.create();
  const skipped = [];
  for (let i = 0; i < entries.length; i++) {
    const { file } = entries[i];
    const only = entries[i].only || includedPages(file);
    setStatus(`${label ? `${label}: ` : ''}Adding ${i + 1} of ${entries.length}: ${file.name}`);
    const before = out.getPageCount();
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
      applyTurns(out, file, before, only);
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
    out.setCreator('Restack');
    out.setProducer(`Restack ${VERSION}`);
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

/** Save PDF's one file: the checked files in the list's order, lumper receipts last. */
async function combineChecked() {
  const [group] = fileGroups(false, false);
  if (!group) return null;
  logOrder(pdfName(), group.entries);
  const result = await buildPdf(group.entries);
  if (!result.pages) {
    setStatus(`None of the checked files could go into a PDF.${leftOutText(result)}`, 'error');
    return null;
  }
  return result;
}

/**
 * The files to email or save. `layout` says whether stops and lumper receipts get files of their
 * own: base-1.pdf, base-2.pdf… per stop (whatever their size), and lumperStem.pdf (or
 * lumperStem-1.pdf… per stop) for lumper receipts. A single POD PDF over 10 MB is split into -1, -2… instead.
 * Returns { files, skipped, pages } or null (after saying why) when nothing could be built.
 */
async function makeFiles(base, layout = layoutOf(null), lumperStem = `${base} Receipt`) {
  const files = [];
  const skipped = [];
  for (const group of fileGroups(layout.splitStops, layout.splitLumpers, layout.noReceipts)) {
    const label = [group.lumper ? word('receipt') : '', group.stop ? word('stop', { n: group.stop }) : ''].filter(Boolean).join(', ');
    const result = await buildPdf(group.entries, label);
    skipped.push(...result.skipped);
    if (!result.pages) continue;
    const stem = `${group.lumper ? lumperStem : base}${group.stop ? `-${group.stop}` : ''}`;
    logOrder(`${stem}.pdf`, group.entries);
    if (!group.lumper && !group.stop) {
      for (const part of await splitIntoFiles(result.bytes, stem)) files.push({ ...part, pages: part.pages || result.pages });
    } else {
      files.push({ name: `${stem}.pdf`, bytes: result.bytes, pages: result.pages, stop: group.stop, lumper: group.lumper, over: false });
    }
  }
  if (!files.length) {
    setStatus(`None of the checked files could go into a PDF.${leftOutText({ skipped })}`, 'error');
    return null;
  }
  for (const file of files) if (file.stop || file.lumper) file.over = file.bytes.length > SPLIT_LIMIT; // these aren't split by size
  if (layout.splitStops || layout.splitLumpers) log(`Files: ${files.map((f) => `${f.name} ${mbText(f.bytes.length)} (${plural(f.pages, 'page')})`).join(', ')}`);
  return { files, skipped, pages: files.reduce((sum, f) => sum + (f.pages || 0), 0) };
}

// --------------------------------------------------------------------------- save, new email

/**
 * Save PDF. Where Outlook allows Windows' own Save As window, that opens straight away (you can
 * name the file and pick the folder there). Everywhere else, you name the file in the panel first.
 */
async function onSave() {
  if (!checkedFiles().length) return;
  const blocked = rcBlock();
  if (blocked) {
    setStatus(blocked, 'error');
    return;
  }
  if (typeof window.showSaveFilePicker === 'function' && readSetting('saveDialog') === 'works' && !state.multiStop && likelyOneFile()) {
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
  updateNameHint();
  input.focus();
  input.select();
}

/** With Multi-stop on, the name box says which files it will save. */
function updateNameHint() {
  const hint = $('name-hint');
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
  const blocked = rcBlock();
  if (blocked) {
    setStatus(blocked, 'error');
    return;
  }
  state.fileName = cleanFileName($('file-name').value);
  hideNameForm();
  const name = pdfName();
  log(`Save: named ${name}${state.multiStop ? ', one file per stop' : ''}`);
  if (state.multiStop || !likelyOneFile()) {
    saveFiles(name).catch(fail); // downloads: one file per stop, or -1, -2… over 10 MB
    return;
  }
  // Straight from the click, so the browser lets Windows' Save As window open if it can.
  pickSaveLocation(name)
    .then((picked) => (picked.cancelled ? null : finishSave(picked.handle || null, name)))
    .catch(fail);
}

/**
 * Save PDF as downloads: one file per stop with Multi-stop on; otherwise one file, or name-1.pdf,
 * name-2.pdf… when it's over 10 MB.
 */
async function saveFiles(name) {
  const result = await withBusy(() => makeFiles(name.replace(/\.pdf$/i, '')));
  if (!result) return;
  await startDownloads(result.files, result);
}

/** Downloads the files (a short pause between them) and says so. `lead` goes before the message. */
async function startDownloads(files, result, lead = '') {
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
  countUsage('save', usageFigures(files.map((f) => ({ ...f, pages: f.pages || (files.length === 1 ? result.pages : 0) }))));
  const what = files.length === 1
    ? `Downloading ${files[0].name} (${plural(files[0].pages || result.pages, 'page')}) to your Downloads folder.`
    : `Downloading ${files.length} files to your Downloads folder: ${files.map((f) => f.name).join(', ')}. If Outlook asks about downloading multiple files, choose Allow.`;
  setStatus(`${lead}${what}${oversizeNote(files)}${leftOutText(result)}`, 'ok');
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

/**
 * Whether Save PDF will most likely make one file (under 10 MB), judging by the checked files' sizes,
 * so Windows' Save As window can be used. Fax scans grow when they become PDF pages, so they count extra.
 */
function likelyOneFile() {
  let bytes = 0;
  for (const file of checkedFiles()) bytes += (file.size || 0) * (file.kind === 'TIFF' ? 8 : 1);
  return bytes < SPLIT_MB * 0.8 * MB;
}

/** Builds the PDF and writes it: into the file picked in Save As, or as a download (split if over 10 MB). */
async function finishSave(handle, name) {
  const result = await withBusy(combineChecked);
  if (!result) return;
  if (result.bytes.length > SPLIT_LIMIT) {
    // Over 10 MB: Save As can only make the one file, so the parts go to Downloads instead
    let lead = '';
    if (handle) {
      const picked = handle.name || name;
      name = picked;
      lead = `It came out at ${mbText(result.bytes.length)}, over ${SPLIT_MB} MB, so it's in parts in Downloads instead of where you picked. `;
      try {
        // Save As may have made an empty file there already (it's left alone: it could be one you had)
        if ((await handle.getFile()).size === 0) lead += `You can delete the empty ${picked} there. `;
      } catch (error) {
        // can't tell; say nothing
      }
      log(`Save dialog: ${picked} came out at ${mbText(result.bytes.length)}, saving it in parts instead`);
    }
    await downloadParts(result, name, lead);
    return;
  }

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
      countUsage('save', usageFigures([{ name: saved, pages: result.pages }]));
      return;
    } catch (error) {
      writeSetting('saveDialog', null);
      log(`Save dialog: couldn't write the file (${error.name}: ${error.message})`);
    }
  }
  try {
    downloadBytes(result.bytes, name); // (under 10 MB: bigger ones were split above)
    log(`Download: started as ${name}. Check whether it showed up in Downloads.`);
    setStatus(`Downloading ${name} (${plural(result.pages, 'page')}) to your Downloads folder.${leftOutText(result)}`, 'ok');
    countUsage('save', usageFigures([{ name, pages: result.pages }]));
  } catch (error) {
    log(`Download: failed (${error.name}: ${error.message})`);
    setStatus("Saving isn't possible here. Use Email instead.", 'error');
  }
}

/** A built PDF over 10 MB, saved as name-1.pdf, name-2.pdf… downloads. */
async function downloadParts(result, name, lead = '') {
  const files = await withBusy(() => splitIntoFiles(result.bytes, name.replace(/\.pdf$/i, '')));
  if (!files) return;
  await startDownloads(files, result, lead);
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

/** "pods@example.com · 4 questions" under a template's name (in Email ▾ and the template list). */
function templateSub(t) {
  const n = questionCount(t);
  return `${recipientsText(t) || 'No recipients set'}${n ? ` · ${plural(n, 'question')}` : ''}`;
}

/** Every email's subject, whatever the template: "3645493 POD(s)" (WORDS.subject). */
function emailSubject(load) {
  return word('subject', { load: String(load || '').trim() }).trim().slice(0, 255);
}

/** The PDF's name for an email, without .pdf: the template's PDF name, or just the load number. */
function emailFileBase(t, load) {
  const base = fill((t && t.file) || '{load}', load).trim().replace(/\.pdf$/i, '');
  return safeName(base || load || state.topic || 'Combined attachments');
}

/** The lumper receipt PDF's name for an email, without .pdf: the template's lumper file name. */
function lumperBase(t, load) {
  const base = fill(receiptName(t), load).trim().replace(/\.pdf$/i, '');
  return safeName(base || `${load} Receipt`);
}

/** The files an email will get, named before anything is built (for the hint under the load number). */
function plannedFiles(base, layout = layoutOf(null), lumperStem = `${base} Receipt`) {
  const groups = fileGroups(layout.splitStops, layout.splitLumpers, layout.noReceipts);
  if (!groups.length) return { names: [`${base}${layout.splitStops ? '-1' : ''}.pdf`], note: layout.splitStops ? '…' : '' };
  const names = groups.map((g) => `${g.lumper ? lumperStem : base}${g.stop ? `-${g.stop}` : ''}.pdf`);
  let note = '';
  if (groups.some((g) => !g.stop && !g.lumper) && !likelyOneFile()) note += ` (split into -1, -2… if it's over ${SPLIT_MB} MB)`;
  if (layout.splitLumpers && !layout.noReceipts && !groups.some((g) => g.lumper)) note += ` · no ${word('receipt').toLowerCase()} tagged`;
  return { names, note };
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
function bodyHtml(text, acc, t, lines = true) {
  const parts = [];
  if (String(text || '').trim()) parts.push(escapeHtml(String(text).slice(0, 30000)).replace(/\r?\n/g, '<br>'));
  const extra = accessorialHtml(acc, t, lines);
  if (extra) parts.push(extra);
  if (!parts.length) return '';
  return `<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 11pt;">${parts.join('<br><br>')}</div>`;
}

/**
 * The new email's To, CC, subject and message from a template (or a blank email when t is null).
 * lines false: the accessorials are left out (they go in their own email).
 */
function emailFields(t, load, acc, lines = true) {
  const form = { subject: emailSubject(load) };
  if (t) {
    const to = parseAddresses(fill(t.to, load));
    const cc = parseAddresses(fill(t.cc, load));
    if (to.length) form.toRecipients = to.slice(0, 100);
    if (cc.length) form.ccRecipients = cc.slice(0, 100);
  }
  const html = bodyHtml(t ? fill(t.body, load) : '', acc, t, lines);
  if (html) form.htmlBody = html;
  return form;
}

// ---- the separate accessorial email (a template setting)

/** The ticked accessorials' names, in the form's order: Detention, Layover, Lumper, then Other's typed type. */
function accessorialTypes(acc) {
  if (!acc) return [];
  const types = [];
  if (acc.det.on) types.push(word('detention'));
  if (acc.lay && acc.lay.on) types.push(word('layover'));
  if (acc.lum.on) types.push(word('lumper'));
  if (acc.oth.on) types.push(String(acc.oth.type || '').trim() || word('other'));
  return types;
}

/** Whether Create email makes two emails: the template says so and an accessorial is ticked. */
function separateAccessorials(t, acc) {
  return !!(t && t.accSeparate && accessorialTypes(acc).length);
}

/** "3645493 - Detention/Lumper Request" */
function accessorialSubject(load, acc) {
  return word('accessorialSubject', { load: String(load || '').trim(), types: accessorialTypes(acc).join('/') }).trim().slice(0, 255);
}

/** The accessorial email's To and CC: the template's own. A blank To uses the POD email's; a blank CC means no CC. */
function accessorialRecipients(t, load) {
  const to = parseAddresses(fill(t.accTo, load));
  return {
    to: (to.length ? to : parseAddresses(fill(t.to, load))).slice(0, 100),
    cc: parseAddresses(fill(t.accCc, load)).slice(0, 100),
  };
}

/** The accessorial email: its To, CC, subject, and a message with the accessorials' details. */
function accessorialFields(t, load, acc) {
  const form = { subject: accessorialSubject(load, acc) };
  const { to, cc } = accessorialRecipients(t, load);
  if (to.length) form.toRecipients = to;
  if (cc.length) form.ccRecipients = cc;
  const types = accessorialTypes(acc).join('/');
  const intro = String(t.accBody || '').trim()
    ? fill(t.accBody, load).replace(/\{types\}/gi, types)
    : word('accessorialIntro', { types, load: String(load || '').trim() });
  form.htmlBody = bodyHtml(intro, acc, null); // (No Lumper stays with the POD email)
  return form;
}

/** Whether any checked file (or page) is tagged Receipt, and whether anything else is. */
function receiptMix() {
  let receipts = false;
  let pods = false;
  for (const file of checkedFiles()) {
    for (const part of partsOfFile(file)) {
      if (part.lumper) receipts = true;
      else pods = true;
    }
  }
  return { receipts, pods };
}

// ---- accessorials (Detention, Layover, Lumper, Other) for the email body

const ACC_FIELDS = {
  det: ['at', 'in', 'appt', 'out', 'hours'],
  lay: ['at', 'in', 'appt', 'out'],               // like Detention, without the hours
  lum: ['at', 'total'],
  oth: ['at', 'type', 'total'],
};

function emptyAccessorials() {
  return {
    det: { on: false, at: '', in: '', appt: '', out: '', hours: '' },
    lay: { on: false, at: '', in: '', appt: '', out: '' },
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

/**
 * Where an accessorial was: Shipper or Receiver. With Multi-stop on, the receiver at each stop in
 * use instead, as "Receiver (Stop 2)", which is also how the email body says it.
 */
function atChoices() {
  if (!state.multiStop) return [word('shipper'), word('receiver')];
  const stops = [...stopsInUse()].sort((a, b) => a - b);
  return [word('shipper'), ...(stops.length ? stops.map((stop) => word('receiverAtStop', { n: stop })) : [word('receiver')])];
}

/** Brings the three Shipper/Receiver drop-downs in line with the stops, keeping each pick that's still there. */
function syncAtChoices() {
  const choices = atChoices();
  for (const kind of Object.keys(ACC_FIELDS)) {
    const select = $(`${kind}-at`);
    const current = select.value;
    const offered = [...select.options].slice(1).map((option) => option.value);
    if (offered.join('\n') !== choices.join('\n')) {
      select.length = 1; // keep the "Shipper or receiver?" prompt
      for (const choice of choices) select.append(new Option(choice, choice));
    }
    select.options[0].textContent = word(state.multiStop ? 'shipperOrStop' : 'shipperOrReceiver');
    select.value = choices.includes(current) ? current : '';
    select.classList.toggle('empty', !select.value);
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
  const where = state.multiStop ? `${word('shipper')} or a stop` : `${word('shipper')} or ${word('receiver')}`;
  if (acc.det.on && !acc.det.at) return ['det-at', `Pick ${where} for ${word('detention')}.`];
  if (acc.lay.on && !acc.lay.at) return ['lay-at', `Pick ${where} for ${word('layover')}.`];
  if (acc.lum.on && !acc.lum.at) return ['lum-at', `Pick ${where} for ${word('lumper')}.`];
  if (acc.lum.on && money(acc.lum.total) === null) return ['lum-total', `The ${word('lumper')} ${word('total').toLowerCase()} should be an amount, like 350.00.`];
  if (acc.oth.on && !acc.oth.at) return ['oth-at', `Pick ${where} for ${word('other')}.`];
  if (acc.oth.on && !acc.oth.type) return ['oth-type', `Type what the ${word('other').toLowerCase()} accessorial is.`];
  if (acc.oth.on && money(acc.oth.total) === null) return ['oth-total', `The ${word('other')} ${word('total').toLowerCase()} should be an amount, like 75.00.`];
  return null;
}

/**
 * The ticked accessorials as lines for the email body (blank entries are left out). A template
 * set to say "No Lumper" gets that line when Lumper isn't ticked.
 */
function accessorialHtml(acc, t, lines = true) {
  if (!acc) acc = emptyAccessorials();
  const blocks = [];
  const add = (title, at, entries) => {
    const lines = [`<b>${escapeHtml(title)}</b> – ${escapeHtml(at)}`];
    for (const [name, value] of entries) if (value) lines.push(`${escapeHtml(name)}: ${escapeHtml(value)}`);
    blocks.push(lines.join('<br>'));
  };
  if (lines && acc.det.on) add(word('detention'), acc.det.at, [[word('timeIn'), acc.det.in], [word('appointment'), acc.det.appt], [word('timeOut'), acc.det.out], [word('hoursRequested'), acc.det.hours]]);
  if (lines && acc.lay.on) add(word('layover'), acc.lay.at, [[word('timeIn'), acc.lay.in], [word('appointment'), acc.lay.appt], [word('timeOut'), acc.lay.out]]);
  if (acc.lum.on) {
    if (lines) add(word('lumper'), acc.lum.at, [[word('total'), money(acc.lum.total)]]);
  } else if (t && t.noLumper) blocks.push(`<b>No ${escapeHtml(word('lumper'))}</b>`);
  if (lines && acc.oth.on) add(word('other'), acc.oth.at, [[word('otherType'), acc.oth.type], [word('total'), money(acc.oth.total)]]);
  return blocks.join('<br><br>');
}

/** "Detention at Receiver, Lumper at Receiver $350.00" for the status line and the report. */
function accessorialSummary(acc) {
  if (!acc) return '';
  const parts = [];
  if (acc.det.on) parts.push(`${word('detention')} at ${acc.det.at}${acc.det.hours ? ` (${acc.det.hours} h)` : ''}`);
  if (acc.lay.on) parts.push(`${word('layover')} at ${acc.lay.at}`);
  if (acc.lum.on) parts.push(`${word('lumper')} at ${acc.lum.at}${money(acc.lum.total) ? ` ${money(acc.lum.total)}` : ''}`);
  if (acc.oth.on) parts.push(`${acc.oth.type} at ${acc.oth.at}${money(acc.oth.total) ? ` ${money(acc.oth.total)}` : ''}`);
  return parts.join(', ');
}

// ---- the menu

/** Templates in A to Z order by name (numbers in order too: Store 9 before Store 10). */
function byName(list) {
  return list.slice().sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base', numeric: true }));
}

function renderEmailMenu() {
  const menu = $('email-menu');
  // Remember what had focus, so pinning, moving or unpinning keeps your place
  const active = document.activeElement && document.activeElement.closest('#email-menu') ? document.activeElement : null;
  const focusedRow = active && active.closest('.menu-row');
  const focusedId = focusedRow ? focusedRow.dataset.id : '';
  const focusedKind = active ? active.dataset.pin || active.dataset.kind || '' : '';
  const shownPins = [...menu.querySelectorAll('.pin-btn')].filter((b) => !b.closest('[hidden]'));
  const focusedAt = active ? shownPins.indexOf(active) : -1;
  menu.textContent = '';
  const nothing = checkedFiles().length === 0;
  const blocked = !nothing && teamRateCon() && rcProblems().length > 0;
  const ready = !state.busy && !nothing && !blocked;
  if (nothing) menu.append(el('div', 'menu-note', 'Check the files to send first.'));
  if (blocked) menu.append(el('div', 'menu-note', 'Leave out the rate confirmation first.'));
  const teamList = byName(teamTemplates());
  const mine = canMakeOwn() ? byName(state.myTemplates) : []; // a rep's own templates stay saved, just unused, while it's off
  const all = [...teamList, ...mine];
  const pinned = pinnedTemplates();
  const isPinned = new Set(pinned.map((t) => t.id));
  const groups = (box, list) => { // the templates under their Team and Mine headings
    const team = list.filter((t) => teamList.includes(t));
    const own = list.filter((t) => mine.includes(t));
    if (team.length) {
      box.append(el('div', 'menu-group', word('team')));
      for (const t of team) box.append(templateRow(t, ready, null));
    }
    if (own.length) {
      box.append(el('div', 'menu-group', word('mine')));
      for (const t of own) box.append(templateRow(t, ready, null));
    }
  };
  if (pinned.length) {
    // Pinned templates first, in your order; the rest wait under Other templates
    menu.append(el('div', 'menu-group', word('pinned')));
    pinned.forEach((t, i) => menu.append(templateRow(t, ready, { at: i, of: pinned.length })));
    const rest = all.filter((t) => !isPinned.has(t.id));
    if (rest.length) {
      const toggle = el('button', 'menu-item menu-other-toggle');
      toggle.type = 'button';
      toggle.setAttribute('role', 'menuitem');
      toggle.setAttribute('aria-expanded', String(!!state.otherOpen));
      toggle.dataset.kind = 'other';
      toggle.append(el('span', 'mi-name', `${word('otherTemplates')} (${rest.length})`));
      toggle.insertAdjacentHTML('beforeend', '<svg class="chev" viewBox="0 0 20 20" width="14" height="14" aria-hidden="true"><path d="M5 7.5 10 12.5l5-5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>');
      const box = el('div', 'menu-other');
      box.id = 'menu-other';
      box.hidden = !state.otherOpen;
      groups(box, rest);
      toggle.addEventListener('click', () => {
        state.otherOpen = !state.otherOpen;
        toggle.setAttribute('aria-expanded', String(state.otherOpen));
        box.hidden = !state.otherOpen;
        if (state.otherOpen) (box.querySelector('.menu-item:not(:disabled)') || box.querySelector('.menu-item') || toggle).scrollIntoView({ block: 'nearest' });
      });
      menu.append(toggle, box);
    }
  } else {
    groups(menu, all);
  }
  if (all.length) menu.append(el('div', 'menu-sep'));
  menu.append(menuItem(word('blankEmail'), word('blankEmailNote'), () => startEmail(null), !ready));
  if (canMakeOwn()) menu.append(menuItem(word(all.length ? 'manageTemplates' : 'createTemplate'), '', openTemplates, false));
  // After pinning, moving or unpinning, focus stays with that template (or, once it's tucked away
  // under Other templates, with the pin now in its place)
  if (focusedId) {
    const row = menu.querySelector(`.menu-row[data-id="${CSS.escape(focusedId)}"]`);
    let target = row && !row.closest('[hidden]') ? row.querySelector(focusedKind ? `[data-pin="${focusedKind}"]` : '.menu-item') : null;
    if (target && target.disabled) target = row.querySelector('.pin-move:not(:disabled)') || row.querySelector('.pin-btn');
    if (!target && focusedAt >= 0) {
      const pins = [...menu.querySelectorAll('.pin-btn')].filter((b) => !b.closest('[hidden]'));
      target = pins[Math.min(focusedAt, pins.length - 1)] || menu.querySelector('.menu-other-toggle');
    }
    if (target) target.focus();
  } else if (focusedKind === 'other') {
    const toggle = menu.querySelector('.menu-other-toggle');
    if (toggle) toggle.focus();
  }
}

/** Your pinned templates that still exist (and that you can use), in your order. */
function pinnedTemplates() {
  const all = [...teamTemplates(), ...(canMakeOwn() ? state.myTemplates : [])];
  return prefs.pins.map((id) => all.find((t) => t.id === id)).filter(Boolean);
}

/** A template in Email ▾: its name and recipients, and its pin (with up and down arrows once it's pinned). */
function templateRow(t, ready, place) {
  const row = el('div', 'menu-row');
  row.dataset.id = t.id;
  row.setAttribute('role', 'none');
  const item = menuItem(t.name, templateSub(t), () => startEmail(t), !ready);
  row.append(item);
  if (place) {
    const arrow = (svg, step, label) => {
      const button = el('button', 'pin-move');
      button.type = 'button';
      button.innerHTML = svg;
      button.title = label;
      button.dataset.pin = step < 0 ? 'up' : 'down';
      button.setAttribute('aria-label', `${label}: ${t.name}`);
      button.disabled = step < 0 ? place.at === 0 : place.at === place.of - 1;
      button.addEventListener('click', () => movePin(t.id, step));
      return button;
    };
    row.classList.add('pinned');
    row.append(arrow(UP_SVG, -1, 'Move up'), arrow(DOWN_SVG, 1, 'Move down'));
  }
  const pin = el('button', 'pin-btn');
  pin.type = 'button';
  pin.dataset.pin = 'pin';
  pin.innerHTML = PIN_SVG;
  pin.setAttribute('aria-pressed', String(!!place));
  pin.title = word(place ? 'unpin' : 'pin');
  pin.setAttribute('aria-label', `${word(place ? 'unpin' : 'pin')}: ${t.name}`);
  pin.addEventListener('click', () => togglePin(t.id));
  row.append(pin);
  return row;
}

/** Pins a template to the top of Email ▾ (last among the pinned ones), or unpins it. Saved per person. */
function togglePin(id) {
  const on = !prefs.pins.includes(id);
  if (on && !pinnedTemplates().length) state.otherOpen = true; // the first pin: the rest stay in view, to pin more
  savePref('pins', on ? [...prefs.pins, id] : prefs.pins.filter((x) => x !== id));
  const t = [...teamTemplates(), ...state.myTemplates].find((x) => x.id === id);
  log(`Templates: ${on ? 'pinned' : 'unpinned'} "${t ? t.name : id}"`);
  renderEmailMenu();
}

/** Moves a pinned template up (-1) or down (1) among the pinned ones. */
function movePin(id, step) {
  const pins = pinnedTemplates().map((t) => t.id); // the ones you can see, in order
  const at = pins.indexOf(id);
  const to = at + step;
  if (at < 0 || to < 0 || to >= pins.length) return;
  [pins[at], pins[to]] = [pins[to], pins[at]];
  savePref('pins', [...pins, ...prefs.pins.filter((x) => !pins.includes(x))]); // pins you can't use right now keep their place at the end
  renderEmailMenu();
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
  state.otherOpen = false; // pinned templates first; Other templates opens the rest
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
  // Right arrow on a template reaches its pin (and arrows); left goes back
  const inRow = document.activeElement && document.activeElement.closest('.menu-row');
  if (inRow && (event.key === 'ArrowRight' || event.key === 'ArrowLeft')) {
    const parts = [...inRow.querySelectorAll('button:not(:disabled)')];
    const at = parts.indexOf(document.activeElement);
    const next = parts[at + (event.key === 'ArrowRight' ? 1 : -1)];
    if (next) next.focus();
    event.preventDefault();
    return;
  }
  if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
  event.preventDefault();
  // Alt+Up / Alt+Down on a pinned template moves it
  const row = document.activeElement && document.activeElement.closest('.menu-row.pinned');
  if (event.altKey && row) {
    movePin(row.dataset.id, event.key === 'ArrowUp' ? -1 : 1);
    const again = $('email-menu').querySelector(`.menu-row[data-id="${CSS.escape(row.dataset.id)}"] .menu-item`);
    if (again) again.focus();
    return;
  }
  const items = [...$('email-menu').querySelectorAll('.menu-item:not(:disabled)')].filter((b) => !b.closest('[hidden]'));
  if (!items.length) return;
  const at = items.indexOf(document.activeElement);
  const step = event.key === 'ArrowDown' ? 1 : -1;
  items[(at + step + items.length) % items.length].focus();
}

// ---- the create-email step: load number (every email's subject has it) and accessorials

let emailDraft = null; // { template } while the step is open; template is null for a blank email

/** Picking a template (or Blank email) from Email ▾: its questions first, if it has any, then the step. */
function startEmail(t) {
  const blocked = rcBlock();
  if (blocked) {
    setStatus(blocked, 'error');
    return;
  }
  if (questionCount(t)) {
    hideNameForm();
    setStatus('');
    askQuestions(t);
    return;
  }
  openLoadForm(t);
}

function openLoadForm(t) {
  emailDraft = { template: t };
  hideNameForm();
  $('save-actions').hidden = true;
  $('load-form').hidden = false;
  setStatus('');
  $('load-number').value = state.loadNumber || guessLoadNumber(state.fileName) || guessLoadNumber(state.topic) || '';
  syncAtChoices(); // before the remembered picks go back in, so a stop picked earlier can be found
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
  const acc = readAccessorials();
  const separate = separateAccessorials(t, acc);
  const layout = layoutOf(t);
  const base = emailFileBase(t, load || '…');
  const lumperStem = lumperBase(t, load || '…');
  const { names, note } = plannedFiles(base, separate ? { ...layout, noReceipts: true } : layout, lumperStem);
  const to = t ? recipientsText(t, load) : '';
  hint.textContent = '';
  const podEmail = !separate || receiptMix().pods; // only receipts checked: just the request
  if (podEmail) hint.append(`${t ? t.name : word('blankEmail')}: attaches `, namesFragment(names), `${note}${to ? ` · to ${to}` : ''}`);
  if (separate) {
    // The second email: the PODs and receipts, with the accessorials' details
    const request = plannedFiles(base, layout, lumperStem);
    const requestTo = accessorialRecipients(t, load).to.join(', ');
    const both = !layout.splitLumpers && receiptMix().receipts ? ' (PODs and receipts)' : '';
    if (podEmail) hint.append(el('br'));
    hint.append(`${accessorialSubject(load || '…', acc)}: attaches `, namesFragment(request.names), `${both}${request.note}${requestTo ? ` · to ${requestTo}` : ''}`);
  } else if (t && t.accSeparate) {
    hint.append(' · accessorials get their own email');
  }
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
  const blocked = rcBlock();
  if (blocked) {
    $('acc-error').textContent = blocked;
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

// ---- a template's questions: a pop-up before the create-email step

let quiz = null; // { template, at, path: [{ at, answer }], stopped } while the pop-up is open

function questionCount(t) {
  return t && t.questions ? t.questions.items.length : 0;
}

/** Fills in {company}: the template's name. (Templates had a Customer name for this until 0.18; it's no longer used.) */
function withCompany(text, t) {
  const company = String((t && t.name) || '').trim();
  return String(text || '').replace(/\{company\}/gi, company);
}

function askQuestions(t) {
  quiz = { template: t, at: 0, path: [], stopped: null };
  $('quiz').hidden = false;
  renderQuiz();
}

function renderQuiz() {
  const t = quiz.template;
  const items = t.questions.items;
  const intro = withCompany(t.questions.intro, t).trim();
  $('quiz-intro').hidden = !intro;
  $('quiz-intro-text').textContent = intro;
  const body = $('quiz-body');
  body.textContent = '';
  if (quiz.path.length) { // the answers so far, so going back makes sense
    const trail = el('ul', 'quiz-trail');
    for (const step of quiz.path) {
      const q = items[step.at];
      const line = el('li');
      line.append(el('span', '', withCompany(q.text, t)), ' ', el('b', '', q.answers[step.answer].label));
      trail.append(line);
    }
    body.append(trail);
  }
  if (quiz.stopped !== null) {
    const box = el('div', 'quiz-stop');
    box.setAttribute('role', 'alert');
    box.innerHTML = STOP_SVG;
    box.append(el('p', '', withCompany(quiz.stopped, t).trim() || "You can't continue with this answer."));
    const row = el('div', 'btn-row');
    const change = el('button', 'btn', 'Change my answer');
    change.type = 'button';
    change.addEventListener('click', quizBack);
    const close = el('button', 'btn primary', 'Close');
    close.type = 'button';
    close.id = 'quiz-close';
    close.addEventListener('click', cancelQuiz);
    row.append(change, close);
    body.append(box, row);
    $('quiz-foot').hidden = true;
    close.focus();
    return;
  }
  const q = items[quiz.at];
  body.append(el('p', 'quiz-q', withCompany(q.text, t)));
  const answers = el('div', 'quiz-answers');
  q.answers.forEach((a, j) => {
    const button = el('button', 'btn quiz-ans', a.label);
    button.type = 'button';
    button.addEventListener('click', () => answerQuiz(j));
    answers.append(button);
  });
  body.append(answers);
  $('quiz-foot').hidden = false;
  $('quiz-back').hidden = !quiz.path.length;
  const first = answers.querySelector('button');
  if (first) first.focus();
}

/** What an answer does: on to the next question, ahead to a later one, on to the email, or stop. */
function answerQuiz(j) {
  if (!quiz) return;
  const t = quiz.template;
  const items = t.questions.items;
  const a = items[quiz.at].answers[j];
  quiz.path.push({ at: quiz.at, answer: j });
  if (a.go === 'stop') {
    quiz.stopped = a.message || '';
    log(`Questions for "${t.name}": ${quizTrail()} -> stopped`);
    renderQuiz();
    return;
  }
  let next = -1;
  if (a.go === 'next') next = quiz.at + 1;
  else if (a.go !== 'continue') next = items.findIndex((q) => q.id === a.go);
  if (next > quiz.at && next < items.length) {
    quiz.at = next;
    renderQuiz();
    return;
  }
  log(`Questions for "${t.name}": ${quizTrail()} -> continue`);
  closeQuiz();
  openLoadForm(t);
}

function quizTrail() {
  const items = quiz.template.questions.items;
  return quiz.path.map((step) => `${items[step.at].text} ${items[step.at].answers[step.answer].label}`).join(' · ');
}

function quizBack() {
  if (!quiz || !quiz.path.length) return;
  quiz.at = quiz.path.pop().at;
  quiz.stopped = null;
  renderQuiz();
}

function closeQuiz() {
  $('quiz').hidden = true;
  $('quiz-body').textContent = '';
  quiz = null;
}

/** Cancel (or Close after a stop): no email. A stop's message stays in the status line as a reminder. */
function cancelQuiz() {
  if (!quiz) return;
  const { template, stopped } = quiz;
  closeQuiz();
  if (stopped !== null) {
    setStatus(`Stopped: ${withCompany(stopped, template).trim() || "this load's answers don't allow an email yet"}`, 'error');
  } else {
    log(`Questions for "${template.name}": cancelled`);
    setStatus('');
  }
  $('btn-email').focus();
}

/** Keeps Tab inside the pop-up while it's open. */
function onQuizKeys(event) {
  if (event.key !== 'Tab') return;
  const buttons = [...$('quiz').querySelectorAll('button')].filter((b) => !b.disabled && b.offsetParent !== null);
  if (!buttons.length) return;
  const first = buttons[0];
  const last = buttons[buttons.length - 1];
  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault();
    first.focus();
  }
}

// ---- creating the email

/**
 * Builds the PDF (one per stop with Multi-stop on, otherwise one, split into -1, -2… over 10 MB),
 * then opens the new email. With the team's Azure function set up, the files are uploaded so
 * Outlook attaches them itself; otherwise (or if the upload fails) they wait for the new email's
 * Attach PDF button.
 */
async function createEmail(template, load, acc) {
  const base = emailFileBase(template, load);
  const layout = layoutOf(template);
  const lumperStem = lumperBase(template, load);
  const separate = separateAccessorials(template, acc);
  const result = await withBusy(() => makeFiles(base, layout, lumperStem));
  if (!result) return;

  // The emails to open, each with its files: the usual one, or with the accessorials in a separate
  // email, the PODs in one and the request (with the PODs and receipts) in the other
  const emails = [];
  if (!separate) {
    emails.push({ form: emailFields(template, load, acc), files: result.files });
  } else {
    const mix = receiptMix();
    let podFiles = result.files;
    if (mix.receipts && layout.splitLumpers) podFiles = result.files.filter((f) => !f.lumper);
    else if (mix.receipts && mix.pods) {
      const pods = await withBusy(() => makeFiles(base, { ...layout, noReceipts: true }, lumperStem));
      if (!pods) return;
      podFiles = pods.files;
    }
    // Only receipts checked (a lumper request on its own): just the request, no empty POD email
    if (mix.pods) emails.push({ form: emailFields(template, load, acc, false), files: podFiles });
    emails.push({ form: accessorialFields(template, load, acc), files: result.files });
  }
  const namesOf = (files) => files.map((f) => f.name).join(', ');
  // The request's file has the same name as the POD email's when the receipts go in it, so say so
  const both = separate && !layout.splitLumpers && receiptMix().receipts ? ' (PODs and receipts)' : '';
  const toOf = (form) => (form.toRecipients || []).join(', ');
  const allFiles = [...new Set(emails.flatMap((e) => e.files))];

  let links = null; // file -> { type, name, url } once uploaded
  let azureProblem = '';
  const azure = teamAzure();
  if (azure && allFiles.length) {
    links = await withBusy(async () => {
      try {
        const list = new Map();
        for (let i = 0; i < allFiles.length; i++) {
          const file = allFiles[i];
          setStatus(`Uploading ${allFiles.length > 1 ? `${i + 1} of ${allFiles.length}: ` : ''}${file.name}…`);
          list.set(file, { type: 'file', name: file.name, url: await uploadToAzure(azure, file.bytes, file.name) });
        }
        log(`Azure: uploaded ${namesOf(allFiles)}`);
        return list;
      } catch (error) {
        azureProblem = error.userMessage || error.message;
        log(`Azure: upload failed (${azureProblem})`);
        return null;
      }
    });
  }

  // For Attach PDF in each new email (it finds its files by the subject)
  const stored = emails.map((e) => ({ subject: e.form.subject || '', files: e.files.map((f) => ({ name: f.name, bytes: f.bytes, pages: f.pages })) }));
  try {
    await idbSet(PENDING_KEY, { emails: stored, files: stored[0].files, subject: stored[0].subject, created: Date.now(), viaLink: !!links });
  } catch (error) {
    log(`New email: couldn't store the PDF (${error.name}: ${error.message})`);
    if (!links) {
      setStatus("Couldn't hand the PDF over to a new email here.", 'error');
      return;
    }
  }
  if (links) for (const e of emails) if (e.files.length) e.form.attachments = e.files.map((f) => links.get(f));
  // The sent with restack line (a team setting), under each email's message
  if (sentWithOn()) for (const e of emails) e.plain = addSentWith(e.form);

  let secondProblem = '';
  for (let i = 0; i < emails.length; i++) {
    if (i) await new Promise((resolve) => setTimeout(resolve, 800)); // a moment between the two, so Outlook opens both
    try {
      await openEmail(emails[i]);
    } catch (error) {
      log(`New email: couldn't open ${i ? 'the accessorial email' : 'one'} (${error.message})`);
      if (!i) {
        setStatus(`Couldn't open a new email: ${error.message}`, 'error');
        return;
      }
      secondProblem = error.message;
    }
  }

  countUsage('email', {
    ...usageFigures(allFiles),
    template: template ? template.name : '',
    templateId: template ? template.id : '',
    emails: emails.length - (secondProblem ? 1 : 0),
    accessorials: accessorialTypes(acc).join('/'),
    separate,
    azure: !!links,
  });
  const summary = accessorialSummary(acc);
  const noLumper = acc && !acc.lum.on && template && template.noLumper ? `No ${word('lumper')}` : '';
  for (const e of emails) {
    log(`New email: opened "${e.form.subject}" (${template ? `template "${template.name}"` : 'blank'}; ${namesOf(e.files) || 'no files'}${toOf(e.form) ? `; to ${toOf(e.form)}` : ''}${links ? '; attached through Azure' : ''}${e.plain ? '; sent with restack line' : ''})`);
  }
  if (summary || noLumper) log(`New email: ${[summary, noLumper].filter(Boolean).join(', ')}${separate ? ' (accessorials in their own email)' : ''}`);
  const tail = `${oversizeNote(allFiles)}${leftOutText(result)}`;
  const why = azureProblem ? ` (Couldn't attach automatically: ${azureProblem}.)` : '';
  const tone = azureProblem || secondProblem ? 'error' : 'ok';

  if (emails.length === 1) {
    const [only] = emails;
    const to = toOf(only.form);
    // (with the accessorials in a separate email, this is the request: only receipts were checked)
    const what = separate ? `: ${only.form.subject}` : '';
    const extras = separate ? summary : [summary, noLumper].filter(Boolean).join(', ');
    const extra = `${extras ? ` In the message: ${extras}.` : ''}${tail}`;
    if (links) setStatus(`New email opened${what}${to ? ` to ${to}` : ''} with ${namesOf(only.files)} attached.${extra}`, 'ok');
    else setStatus(`New email opened${what}${to ? ` to ${to}` : ''}. Click ${word('attachPdf')} on its ribbon to add ${namesOf(only.files)}.${why}${extra}`, tone);
    return;
  }

  const [pod, request] = emails;
  const describe = (e) => `${e.form.subject}${toOf(e.form) ? ` to ${toOf(e.form)}` : ''}`;
  const extra = ` In the request: ${summary}.${noLumper ? ` The POD email says ${noLumper}.` : ''}${tail}`;
  let text;
  if (secondProblem) {
    text = `Opened ${describe(pod)}, but couldn't open the accessorial email: ${secondProblem}.${links ? '' : ` Click ${word('attachPdf')} on the POD email's ribbon to add ${namesOf(pod.files) || 'its files'}.`}`;
  } else if (links) {
    text = `Opened 2 emails: ${describe(pod)} with ${namesOf(pod.files) || 'nothing'} attached, and ${describe(request)} with ${namesOf(request.files)}${both} attached.`;
  } else {
    text = `Opened 2 emails: ${describe(pod)}, and ${describe(request)}. Click ${word('attachPdf')} on each one's ribbon: the POD email gets ${namesOf(pod.files) || 'nothing'}, the request gets ${namesOf(request.files)}${both}.${why}`;
  }
  setStatus(`${text}${extra}`, tone);
  showReopen(request.form, !!secondProblem);
}

/** Whether new emails get the sent with restack line: the team's setting (off to start). */
function sentWithOn() {
  const settings = teamSettings();
  return !!(settings && settings.sentWith);
}

/**
 * Puts the sent with restack line under a new email's message: the picture, which Outlook attaches
 * inside the email from this site, after a blank line. Returns the fields as they were, to open the
 * email without it if Outlook can't add the picture.
 */
function addSentWith(form) {
  const plain = { ...form };
  if (form.attachments) plain.attachments = form.attachments.slice();
  const { file, width, height, alt } = SENT_WITH;
  const picture = `<img src="cid:${file}" width="${width}" height="${height}" alt="${alt}" style="display: block; width: ${width}px; height: ${height}px; border: 0;">`;
  form.htmlBody = `${form.htmlBody || ''}<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 11pt;"><br>${picture}</div>`;
  form.attachments = [...(form.attachments || []), { type: 'file', name: file, url: new URL(`${file}?v=${VERSION}`, window.location.href).href, isInline: true }];
  return plain;
}

/** Opens one of Create email's emails. If Outlook can't add the sent with restack line, it opens without it. */
async function openEmail(email) {
  try {
    await openNewEmail(email.form);
  } catch (error) {
    if (!email.plain) throw error;
    log(`New email: Outlook couldn't add the sent with restack line (${error.message}), so it opens without it`);
    email.form = email.plain;
    email.plain = null;
    await openNewEmail(email.form);
  }
}

/** Opens a new email in Outlook with these fields. */
async function openNewEmail(form) {
  const mailbox = Office.context.mailbox;
  if (typeof mailbox.displayNewMessageFormAsync === 'function') {
    await call((done) => mailbox.displayNewMessageFormAsync(form, done));
  } else {
    mailbox.displayNewMessageForm(form);
  }
}

/** Under the status: a button for the accessorial email, in case Outlook didn't show it. */
function showReopen(form, failed) {
  const box = $('parts');
  box.textContent = '';
  box.append(el('span', 'muted', failed ? '' : "Didn't see the accessorial email?"));
  const button = el('button', 'link-btn', failed ? 'Try opening the accessorial email again' : 'Open it again');
  button.type = 'button';
  button.addEventListener('click', () => {
    openNewEmail(form)
      .then(() => log(`New email: opened "${form.subject}" again`))
      .catch((error) => setStatus(`Couldn't open the accessorial email: ${error.message}`, 'error'));
  });
  box.append(button);
  box.hidden = false;
}

// ---- splitting by size

function mbText(bytes) {
  return `${(bytes / MB).toFixed(1)} MB`;
}

/** Copies pages [from, to) of a loaded PDF into a new PDF and returns its bytes. */
async function pagesToPdf(src, from, to) {
  const out = await PDFLib.PDFDocument.create();
  const pages = await out.copyPages(src, Array.from({ length: to - from }, (_, i) => from + i));
  pages.forEach((page) => out.addPage(page));
  if (state.topic) out.setTitle(state.topic);
  out.setCreator('Restack');
  out.setProducer(`Restack ${VERSION}`);
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

/** The files to send or save: base.pdf, or base-1.pdf, base-2.pdf… when it's over 10 MB. */
async function splitIntoFiles(bytes, base) {
  if (bytes.length <= SPLIT_LIMIT) return [{ name: `${base}.pdf`, bytes, pages: null, over: false }];
  const parts = await splitPdfBySize(bytes, SPLIT_LIMIT);
  if (parts.length === 1) return [{ name: `${base}.pdf`, bytes: parts[0].bytes, pages: parts[0].pages, over: parts[0].over }];
  const files = parts.map((part, i) => ({ name: `${base}-${i + 1}.pdf`, bytes: part.bytes, pages: part.pages, over: part.over }));
  log(`Split (over ${SPLIT_MB} MB): ${files.map((f) => `${f.name} ${mbText(f.bytes.length)} (${plural(f.pages, 'page')})`).join(', ')}`);
  return files;
}

function oversizeNote(files) {
  const over = files.filter((f) => f.over);
  if (!over.length) return '';
  const why = over[0].stop || over[0].lumper ? ' (stop and receipt files aren\'t split by size)' : ', because a single page is that big';
  return ` ${over.map((f) => `${f.name} is ${mbText(f.bytes.length)}`).join('; ')}: over ${SPLIT_MB} MB${why}.`;
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
  const source = teamSettings();
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

/**
 * A template's receipt file name: what it says, or {load} Receipt. Templates saved before 0.16 that
 * were still on the old default ({load} Lumper) get the new one; one saved since keeps whatever was typed.
 */
function receiptName(t) {
  const name = String((t && t.lumperName) || '').trim();
  if (!name || (name === OLD_RECEIPT_NAME && !(t && t.receiptNamed))) return NEW_TEMPLATE.lumperName;
  return name;
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
      stopFiles: t.stopFiles !== false,             // Multi-stop: each stop its own file (else all stops in one)
      lumperSeparate: !!t.lumperSeparate,           // lumper receipts in their own file
      lumperName: receiptName(t),
      receiptNamed: !!t.receiptNamed,             // saved in 0.16 or later: its receipt name is what was typed
      noLumper: !!t.noLumper,                       // say "No Lumper" when Lumper isn't ticked
      accSeparate: !!t.accSeparate,                 // accessorials in a separate email, when one is ticked
      accTo: String(t.accTo || ''),                 // that email's To (blank: the POD email's) and CC (blank: none)
      accCc: String(t.accCc || ''),
      accBody: String(t.accBody || ''),             // its message (blank: WORDS.accessorialIntro)
      questions: cleanQuestions(t.questions),       // asked in a pop-up before the create-email step
      // (a checklist and a customer name saved before 0.18 are dropped here: neither is used any more)
    }));
}

/**
 * A template's questions: { intro, items: [{ id, text, answers: [{ label, go, message }] }] }, or
 * null. An answer's `go` is 'next', 'continue', 'stop' (with a message), or the id of a LATER
 * question, so the questions always come to an end.
 */
function cleanQuestions(value) {
  if (!value || typeof value !== 'object' || !Array.isArray(value.items)) return null;
  const items = value.items
    .filter((q) => q && typeof q === 'object')
    .slice(0, QUIZ.maxQuestions)
    .map((q) => ({
      id: /^[A-Za-z0-9_-]{1,24}$/.test(String(q.id || '')) ? String(q.id) : newId(),
      text: oneLine(q.text).slice(0, 200),
      answers: (Array.isArray(q.answers) ? q.answers : [])
        .filter((a) => a && typeof a === 'object' && oneLine(a.label))
        .slice(0, QUIZ.maxAnswers)
        .map((a) => ({ label: oneLine(a.label).slice(0, 60), go: String(a.go || 'next'), message: oneLine(a.message).slice(0, 300) })),
    }))
    .filter((q) => q.text && q.answers.length);
  const seen = new Set();
  for (const q of items) {
    while (seen.has(q.id)) q.id = newId();
    seen.add(q.id);
  }
  items.forEach((q, i) => {
    for (const a of q.answers) {
      if (!['next', 'continue', 'stop'].includes(a.go) && !(items.findIndex((x) => x.id === a.go) > i)) a.go = 'next';
      if (a.go !== 'stop') a.message = '';
    }
  });
  return items.length ? { intro: oneLine(value.intro).slice(0, 300), items } : null;
}

/** Text with its line breaks and runs of spaces turned into single spaces. */
function oneLine(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
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
    const raw = localStorage.getItem(`${STORE_PREFIX}${name}`);
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
  if (value === null) localStorage.removeItem(`${STORE_PREFIX}${name}`);
  else localStorage.setItem(`${STORE_PREFIX}${name}`, JSON.stringify(value));
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
    repTemplates: !!(data && data.repTemplates), // reps may make their own templates (off to start)
    rateCon: cleanRateCon(data && data.rateCon),  // the sample rate confirmation (null: check is off)
    receiptWords: cleanReceiptWords(data && data.receiptWords), // file names that start out tagged Receipt (null: lumper, receipt)
    usage: cleanUsage(data && data.usage),       // the Restack Usage function (null: nothing is counted)
    sentWith: !!(data && data.sentWith),         // the sent with restack line on every email (off to start)
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
      response = await fetch(new URL(`${TEAM_FILE}?t=${Date.now()}`, SITE_ROOT).href, { cache: 'no-store' });
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
  if (CONSOLE) return team.status === 'ok'; // only someone with the team key gets this far
  if (team.status === 'nofile') return true;
  return team.status === 'ok' && !!team.me && team.published.admins.includes(team.me);
}

/** The team settings in force for this person: the admin's unpublished changes, else what's published. */
function teamSettings() {
  return (isAdmin() && team.draft) || team.published || null;
}

/** Making your own templates: always for the admin; for reps only once the admin turns it on. */
function canMakeOwn() {
  if (isAdmin()) return true;
  const settings = teamSettings();
  return !!(settings && settings.repTemplates);
}

function teamTemplates() {
  if (isAdmin() && team.draft) return team.draft.templates;
  return team.published ? team.published.templates : [];
}

function showLocked() {
  const site = SITE_ROOT;
  const [title, body] = {
    nokey: ["This copy of Restack isn't set up for a team.",
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

let templatesOpener = null; // the button that opened Manage templates (Email or Settings), for focus after

function openTemplates() {
  closeEmailMenu(false);
  if (!canMakeOwn()) return;
  editing = null;
  $('templates').hidden = false;
  renderTemplateSheet();
  $('tpl-back').focus();
}

function closeTemplates() {
  $('templates').hidden = true;
  editing = null;
  if (!$('read-view').hidden) renderList(); // a template may have turned lumper receipt tags on or off
  const back = templatesOpener || $('btn-email');
  templatesOpener = null;
  back.focus();
}

function cancelEdit() {
  editing = null;
  formSnapshot = '';
  renderTemplateSheet();
}

function renderTemplateSheet() {
  const formOpen = !!editing;
  $('tpl-list-view').hidden = formOpen && !CONSOLE; // the console keeps the list next to the form
  $('tpl-form').hidden = !formOpen;
  $('tpl-pick').hidden = !CONSOLE || formOpen;
  $('con-tpl-tools').hidden = !CONSOLE;
  if (formOpen && !CONSOLE) return;
  if (!formOpen) $('tpl-title').textContent = 'Email templates';
  const admin = isAdmin();
  const teamList = teamTemplates();
  $('team-note').textContent = CONSOLE ? '' : admin ? 'You manage these' : 'Set by your team admin';
  $('team-count').textContent = CONSOLE && teamList.length ? `(${teamList.length})` : '';
  const query = CONSOLE ? consoleSearch.trim().toLowerCase() : '';
  const shown = query ? teamList.filter((t) => [t.name, t.to, t.cc, t.accTo].some((v) => String(v || '').toLowerCase().includes(query))) : teamList;
  fillTemplateList($('team-list'), byName(shown), admin ? 'team' : null);
  $('team-empty').hidden = shown.length > 0;
  $('team-empty').textContent = query
    ? `No templates match "${consoleSearch.trim()}".`
    : admin
      ? 'Add the templates your reps should use, then publish them.'
      : "Your team admin hasn't added any templates yet.";
  $('team-admin').hidden = !admin;
  if (admin) renderPublishBox();
  $('mine-section').hidden = !canMakeOwn();
  fillTemplateList($('mine-list'), byName(state.myTemplates), 'mine');
  $('mine-empty').hidden = state.myTemplates.length > 0;
}

/** The rate confirmation check's status, for the team admin. */
function renderRcBox(message) {
  const sample = teamRateCon();
  teamSubtitle('rc', sample ? 'On' : 'Off');
  $('rc-remove').hidden = !sample;
  $('rc-choose').hidden = !CONSOLE;
  $('rc-choose').textContent = sample ? 'Use a different PDF…' : 'Choose a rate con PDF…';
  const saved = sample && sample.saved ? new Date(sample.saved).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : '';
  $('rc-text').textContent = message || (sample
    ? `Checking against ${sample.name} (${plural(sample.pages.length, 'page')}${saved ? `, set ${saved}` : ''}${sample.pages.some((page) => page.words) ? '' : '; pictures of pages, so it\'s matched by layout'}). Pages that look like it are flagged, as are pages that say "Rate Confirmation" or "Load Confirmation" and pages with the same barcode as a flagged one. They have to be left out before anyone emails or saves.`
    : CONSOLE
      ? 'Off. To turn it on, choose one of your rate confirmations (a PDF) as the sample. Its first 3 pages are kept.'
      : `Off. To turn it on, open one of your rate confirmations from your inbox, preview it, and click ${word('useAsRcSample')} at the bottom of the preview.`);
}

async function onRcRemove() {
  try {
    await saveTeamDraft({ rateCon: null });
  } catch (error) {
    renderRcBox(`Couldn't remove it: ${error.message}`);
    return;
  }
  log('Rate con check: sample removed (not published yet)');
  renderPublishBox();
  renderRcBox(`Removed. ${publishWhere()} to turn the check off for your team.`);
  recheckRateCons();
}

/** The admin's switch for the sent with restack line on every email Restack creates. Published with the templates. */
async function onSentWithToggle(on) {
  const id = CONSOLE ? 'con-sent-with' : 'team-sent-with';
  markSwitch(id, on);
  try {
    await saveTeamDraft({ sentWith: on });
  } catch (error) {
    markSwitch(id, !on);
    teamNote(`Couldn't save it: ${error.message}`, 'error');
    return;
  }
  log(`Team: sent with restack line ${on ? 'on' : 'off'} (not published yet)`);
  renderPublishBox();
  if (CONSOLE) {
    teamNote(on ? 'Every email Restack creates gets the sent with restack line once you publish.' : 'Emails stop getting the sent with restack line once you publish.');
  } else {
    teamNote(on
      ? "Your emails get the sent with restack line now. Your reps' get it once you publish."
      : "Your emails don't get the sent with restack line now. Your reps' stop getting it once you publish.");
  }
}

/** The admin's switch for letting reps make their own templates. Published with the team templates. */
async function onRepTemplatesToggle(on) {
  markSwitch('team-rep-templates', on);
  try {
    await saveTeamDraft({ repTemplates: on });
  } catch (error) {
    markSwitch('team-rep-templates', !on);
    teamNote(`Couldn't save it: ${error.message}`, 'error');
    return;
  }
  log(`Team: reps ${on ? 'can' : "can't"} make their own templates (not published yet)`);
  renderPublishBox();
  teamNote(`Reps ${on ? 'can' : "can't"} make their own templates once you publish.`);
}

function fillTemplateList(list, templates, scope) {
  list.textContent = '';
  for (const t of templates) {
    const row = el('li', 'tpl-row');
    const text = el('div', 'tpl-text');
    text.append(el('div', 'tpl-name', t.name), el('div', 'tpl-sub', templateSub(t)));
    row.append(text);
    const open = !!(editing && editing.scope === scope && editing.id === t.id);
    row.classList.toggle('editing', open);
    if (open) row.setAttribute('aria-current', 'true');
    if (scope) {
      const edit = el('button', 'btn', open ? 'Editing' : 'Edit');
      edit.type = 'button';
      edit.setAttribute('aria-label', `Edit ${t.name}`);
      edit.addEventListener('click', () => openTemplateEditor(scope, t.id));
      row.append(edit);
      if (CONSOLE) row.addEventListener('click', (event) => {
        if (!event.target.closest('button')) openTemplateEditor(scope, t.id);
      });
    }
    list.append(row);
  }
}

/** Whether there's something to publish: the Team group in Settings, and a banner in Manage templates. */
function renderPublishBox(message) {
  renderTeamMenu();
  const box = $('tpl-publish-box');
  if (!box) return;
  const pending = !!team.draft && !team.justPublished;
  let text = '';
  if (message) text = message;
  else if (team.justPublished) text = PUBLISHED_TEXT;
  else if (pending) text = "You have changes your team can't see yet. Publish, then upload team.json to GitHub.";
  box.hidden = !text || !isAdmin();
  box.classList.toggle('pending', pending && !message);
  $('tpl-publish-text').textContent = text;
  $('tpl-publish').hidden = !pending;
}

/** Opens a template in the editor (in the console, after asking about unsaved changes to another). */
function openTemplateEditor(scope, id) {
  if (CONSOLE && editing && !(editing.scope === scope && editing.id === id) && templateFormChanged()
    && !window.confirm("You haven't saved your changes to this template. Leave them?")) return;
  editTemplate(scope, id);
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
  $('tpl-no-lumper').checked = !!(t && t.noLumper);
  $('tpl-stop-files').checked = !!t && t.stopFiles !== false; // a new template starts with every box unticked
  $('tpl-lumper-sep').checked = !!(t && t.lumperSeparate);
  $('tpl-lumper-name').value = receiptName(t);
  syncLumperNameRow();
  $('tpl-acc-sep').checked = !!(t && t.accSeparate);
  $('tpl-acc-to').value = t ? t.accTo || '' : '';
  $('tpl-acc-cc').value = t ? t.accCc || '' : '';
  $('tpl-acc-message').value = t ? t.accBody || '' : '';
  $('tpl-acc-message').placeholder = word('accessorialIntro');
  syncAccRow();
  qDraft = t && t.questions ? JSON.parse(JSON.stringify(t.questions)) : { intro: '', items: [] };
  renderQuestionBuilder();
  $('tpl-delete').hidden = !t;
  $('tpl-delete').textContent = 'Delete this template';
  $('tpl-duplicate').hidden = !t;
  $('tpl-error').textContent = '';
  renderTemplateSheet();
  fitQuestionBoxes(); // now that the form is showing, the question boxes can be measured
  setTemplateTitle(t ? t.name : '');
  formSnapshot = templateFormState();
  $('tpl-name').focus();
}

function setTemplateTitle(name) {
  const scope = editing ? editing.scope : 'team';
  const kind = scope === 'team' ? 'team template' : 'template';
  $('tpl-title').textContent = `${editing && editing.id ? 'Edit' : 'New'} ${kind}`;
  $('tpl-form-title').hidden = !CONSOLE;
  $('tpl-form-title').textContent = editing && editing.id ? name : `New ${kind}`;
}

/** Make a copy: the form becomes a new template with the same settings, named "… (copy)". */
function onTemplateDuplicate() {
  if (!editing || !editing.id) return;
  const name = $('tpl-name').value.trim();
  editing = { scope: editing.scope, id: null };
  $('tpl-name').value = `${name} (copy)`.slice(0, 60);
  $('tpl-delete').hidden = true;
  $('tpl-duplicate').hidden = true;
  $('tpl-error').textContent = '';
  if (qDraft && Array.isArray(qDraft.items)) qDraft = JSON.parse(JSON.stringify(qDraft));
  renderTemplateSheet();
  setTemplateTitle('');
  formSnapshot = 'copy'; // not saved yet: leaving it asks first
  $('tpl-name').focus();
  $('tpl-name').select();
  log(`Templates: copying "${name}"`);
}

/** Shows the lumper file name box only when lumper receipts get their own file. */
function syncLumperNameRow() {
  $('tpl-lumper-row').hidden = !$('tpl-lumper-sep').checked;
}

/** Shows the accessorial email's To, CC and message only when the template sends one. */
function syncAccRow() {
  $('tpl-acc-row').hidden = !$('tpl-acc-sep').checked;
}

// ---- the question builder (in the template form)

let qDraft = { intro: '', items: [] }; // the questions of the template being edited

function clearTplError() {
  $('tpl-error').textContent = '';
}

function iconButton(svg, label, onClick) {
  const button = el('button', 'icon-btn');
  button.type = 'button';
  button.innerHTML = svg;
  button.title = label;
  button.setAttribute('aria-label', label);
  button.addEventListener('click', onClick);
  return button;
}

/** [value, text] choices for what an answer to question i does. Jumps only go forward. */
function goChoices(i) {
  const items = qDraft.items;
  const choices = [['next', i < items.length - 1 ? `Next question (Q${i + 2})` : 'Next question (or the email)']];
  for (let k = i + 2; k < items.length; k++) choices.push([items[k].id, `Go to Q${k + 1}`]);
  choices.push(['continue', 'Continue to the email'], ['stop', 'Stop with a message']);
  return choices;
}

/** Redraws the builder from qDraft (after adding or removing something), focusing `focus` if given. */
function renderQuestionBuilder(focus) {
  const items = qDraft.items;
  $('q-intro-row').hidden = !items.length;
  $('q-intro').value = qDraft.intro;
  const list = $('q-list');
  list.textContent = '';
  items.forEach((q, i) => list.append(questionEditor(q, i)));
  $('q-add').disabled = items.length >= QUIZ.maxQuestions;
  $('q-add').textContent = items.length ? 'Add another question' : 'Add a question';
  fitQuestionBoxes();
  const target = focus ? document.querySelector(focus) : null;
  if (target) target.focus();
}

/** Grows a builder text box to fit what's in it (one line to start). Hidden boxes wait until shown. */
function fitHeight(area) {
  if (!area.offsetParent) return;
  area.style.height = 'auto';
  area.style.height = `${area.scrollHeight + 2}px`;
}

function fitQuestionBoxes() {
  for (const area of document.querySelectorAll('#q-list textarea')) fitHeight(area);
}

function questionEditor(q, i) {
  const item = el('li', 'q-item');
  item.dataset.qid = q.id;
  const head = el('div', 'q-item-head');
  const text = document.createElement('textarea');
  text.className = 'q-text';
  text.rows = 1;
  text.maxLength = 200;
  text.value = q.text;
  text.placeholder = i === 0 ? 'Did the driver sign the POD?' : 'Type the question';
  text.setAttribute('aria-label', `Question ${i + 1}`);
  text.addEventListener('input', () => {
    q.text = text.value;
    fitHeight(text);
    clearTplError();
  });
  head.append(el('span', 'q-num', `Q${i + 1}`), text, iconButton(X_SVG, `Remove question ${i + 1}`, () => removeQuestion(i)));
  const answers = el('div', 'q-answers');
  q.answers.forEach((a, j) => answers.append(answerEditor(q, i, a, j)));
  const add = el('button', 'link-btn q-ans-add', 'Add an answer');
  add.type = 'button';
  add.disabled = q.answers.length >= QUIZ.maxAnswers;
  add.addEventListener('click', () => {
    q.answers.push({ label: '', go: 'next', message: '' });
    renderQuestionBuilder(`[data-qid="${q.id}"] .q-ans:last-child .q-ans-label`);
  });
  item.append(head, answers, add);
  return item;
}

function answerEditor(q, i, a, j) {
  const row = el('div', 'q-ans');
  const label = document.createElement('input');
  label.type = 'text';
  label.className = 'q-ans-label';
  label.maxLength = 60;
  label.value = a.label;
  label.placeholder = 'Answer';
  label.setAttribute('aria-label', `Question ${i + 1}, answer ${j + 1}`);
  label.addEventListener('input', () => {
    a.label = label.value;
    clearTplError();
  });
  const go = document.createElement('select');
  go.className = 'q-ans-go';
  go.setAttribute('aria-label', `What answer ${j + 1} to question ${i + 1} does`);
  const choices = goChoices(i);
  for (const [value, text] of choices) go.append(new Option(text, value));
  if (!choices.some(([value]) => value === a.go)) a.go = 'next'; // its question was removed
  go.value = a.go;
  const message = document.createElement('textarea');
  message.className = 'q-ans-msg';
  message.rows = 1;
  message.maxLength = 300;
  message.value = a.message || '';
  message.placeholder = 'Message to show, e.g. Please get the exit slip from the carrier before continuing';
  message.setAttribute('aria-label', `Stop message for answer ${j + 1} to question ${i + 1}`);
  message.hidden = a.go !== 'stop';
  message.addEventListener('input', () => {
    a.message = message.value;
    fitHeight(message);
    clearTplError();
  });
  go.addEventListener('change', () => {
    a.go = go.value;
    message.hidden = a.go !== 'stop';
    if (!message.hidden) {
      fitHeight(message);
      message.focus();
    }
    clearTplError();
  });
  const remove = iconButton(X_SVG, `Remove answer ${j + 1} from question ${i + 1}`, () => {
    q.answers.splice(j, 1);
    renderQuestionBuilder(`[data-qid="${q.id}"] .q-ans-add`);
  });
  row.append(label, go, remove, message);
  return row;
}

function addQuestion() {
  if (qDraft.items.length >= QUIZ.maxQuestions) return;
  if (!qDraft.items.length && !qDraft.intro.trim()) qDraft.intro = QUIZ.intro;
  const q = { id: newId(), text: '', answers: [{ label: 'Yes', go: 'next', message: '' }, { label: 'No', go: 'stop', message: '' }] };
  qDraft.items.push(q);
  renderQuestionBuilder(`[data-qid="${q.id}"] .q-text`);
}

/** Removing a question: answers that jumped to it go to the next question instead. */
function removeQuestion(i) {
  const [gone] = qDraft.items.splice(i, 1);
  for (const q of qDraft.items) for (const a of q.answers) if (a.go === gone.id) a.go = 'next';
  const near = qDraft.items[i] || qDraft.items[i - 1];
  renderQuestionBuilder(near ? `[data-qid="${near.id}"] .q-text` : '#q-add');
  clearTplError();
}

/** What's wrong with the questions being edited, or ''. Blank answers are simply left out. */
function questionsProblem(draft) {
  for (let i = 0; i < draft.items.length; i++) {
    const q = draft.items[i];
    const n = `Q${i + 1}`;
    if (!q.text.trim()) return `${n}: type the question, or remove it.`;
    const answers = q.answers.filter((a) => a.label.trim());
    if (!answers.length) return `${n}: add at least one answer.`;
    const seen = new Set();
    for (const a of answers) {
      const name = a.label.trim();
      if (seen.has(name.toLowerCase())) return `${n}: two answers are called "${name}".`;
      seen.add(name.toLowerCase());
      if (a.go === 'stop' && !String(a.message || '').trim()) return `${n}: type the message to show when someone answers "${name}".`;
    }
  }
  return '';
}

function templateProblem(t) {
  if (!t.name) return 'Give the template a name.';
  if (t.lumperSeparate && fill(t.lumperName, 'X').toLowerCase() === fill(t.file || '{load}', 'X').toLowerCase()) {
    return 'The receipt file needs a different name from the PDF name.';
  }
  for (const [label, text] of [['To', t.to], ['CC', t.cc], ['Accessorial email To', t.accTo], ['Accessorial email CC', t.accCc]]) {
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

/** Saves the admin's unpublished team changes: { templates }, { azure }, { rateCon } and so on. The rest stays as it is. */
async function saveTeamDraft(changes) {
  const base = team.draft || team.published || cleanTeam({});
  const kept = base.admins && base.admins.length ? base.admins : [team.me].filter(Boolean);
  const admins = Array.isArray(changes.admins) && changes.admins.length ? changes.admins : kept; // Team admins changes them
  const draft = { ...cleanTeam(base), ...changes, v: 1, updated: Date.now(), by: team.me, admins };
  await writeRoaming(DRAFT_KEY, draft);
  team.draft = draft;
  team.justPublished = false;
  renderTeamMenu();
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
    noLumper: $('tpl-no-lumper').checked,
    accSeparate: $('tpl-acc-sep').checked,
    accTo: $('tpl-acc-to').value.trim(),
    accCc: $('tpl-acc-cc').value.trim(),
    accBody: $('tpl-acc-message').value.replace(/\s+$/, ''),
    stopFiles: $('tpl-stop-files').checked,
    lumperSeparate: $('tpl-lumper-sep').checked,
    lumperName: $('tpl-lumper-name').value.trim().replace(/\.pdf$/i, '') || NEW_TEMPLATE.lumperName,
    receiptNamed: true,
    questions: null,
  };
  const problem = templateProblem(data) || questionsProblem(qDraft);
  if (problem) {
    $('tpl-error').textContent = problem;
    return;
  }
  data.questions = cleanQuestions(qDraft);
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
  if (CONSOLE) {
    editing = { scope: editing.scope, id: data.id }; // stays open, now saved
    $('tpl-delete').hidden = false;
    $('tpl-delete').textContent = 'Delete this template';
    $('tpl-duplicate').hidden = false;
    setTemplateTitle(data.name);
    formSnapshot = templateFormState();
    renderTemplateSheet();
    consoleNote(`Saved "${data.name}". ${publishWhere()} so your reps get it.`, 'ok');
    return;
  }
  editing = null;
  formSnapshot = '';
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
  formSnapshot = '';
  renderTemplateSheet();
  if (CONSOLE) consoleNote(`Deleted. ${publishWhere()} to take it away from your reps too.`, 'ok');
}

const PUBLISHED_TEXT = 'team.json downloaded. Upload it to your GitHub repository (Add file → Upload files → Commit changes). If it saved as "team (1).json", rename it to team.json first. Reps get the changes the next time they open the panel.';

/** Encrypts the team templates and settings with the team key and downloads team.json to upload to GitHub. */
async function onPublish() {
  const source = team.draft || team.published;
  if (!source || !team.key) return;
  if (!team.me) {
    const why = "Couldn't read your email address from Outlook, so this can't be published.";
    renderPublishBox(why);
    teamNote(why, 'error');
    return;
  }
  const admins = source.admins.length ? source.admins : [team.me];
  const payload = { ...cleanTeam(source), v: 1, updated: source.updated || Date.now(), by: team.me, admins };
  const text = await encryptTeam(payload, team.key);
  if (githubReady()) {
    $('con-publish').disabled = true;
    consoleNote('Publishing to GitHub…');
    let where;
    try {
      where = await publishToGithub(text);
    } catch (error) {
      renderConsolePublish();
      consoleNote(`Couldn't publish to GitHub: ${error.userMessage || error.message}. Nothing changed for your team. (Team settings → Publishing has the token.)`, 'error');
      return;
    }
    team.justPublished = true;
    log(`Team: published ${plural(payload.templates.length, 'template')} to GitHub (${where.repo}, ${where.branch})`);
    renderPublishBox();
    consoleNote(`Published to ${where.repo}. GitHub updates your site within a minute or two, and reps get the changes the next time they open the panel.`, 'ok');
    return;
  }
  downloadText(text, TEAM_FILE, 'application/json');
  team.justPublished = true;
  log(`Team: published ${plural(payload.templates.length, 'template')} (team.json downloaded)`);
  renderPublishBox();
  teamNote(PUBLISHED_TEXT, 'ok');
}

// ---- the Team screen: one of the admin's settings at a time (from Settings → Team)

const TEAM_SCREENS = {
  admins: { section: 'admins-section', title: 'Team admins', render: () => renderAdminsBox() },
  rc: { section: 'rc-section', title: 'Rate confirmation check', render: () => renderRcBox() },
  azure: { section: 'azure-section', title: 'Attach automatically', render: () => renderAzureBox() },
  receipts: { section: 'receipts-section', title: 'Receipt file names', render: () => renderReceiptWords() },
  usage: { section: 'usage-section', title: 'Usage counts', render: () => renderUsageBox() },
};
let teamShowing = null;

function openTeamSheet(which) {
  const screen = TEAM_SCREENS[which];
  if (!screen || !isAdmin()) return;
  closeSettings(false);
  teamShowing = which;
  for (const [key, other] of Object.entries(TEAM_SCREENS)) $(other.section).hidden = key !== which;
  $('team-title').textContent = screen.title;
  $('team-subtitle').textContent = '';
  $('team-sheet').hidden = false;
  screen.render();
  $('team-back').focus();
}

function closeTeamSheet() {
  if ($('team-sheet').hidden) return;
  $('team-sheet').hidden = true;
  teamShowing = null;
  if (!$('read-view').hidden) renderList(); // receipt words may have changed which files start out tagged
  $('btn-settings').focus();
}

/** The line under the Team screen's title (On, Off…), when that setting is the one showing. */
function teamSubtitle(which, text) {
  if (teamShowing === which || CONSOLE) {
    const node = CONSOLE ? $(`${which}-state`) : $('team-subtitle');
    if (node) node.textContent = text;
  }
}

function setStatusIn(id, text, tone) {
  const node = $(id);
  node.textContent = text || '';
  node.className = `hint${tone ? ` ${tone}` : ''}`;
}

// ---- receipt file names (the team's words; each rep turns the tagging on in Settings)

const DEFAULT_RECEIPT_WORDS = ['lumper', 'receipt'];

function cleanReceiptWords(value) {
  if (!Array.isArray(value)) return null;
  const words = value.map((w) => String(w).toLowerCase().replace(/\s+/g, ' ').trim()).filter((w) => w && w.length <= 40);
  return [...new Set(words)].slice(0, 20);
}

/** The words that make a file start out tagged Receipt (the team's, else lumper and receipt). */
function receiptWords() {
  const settings = teamSettings();
  return settings && settings.receiptWords ? settings.receiptWords : DEFAULT_RECEIPT_WORDS;
}

function renderReceiptWords(message, tone) {
  $('rw-words').value = receiptWords().join(', ');
  teamSubtitle('receipts', receiptWords().length ? `${plural(receiptWords().length, 'word')}` : 'No words');
  setStatusIn('rw-status', message || '', tone);
}

async function onReceiptWordsSave(words) {
  const list = cleanReceiptWords(words) || [];
  try {
    await saveTeamDraft({ receiptWords: list });
  } catch (error) {
    setStatusIn('rw-status', `Couldn't save them: ${error.message}`, 'error');
    return;
  }
  log(`Receipt file names: ${list.join(', ') || 'none'} (not published yet)`);
  renderPublishBox();
  renderReceiptWords(list.length ? `Saved. ${publishWhere()} to share them with your reps.` : 'Saved with no words, so nothing is tagged by its name. Publish to share it.', 'ok');
  applyReceiptNames();
}

// ---- team admins: who gets Settings → Team (published with the templates)

/** The team's admins, as they'll be published: the admin's changes, else what's published, else you. */
function teamAdmins() {
  const settings = teamSettings();
  return settings && settings.admins && settings.admins.length ? settings.admins : [team.me].filter(Boolean);
}

/** The list, each with Remove (not for yourself, and never the last one), and the box to add one. */
function renderAdminsBox(message, tone) {
  const list = teamAdmins();
  teamSubtitle('admins', plural(list.length, 'admin'));
  const box = $('admin-list');
  box.textContent = '';
  for (const address of list) {
    const row = el('li', 'admin-row');
    row.append(el('span', 'admin-email', address));
    const you = !CONSOLE && address === team.me; // (the console doesn't know who's using it)
    if (you) row.append(el('span', 'admin-you', 'you'));
    const remove = el('button', 'link-btn', 'Remove');
    remove.type = 'button';
    remove.setAttribute('aria-label', `Remove ${address}`);
    remove.disabled = you || list.length < 2;
    remove.addEventListener('click', () => onAdminRemove(address).catch(fail));
    row.append(remove);
    box.append(row);
  }
  $('admin-self').hidden = CONSOLE;
  $('admin-new').disabled = list.length >= MAX_ADMINS;
  $('admin-add').disabled = list.length >= MAX_ADMINS;
  setStatusIn('admin-status', message || (list.length >= MAX_ADMINS ? `That's the most there can be (${MAX_ADMINS}).` : ''), tone);
}

async function onAdminAdd(event) {
  event.preventDefault();
  const input = $('admin-new');
  const address = input.value.trim().toLowerCase();
  const list = teamAdmins();
  let problem = '';
  if (!address) problem = 'Type their email address first.';
  else if (!/^[^\s@<>()",;]+@[^\s@<>()",;]+\.[^\s@<>()",;]+$/.test(address)) problem = `"${address}" doesn't look like an email address.`;
  else if (list.includes(address)) problem = `${address} is already an admin.`;
  else if (list.length >= MAX_ADMINS) problem = `That's the most there can be (${MAX_ADMINS}).`;
  if (problem) {
    renderAdminsBox(problem, 'error');
    input.focus();
    return;
  }
  try {
    await saveTeamDraft({ admins: [...list, address] });
  } catch (error) {
    renderAdminsBox(`Couldn't save it: ${error.message}`, 'error');
    return;
  }
  input.value = '';
  log(`Team admins: added one, ${list.length + 1} now (not published yet)`);
  renderPublishBox();
  renderAdminsBox(`Added. ${publishWhere()} so ${address} gets the admin settings.`, 'ok');
  input.focus();
}

async function onAdminRemove(address) {
  const list = teamAdmins();
  if (list.length < 2 || (!CONSOLE && address === team.me) || !list.includes(address)) return;
  try {
    await saveTeamDraft({ admins: list.filter((a) => a !== address) });
  } catch (error) {
    renderAdminsBox(`Couldn't remove them: ${error.message}`, 'error');
    return;
  }
  log(`Team admins: removed one, ${list.length - 1} now (not published yet)`);
  renderPublishBox();
  renderAdminsBox(`Removed. ${publishWhere()} to take the admin settings away from ${address}.`, 'ok');
  $('admin-new').focus();
}

// ---- usage counts: the Restack Usage function (its own Function App)

const USAGE_FUNCTION = 'usage'; // the function's name in the Restack Usage Function App

function cleanUsage(value) {
  if (!value || typeof value !== 'object') return null;
  const url = String(value.url || '').trim().replace(/\/+$/, '');
  const key = String(value.key || '').trim();
  return okAzureUrl(url) && key.length >= 16 ? { url, key } : null;
}

/** The team's Restack Usage function ({ url, key }), or null when nothing is counted. */
function teamUsage() {
  const settings = teamSettings();
  return settings ? settings.usage || null : null;
}

/** The function address as typed: the Function App's domain alone becomes https://<app>.azurewebsites.net/api/usage. */
function usageBaseUrl(text) {
  let url = String(text || '').trim().replace(/[?#].*$/, '').replace(/\/+$/, '');
  if (url && !/^[a-z][a-z0-9+.-]*:\/\//i.test(url) && /^[^\s/:]+\.[^\s/:]+(:\d+)?(\/\S*)?$/.test(url)) url = `https://${url}`;
  if (/^https?:\/\/[^/\s]+$/i.test(url)) url += `/api/${USAGE_FUNCTION}`;
  else if (/^https?:\/\/[^/\s]+\/api$/i.test(url)) url += `/${USAGE_FUNCTION}`;
  return url;
}

/** The admin key that reads everyone's counts: kept in the admin's own Outlook, never published. */
function usageAdminKey() {
  return String(readRoaming('usageAdmin') || '');
}

function renderUsageBox(message, tone) {
  const usage = teamUsage();
  $('us-url').value = usage ? usage.url : '';
  $('us-key').value = usage ? usage.key : '';
  $('us-admin').value = usageAdminKey();
  teamSubtitle('usage', usage ? 'On' : 'Off');
  setStatusIn('us-status', message || (usage
    ? 'Counting. The admin console\'s Usage tab shows the numbers.'
    : 'Not set up. Nothing is counted.'), tone);
}

function usageFromFields() {
  const typed = $('us-url').value.trim();
  const key = $('us-key').value.trim();
  const admin = $('us-admin').value.trim();
  if (!typed && !key) return { off: true, admin };
  const url = usageBaseUrl(typed);
  if (!okAzureUrl(url)) return { problem: "The function address should be your Restack Usage Function App's Default domain, like yourusageapp.azurewebsites.net." };
  if (key.length < 16) return { problem: 'The counting key should be at least 16 characters. Click Generate to make one.' };
  if (admin && admin.length < 16) return { problem: 'The admin key should be at least 16 characters. Click Generate to make one.' };
  if (admin && admin === key) return { problem: 'The admin key has to be different from the counting key.' };
  return { usage: { url, key }, admin };
}

async function copyFrom(input, statusId) {
  const value = $(input).value.trim();
  if (!value) return;
  try {
    await navigator.clipboard.writeText(value);
    setStatusIn(statusId, 'Copied.');
  } catch (error) {
    $(input).select();
    let copied = false;
    try {
      copied = document.execCommand('copy');
    } catch (ignored) {
      copied = false;
    }
    setStatusIn(statusId, copied ? 'Copied.' : 'Press Ctrl+C to copy it.');
  }
}

async function onUsageSave() {
  const { usage, admin, problem, off } = usageFromFields();
  if (problem) {
    setStatusIn('us-status', problem, 'error');
    return;
  }
  try {
    await saveTeamDraft({ usage: off ? null : usage });
    await writeRoaming('usageAdmin', admin || null);
  } catch (error) {
    setStatusIn('us-status', `Couldn't save it: ${error.message}`, 'error');
    return;
  }
  log(`Usage counts: settings ${off ? 'cleared' : 'saved'} (not published yet)`);
  renderPublishBox();
  usageReport.key = ''; // the Usage tab reads again
  renderUsageBox(off ? `Turned off. ${publishWhere()} to stop the counting.` : `Saved. ${publishWhere()} so your reps' panels start counting.`, 'ok');
}

/** Checks the function is there and set up, then that both keys work (nothing is counted). */
async function onUsageTest() {
  const { usage, admin, problem } = usageFromFields();
  if (problem || !usage) {
    setStatusIn('us-status', problem || 'Fill in the function address and counting key first.', 'error');
    return;
  }
  $('us-test').disabled = true;
  setStatusIn('us-status', 'Testing… (the first call can take a few seconds while Azure starts the function)');
  try {
    const info = await usageHealth(usage.url);
    if (!info.keySet) throw userError("RESTACK_USAGE_KEY isn't set in the Function App yet");
    if (info.storage === false) throw userError(info.storageError || "the function can't reach its storage");
    await sendUsage(usage, { check: true });
    let read = '';
    if (admin) {
      if (!info.adminKeySet) throw userError("RESTACK_USAGE_ADMIN_KEY isn't set in the Function App yet");
      const today = new Date().toISOString().slice(0, 10);
      await readUsage(usage.url, admin, today, today);
      read = ' The admin key reads the numbers.';
    }
    log(`Usage counts: test connection worked (${info.version || 'version unknown'})`);
    setStatusIn('us-status', `Connection works.${read} Click Save, then Publish.`, 'ok');
  } catch (error) {
    const reason = error.userMessage || error.message;
    log(`Usage counts: test failed (${reason})`);
    setStatusIn('us-status', `Didn't work: ${reason}.`, 'error');
  } finally {
    $('us-test').disabled = false;
  }
}

// ---- the admin console: the team's templates and settings in a browser window

function consoleUrl() {
  const parts = [`k=${encodeURIComponent(team.key || '')}`];
  const admin = usageAdminKey();
  if (admin) parts.push(`u=${encodeURIComponent(admin)}`);
  return `${new URL('console.html', window.location.href).href}#${parts.join('&')}`; // the beta opens the beta's console
}

let consoleArmed = 0;

async function openConsole() {
  const pending = !!team.draft && !team.justPublished;
  if (pending && Date.now() - consoleArmed > 8000) { // the console starts from what's published
    consoleArmed = Date.now();
    teamNote("You have changes here that aren't published. The console starts from what's published, so publish first to take them along. Click Open admin console again to open it anyway.", 'error');
    return;
  }
  consoleArmed = 0;
  const url = consoleUrl();
  log('Team: admin console');
  const opened = (how) => teamNote(`The admin console opened ${how}.${pending ? " It starts from what's published, so publish your changes here first if you want them there." : ''}`, pending ? 'error' : 'ok');
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.openBrowserWindow === 'function') {
      Office.context.ui.openBrowserWindow(url);
      opened('in your browser');
      return;
    }
  } catch (error) {
    // try a window below
  }
  const win = window.open(url, '_blank');
  if (win) {
    opened('in a new window');
    return;
  }
  try {
    await navigator.clipboard.writeText(url);
    teamNote("Outlook didn't open your browser, so the console's link is copied: paste it into your browser's address bar. It has your team key in it, so don't share it.", 'error');
  } catch (error) {
    teamNote(`Outlook didn't open your browser. Open ${SITE_ROOT}console.html from Settings → Team on your computer's Outlook instead.`, 'error');
  }
}

// ---- Azure settings (team admin)

function renderAzureBox(message, tone) {
  const azure = teamAzure();
  $('az-url').value = azure ? azure.url : '';
  $('az-key').value = azure ? azure.key : '';
  teamSubtitle('azure', azure ? 'On' : 'Off');
  setAzureStatus(message || (azure
    ? 'Create email attaches PDFs through your Azure function.'
    : `Not set up. ${word('createEmail')} uses the ${word('attachPdf')} button instead.`), tone);
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
    ? `Turned off. ${publishWhere()} to update your team.`
    : `Saved. ${publishWhere()} to share it with your reps.`, 'ok');
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
    doc.addPage([200, 200]).drawText('Restack test', { x: 20, y: 100, size: 12 });
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

/** The add-in's manifest template, from the main site (the beta reads the same one). */
async function manifestTemplate() {
  const response = await fetch(new URL('manifest.template.xml', SITE_ROOT).href, { cache: 'no-store' });
  if (!response.ok) throw new Error(`the site answered ${response.status}`);
  return response.text();
}

/** The install file for reps: the add-in's manifest with this team's key in it. */
async function onShareInstall() {
  let xml;
  try {
    xml = (await manifestTemplate()).split('{{BASE}}').join(SITE_ROOT).split('{{QUERY}}').join(`?team=${team.key}`);
  } catch (error) {
    teamNote(`Couldn't make the install file: ${error.message}`, 'error');
    return;
  }
  downloadText(xml, 'restack-team.xml', 'application/xml');
  log('Team: install file downloaded');
  teamNote('restack-team.xml downloaded. Email it to a rep; they install it at aka.ms/olksideload (My add-ins → Add a custom add-in → Add from File). Anyone with this file can use the add-in and see the team templates, so keep it within your team.', 'ok');
}

/**
 * The beta's install file: the same add-in under its own ID and name (Restack Beta), loading from the
 * site's beta folder, so it installs next to Restack. For trying a new version before reps get it.
 */
function betaManifest(template) {
  return template
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${BETA_ID}</Id>`)
    .replace('<DisplayName DefaultValue="Restack"/>', '<DisplayName DefaultValue="Restack Beta"/>')
    .replace(/(<bt:String id="(?:Group|ReadButton)\.Label" DefaultValue=")Restack(")/g, '$1Restack Beta$2')
    .replace(/icon-(16|32|64|80|128)\.png/g, 'icon-beta-$1.png')
    .split('{{BASE}}').join(`${SITE_ROOT}beta/`)
    .split('{{QUERY}}').join(`?team=${team.key}`);
}

async function onShareBeta() {
  let xml;
  try {
    xml = betaManifest(await manifestTemplate());
  } catch (error) {
    teamNote(`Couldn't make the beta install file: ${error.message}`, 'error');
    return;
  }
  downloadText(xml, 'restack-beta.xml', 'application/xml');
  log('Team: beta install file downloaded');
  teamNote("restack-beta.xml downloaded. Install it like Restack's file (aka.ms/olksideload → My add-ins → Add a custom add-in → Add from File) and it shows up as Restack Beta, next to Restack. It runs whatever is in your repository's beta folder, so put a new version there first, try it, then upload the same files to the main folder. Only you need it: keep it to yourself.", 'ok');
}

// --------------------------------------------------------------------------- usage counts
// With the team's Restack Usage function set up, the panel counts what each rep makes: PDFs saved,
// emails created (with their template), rate cons left out. Only counts go: no files, file names, load
// numbers or customer addresses. They wait in this browser until they're sent, so none are lost offline.

const USAGE_QUEUE = 'usageQueue';

function usageUnreachable() {
  return userError(`couldn't reach the Restack Usage function (check its address, and that its CORS settings list ${window.location.origin})`);
}

async function usageProblem(response) {
  let detail = '';
  try {
    const body = await response.json();
    detail = body && body.error ? String(body.error).replace(/\.$/, '') : '';
  } catch (error) {
    detail = '';
  }
  if (response.status === 401) return userError("the key doesn't match the Function App's setting");
  if (response.status === 503) return userError(detail || "the Function App isn't set up yet (its keys or storage)");
  if (response.status === 404) return userError(`there's no ${USAGE_FUNCTION} function at that address`);
  return userError(`Azure answered ${response.status}${detail ? `: ${detail}` : ''}`);
}

/** The function's own report: { service, version, keySet, adminKeySet }. */
async function usageHealth(url) {
  let response;
  try {
    response = await fetch(url, { cache: 'no-store' });
  } catch (error) {
    throw usageUnreachable();
  }
  if (!response.ok) throw await usageProblem(response);
  let info = {};
  try {
    info = await response.json();
  } catch (error) {
    info = {};
  }
  if (info.service !== 'restack-usage') throw userError("that address answers, but it isn't the Restack Usage function");
  return info;
}

/** Sends counts ({ events }) or a key check ({ check: true }) with the counting key. */
async function sendUsage(usage, body) {
  let response;
  try {
    response = await fetch(usage.url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-usage-key': usage.key },
      body: JSON.stringify(body),
      keepalive: true,
    });
  } catch (error) {
    throw usageUnreachable();
  }
  if (!response.ok) throw await usageProblem(response);
  try {
    return await response.json();
  } catch (error) {
    return {};
  }
}

/** Everyone's counts from `from` to `to` (YYYY-MM-DD), with the admin key. */
async function readUsage(url, adminKey, from, to) {
  let response;
  try {
    response = await fetch(`${url}?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`, { headers: { 'x-usage-admin-key': adminKey }, cache: 'no-store' });
  } catch (error) {
    throw usageUnreachable();
  }
  if (!response.ok) throw await usageProblem(response);
  const data = await response.json();
  const events = Array.isArray(data && data.events) ? data.events : [];
  events.truncated = !!(data && data.truncated);
  return events;
}

function usageQueue() {
  try {
    const queue = JSON.parse(readSetting(USAGE_QUEUE) || '[]');
    return Array.isArray(queue) ? queue : [];
  } catch (error) {
    return [];
  }
}

/** Counts one thing a rep did, and sends what's waiting. */
function countUsage(action, details = {}) {
  if (CONSOLE || !teamUsage()) return;
  const event = { id: randomB64url(9), at: new Date().toISOString(), rep: team.me || '', action, version: VERSION, beta: BETA, ...details };
  writeSetting(USAGE_QUEUE, JSON.stringify([...usageQueue(), event].slice(-300)));
  flushUsage().catch(() => {});
}

let usageSending = false;

/** Sends what's waiting, 50 at a time; counts made while a batch is on its way go in the next one. */
async function flushUsage() {
  const usage = teamUsage();
  if (!usage || usageSending) return;
  usageSending = true;
  try {
    for (let round = 0; round < 10; round++) {
      const batch = usageQueue().slice(0, 50);
      if (!batch.length) break;
      await sendUsage(usage, { events: batch });
      const sent = new Set(batch.map((event) => event.id));
      writeSetting(USAGE_QUEUE, JSON.stringify(usageQueue().filter((event) => !sent.has(event.id))));
      log(`Usage counts: sent ${plural(batch.length, 'count')}`);
    }
  } catch (error) {
    log(`Usage counts: not sent yet (${error.userMessage || error.message})`);
  } finally {
    usageSending = false;
  }
}

/** The figures a count carries about the files: how many PDFs, pages and checked files, and how they were made. */
function usageFigures(files) {
  return {
    pdfs: files.length,
    pages: files.reduce((sum, file) => sum + (file.pages || 0), 0),
    files: checkedFiles().length,
    split: files.filter((file) => !file.stop && !file.lumper).length > 1, // over 10 MB, in parts
    multiStop: !!state.multiStop,
  };
}

// --------------------------------------------------------------------------- the admin console
// taskpane.html?console (console.html sends you here) in a browser window: the team's templates and
// settings at full size, and the Usage tab. Settings → Team → Open admin console passes the team key
// (and the usage admin key) after # in the address, so they never go to a server. They're kept in this
// browser (thread-attachments.console.*), so a bookmark to the console works next time. The console
// works from what's published in team.json: unpublished changes made in Outlook don't show here.

const CONSOLE_TABS = ['templates', 'team', 'usage'];
let consoleTab = 'templates';
let consoleSearch = '';

/** Where to go to publish, in messages: the panel's Settings, or the console's header. */
function publishWhere() {
  return CONSOLE ? 'Publish to team (top right)' : 'Publish (Settings → Team)';
}

function consoleNote(text, tone) {
  const note = $('con-note');
  if (!note) return;
  note.textContent = text || '';
  note.hidden = !text;
  note.className = `con-note${tone ? ` ${tone}` : ''}`;
}

/** Keeps the keys from the link (#k=…&u=…) in this browser, and takes them out of the address bar. */
function readConsoleHash() {
  const hash = new URLSearchParams(window.location.hash.replace(/^#/, ''));
  const key = hash.get('k');
  const admin = hash.get('u');
  if (key && /^[A-Za-z0-9_-]{22}$/.test(key)) writeSetting('teamKey', key);
  if (admin && admin.length >= 16) writeSetting('usageAdmin', JSON.stringify(admin)); // what readRoaming reads outside Outlook
  if (window.location.hash) {
    try {
      history.replaceState(null, '', `${window.location.pathname}${window.location.search}`);
    } catch (error) {
      // the key stays in the address bar; harmless
    }
  }
}

async function startConsole() {
  document.documentElement.classList.add('is-console');
  document.title = BETA ? 'Restack Beta admin console' : 'Restack admin console';
  $('app').hidden = true;
  $('console').hidden = false;
  $('con-beta').hidden = !BETA;
  readConsoleHash();
  loadAppearance(); // Match follows the computer here
  // The panel's own screens move into the console's tabs
  $('con-templates').append($('templates'));
  $('templates').hidden = false;
  for (const card of document.querySelectorAll('.con-card[data-section]')) {
    const section = $(card.dataset.section);
    card.querySelector('.con-card-body').append(section);
    section.hidden = false;
  }
  $('con-foot').append($('diagnostics')); // the test report, at the bottom
  wireReadUi(); // the template and team screens' buttons
  wireConsole();
  consoleNote('Opening your team’s templates…');
  await loadTeam();
  if (team.status === 'ok') team.me = team.published.admins[0] || '';
  consoleNote('');
  if (team.status !== 'ok') {
    showConsoleLocked();
    return;
  }
  log(`Admin console: ${plural(teamTemplates().length, 'team template')}${team.draft ? ', with changes not published yet' : ''}`);
  $('con-tabs').hidden = false;
  renderConsole();
  showConsoleTab(CONSOLE_TABS.includes(readSetting('tab')) ? readSetting('tab') : 'templates', false);
}

/** No key, the wrong key, nothing published yet, or the site couldn't be reached. */
function showConsoleLocked() {
  const box = $('con-locked');
  box.textContent = '';
  const card = el('div', 'con-locked-card');
  const [title, ...lines] = {
    nokey: ['Open the console from Outlook',
      'In Restack, choose Settings (the gear at the bottom) → Team → Open admin console. That link carries your team key, and this browser keeps it so the console opens straight away next time.'],
    wrongkey: ["This key doesn't open your team's templates",
      'It may be from another team or an old install file. Remove it, then open the console again from Outlook: Settings → Team → Open admin console.'],
    nofile: ['Nothing is published yet',
      'Set up your team in Outlook first: open Restack, choose Settings → Manage templates, add a template and publish it. Then open the console again.'],
    error: ["Couldn't load your team's templates",
      `Check your internet connection and try again. (${team.error})`],
  }[team.status] || ['Something went wrong', team.error || ''];
  card.append(el('h1', '', title), ...lines.map((line) => el('p', '', line)));
  const row = el('div', 'btn-row');
  if (team.status === 'wrongkey') {
    const forget = el('button', 'btn', 'Remove the key from this browser');
    forget.type = 'button';
    forget.addEventListener('click', () => {
      forgetConsoleKeys();
      window.location.reload();
    });
    row.append(forget);
  }
  if (team.status === 'error') {
    const again = el('button', 'btn primary', 'Try again');
    again.type = 'button';
    again.addEventListener('click', () => window.location.reload());
    row.append(again);
  }
  if (row.childElementCount) card.append(row);
  box.append(card);
  box.hidden = false;
}

/** Everything the console keeps in this browser: the keys, the GitHub token, unpublished changes and its choices. */
function forgetConsoleKeys() {
  try {
    for (const key of Object.keys(localStorage)) if (key.startsWith(STORE_PREFIX)) localStorage.removeItem(key);
    sessionStorage.removeItem(`${STORE_PREFIX}githubToken`);
  } catch (error) {
    // nothing kept, or storage is off
  }
}

let forgetArmed = 0;

function showConsoleTab(tab, focus = true) {
  if (!CONSOLE_TABS.includes(tab)) return;
  consoleTab = tab;
  writeSetting('tab', tab);
  for (const name of CONSOLE_TABS) {
    const button = $(`con-tab-${name}`);
    const on = name === tab;
    button.setAttribute('aria-selected', String(on));
    button.tabIndex = on ? 0 : -1;
    $(`con-${name}`).hidden = !on;
  }
  if (focus) $(`con-tab-${tab}`).focus();
  if (tab === 'usage') loadUsageReport().catch(fail);
}

function renderConsole() {
  renderConsolePublish();
  renderTemplateSheet();
  renderAdminsBox();
  renderRcBox();
  renderAzureBox();
  renderReceiptWords();
  renderUsageBox();
  renderGithubBox();
  const settings = teamSettings();
  markSwitch('con-rep-templates', !!(settings && settings.repTemplates));
}

/** The header: whether there are changes to publish, and the button. */
function renderConsolePublish() {
  if (!CONSOLE || team.status !== 'ok') return;
  const pending = !!team.draft && !team.justPublished;
  const box = $('con-publish-box');
  box.hidden = false;
  box.classList.toggle('pending', pending);
  const when = team.published && team.published.updated
    ? new Date(team.published.updated).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })
    : '';
  $('con-publish-text').textContent = pending
    ? "Changes your team can't see yet"
    : team.justPublished ? (githubReady() ? 'Published to GitHub' : 'team.json downloaded') : `Published${when ? ` ${when}` : ''}`;
  $('con-publish').disabled = !pending;
  const settings = teamSettings();
  markSwitch('con-rep-templates', !!(settings && settings.repTemplates));
  markSwitch('con-sent-with', !!(settings && settings.sentWith));
}

function wireConsole() {
  const tabs = $('con-tabs');
  for (const name of CONSOLE_TABS) $(`con-tab-${name}`).addEventListener('click', () => showConsoleTab(name));
  tabs.addEventListener('keydown', (event) => {
    const at = CONSOLE_TABS.indexOf(consoleTab);
    const next = { ArrowRight: at + 1, ArrowLeft: at - 1, Home: 0, End: CONSOLE_TABS.length - 1 }[event.key];
    if (next === undefined) return;
    event.preventDefault();
    showConsoleTab(CONSOLE_TABS[(next + CONSOLE_TABS.length) % CONSOLE_TABS.length]);
  });
  $('con-publish').addEventListener('click', () => onPublish().catch(fail));
  $('con-new').addEventListener('click', () => editTemplate('team', null));
  $('tpl-search').addEventListener('input', (event) => {
    consoleSearch = event.target.value;
    renderTemplateSheet();
  });
  $('con-share').addEventListener('click', () => onShareInstall().catch(fail));
  $('con-share-beta').addEventListener('click', () => onShareBeta().catch(fail));
  $('con-rep-templates').addEventListener('click', (event) => onRepTemplatesToggle(event.currentTarget.getAttribute('aria-checked') !== 'true').catch(fail));
  $('con-sent-with')?.addEventListener('click', (event) => onSentWithToggle(event.currentTarget.getAttribute('aria-checked') !== 'true').catch(fail));
  $('rc-choose').addEventListener('click', () => $('rc-file').click());
  $('rc-file').addEventListener('change', (event) => onRcFile(event.target).catch(fail));
  $('gh-save').addEventListener('click', onGithubSave);
  $('gh-test').addEventListener('click', () => onGithubTest().catch(fail));
  $('gh-forget').addEventListener('click', onGithubForget);
  $('con-forget').addEventListener('click', () => {
    if (team.draft && !team.justPublished && Date.now() - forgetArmed > 8000) {
      forgetArmed = Date.now();
      setStatusIn('forget-status', "Your changes that aren't published yet are kept here too, and they'd go as well. Publish first, or click again to remove everything.", 'error');
      return;
    }
    forgetArmed = 0;
    forgetConsoleKeys();
    team.draft = null;
    renderGithubBox();
    log('Admin console: removed what it kept in this browser');
    setStatusIn('forget-status', 'Removed. To open the console again, use Settings → Team → Open admin console in Outlook.', 'ok');
  });
  for (const button of document.querySelectorAll('#con-usage .us-range')) button.addEventListener('click', () => pickUsageRange(button.dataset.range));
  for (const id of ['us-from', 'us-to']) $(id).addEventListener('change', () => pickUsageDates());
  $('us-refresh').addEventListener('click', () => loadUsageReport(true).catch(fail));
  $('us-csv').addEventListener('click', downloadUsageCsv);
  $('us-minutes').value = String(usageMinutes());
  $('us-minutes').addEventListener('change', (event) => {
    const minutes = Math.min(60, Math.max(0, Number(event.target.value) || 0));
    writeSetting('minutesSaved', String(minutes));
    if (usageReport.events) renderUsageReport();
  });
  window.addEventListener('beforeunload', (event) => {
    if (editing && templateFormChanged()) {
      event.preventDefault();
      event.returnValue = '';
    }
  });
}

/** A rate confirmation PDF from the computer becomes the sample (the panel uses one previewed from an email). */
async function onRcFile(input) {
  const chosen = input.files && input.files[0];
  input.value = '';
  if (!chosen) return;
  const button = $('rc-choose');
  button.disabled = true;
  renderRcBox(`Reading ${chosen.name}…`);
  try {
    const file = { name: chosen.name, kind: 'PDF', bytes: new Uint8Array(await chosen.arrayBuffer()) };
    const sample = await rateConSample(file);
    if (!sample.pages.some((page) => page.words || page.grid)) throw userError('nothing could be read from this file');
    await saveTeamDraft({ rateCon: sample });
    log(`Rate con check: sample set from ${chosen.name} (${plural(sample.pages.length, 'page')}), not published yet`);
    renderConsolePublish();
    renderRcBox();
    consoleNote(`${chosen.name} is the sample now. ${publishWhere()} to turn the check on for your reps.`, 'ok');
  } catch (error) {
    renderRcBox(`Couldn't use ${chosen.name}: ${error.userMessage || error.message}.`);
  } finally {
    button.disabled = false;
  }
}

// ---- the template editor's unsaved changes (so closing the tab asks first)

let formSnapshot = '';

function templateFormState() {
  return JSON.stringify([...document.querySelectorAll('#tpl-form input, #tpl-form textarea')].map((node) => (node.type === 'checkbox' ? node.checked : node.value)).concat(JSON.stringify(qDraft)));
}

function templateFormChanged() {
  return !!formSnapshot && templateFormState() !== formSnapshot;
}

// ---- publishing straight to GitHub (optional): a fine-grained token for the one repository

const GITHUB_API = 'https://api.github.com';

/** owner/repository from the site's address (https://owner.github.io/repository/), as a starting point. */
function defaultGithubRepo() {
  const match = /^https:\/\/([a-z0-9-]+)\.github\.io\/([^/]+)\//i.exec(SITE_ROOT);
  return match ? `${match[1]}/${match[2]}` : '';
}

function githubToken() {
  try {
    const now = sessionStorage.getItem(`${STORE_PREFIX}githubToken`);
    if (now) return now;
  } catch (error) {
    // not available; try the remembered one
  }
  return readSetting('githubToken') || '';
}

function githubRepo() {
  return readSetting('githubRepo') || defaultGithubRepo();
}

function githubReady() {
  return CONSOLE && !!githubToken() && /^[\w.-]+\/[\w.-]+$/.test(githubRepo());
}

function renderGithubBox(message, tone) {
  if (!CONSOLE) return;
  $('gh-repo').value = githubRepo();
  $('gh-token').value = githubToken();
  $('gh-remember').checked = !!readSetting('githubToken');
  $('gh-forget').hidden = !githubToken();
  $('github-state').textContent = githubReady() ? 'Straight to GitHub' : 'Download and upload';
  setStatusIn('gh-status', message || (githubReady()
    ? `Publish to team commits team.json to ${githubRepo()} for you.`
    : 'Without a token, Publish to team downloads team.json for you to upload.'), tone);
}

function onGithubSave() {
  const repo = $('gh-repo').value.trim().replace(/^https:\/\/github\.com\//i, '').replace(/\.git$/i, '').replace(/\/+$/, '');
  const token = $('gh-token').value.trim();
  if (repo && !/^[\w.-]+\/[\w.-]+$/.test(repo)) {
    setStatusIn('gh-status', 'The repository should look like owner/repository, for example huntrbwdn1/restack_hb.', 'error');
    return;
  }
  if (token && !/^(github_pat_|ghp_)[A-Za-z0-9_]{20,}$/.test(token)) {
    setStatusIn('gh-status', "That doesn't look like a GitHub token. Fine-grained tokens start with github_pat_.", 'error');
    return;
  }
  writeSetting('githubRepo', repo && repo !== defaultGithubRepo() ? repo : null);
  const remember = $('gh-remember').checked;
  writeSetting('githubToken', token && remember ? token : null);
  try {
    if (token && !remember) sessionStorage.setItem(`${STORE_PREFIX}githubToken`, token);
    else sessionStorage.removeItem(`${STORE_PREFIX}githubToken`);
  } catch (error) {
    // sessionStorage unavailable: only a remembered token works
  }
  log(`Admin console: GitHub publishing ${token ? 'set up' : 'off'}${token ? (remember ? ' (remembered)' : ' (this tab only)') : ''}`);
  renderGithubBox(token ? 'Saved. Click Test to check the token can reach the repository.' : 'Saved. Publish to team downloads team.json.', 'ok');
  renderConsolePublish();
}

function onGithubForget() {
  writeSetting('githubToken', null);
  try {
    sessionStorage.removeItem(`${STORE_PREFIX}githubToken`);
  } catch (error) {
    // nothing kept there
  }
  log('Admin console: GitHub token removed');
  renderGithubBox('Removed. Publish to team downloads team.json again.', 'ok');
}

async function github(path, options = {}) {
  const token = githubToken();
  let response;
  try {
    response = await fetch(`${GITHUB_API}${path}`, {
      method: options.method || 'GET',
      headers: {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${token}`,
        'X-GitHub-Api-Version': '2022-11-28',
        ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      },
      body: options.body,
      cache: 'no-store',
    });
  } catch (error) {
    throw userError("couldn't reach GitHub (check your internet connection)");
  }
  if (response.status === 404 && options.missingOk) return null;
  if (response.status === 401) throw userError("GitHub didn't accept the token (it may have expired)");
  if (response.status === 403) throw userError("the token isn't allowed to change that repository (it needs Contents: Read and write)");
  if (response.status === 404) throw userError(`GitHub can't find ${githubRepo()} with this token (check the name, and that the token includes that repository)`);
  if (response.status === 409 || response.status === 422) throw userError('team.json changed on GitHub at the same moment; publish again');
  if (!response.ok) throw userError(`GitHub answered ${response.status}`);
  return response.json();
}

async function onGithubTest() {
  if (!githubToken()) {
    renderGithubBox('Paste a token and click Save first.', 'error');
    return;
  }
  $('gh-test').disabled = true;
  setStatusIn('gh-status', 'Testing…');
  try {
    const repo = githubRepo();
    const info = await github(`/repos/${repo}`);
    const branch = info.default_branch || 'main';
    const file = await github(`/repos/${repo}/contents/${TEAM_FILE}?ref=${encodeURIComponent(branch)}`, { missingOk: true });
    log(`Admin console: GitHub test worked (${repo}, ${branch}${file ? ', team.json found' : ', no team.json yet'})`);
    setStatusIn('gh-status', `Works: the token reaches ${repo}${file ? `, and team.json is on its ${branch} branch` : ` (${branch} branch; no team.json there yet)`}. Publish to team will commit team.json there.`, 'ok');
  } catch (error) {
    setStatusIn('gh-status', `Didn't work: ${error.userMessage || error.message}.`, 'error');
  } finally {
    $('gh-test').disabled = false;
  }
}

/** Commits team.json to the repository's default branch, replacing the one there. */
async function publishToGithub(text) {
  const repo = githubRepo();
  const info = await github(`/repos/${repo}`);
  const branch = info.default_branch || 'main';
  const now = await github(`/repos/${repo}/contents/${TEAM_FILE}?ref=${encodeURIComponent(branch)}`, { missingOk: true });
  await github(`/repos/${repo}/contents/${TEAM_FILE}`, {
    method: 'PUT',
    body: JSON.stringify({
      message: 'Publish team templates (Restack admin console)',
      content: bytesToBase64(new TextEncoder().encode(text)),
      branch,
      ...(now && now.sha ? { sha: now.sha } : {}),
    }),
  });
  return { repo, branch };
}

// ---- the Usage tab: everyone's counts, read from the Restack Usage function with the admin key

const usageReport = { from: '', to: '', events: null, loadedAt: 0, key: '' };

function usageMinutes() {
  const saved = Number(readSetting('minutesSaved'));
  return Number.isFinite(saved) && saved >= 0 && readSetting('minutesSaved') !== null ? saved : 3;
}

function isoDay(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function dayFromIso(text) {
  const [y, m, d] = String(text).split('-').map(Number);
  return new Date(y, m - 1, d);
}

/** [from, to] (YYYY-MM-DD, this computer's days) for a range button. */
function usageRange(range) {
  const today = new Date();
  const start = new Date(today);
  if (range === '7') start.setDate(today.getDate() - 6);
  else if (range === '30') start.setDate(today.getDate() - 29);
  else if (range === 'month') start.setDate(1);
  else if (range === 'year') start.setMonth(0, 1);
  else if (range === 'last-month') {
    const first = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const last = new Date(today.getFullYear(), today.getMonth(), 0);
    return [isoDay(first), isoDay(last)];
  }
  return [isoDay(start), isoDay(today)];
}

function pickUsageRange(range) {
  const [from, to] = usageRange(range);
  $('us-from').value = from;
  $('us-to').value = to;
  writeSetting('usageRange', range);
  for (const button of document.querySelectorAll('#con-usage .us-range')) button.setAttribute('aria-pressed', String(button.dataset.range === range));
  loadUsageReport().catch(fail);
}

function pickUsageDates() {
  const from = $('us-from').value;
  const to = $('us-to').value;
  if (!from || !to) return;
  for (const button of document.querySelectorAll('#con-usage .us-range')) button.setAttribute('aria-pressed', 'false');
  writeSetting('usageRange', null);
  loadUsageReport().catch(fail);
}

function usageEmpty(text, button) {
  $('us-report').hidden = true;
  $('us-csv').disabled = true;
  const box = $('us-empty');
  box.textContent = '';
  box.append(el('p', '', text));
  if (button) {
    const go = el('button', 'btn', button);
    go.type = 'button';
    go.addEventListener('click', () => {
      showConsoleTab('team');
      $('con-card-usage').scrollIntoView({ block: 'start' });
      $('us-url').focus();
    });
    box.append(go);
  }
  box.hidden = false;
}

async function loadUsageReport(force = false) {
  if (!$('us-from').value || !$('us-to').value) {
    const range = readSetting('usageRange') || '30';
    const [from, to] = usageRange(range);
    $('us-from').value = from;
    $('us-to').value = to;
    for (const button of document.querySelectorAll('#con-usage .us-range')) button.setAttribute('aria-pressed', String(button.dataset.range === range));
  }
  const usage = teamUsage();
  const adminKey = usageAdminKey();
  if (!usage) {
    $('us-line').textContent = '';
    usageEmpty("Usage counts aren't on yet. Set up your Restack Usage function under Team settings → Usage counts, then publish, and each rep's panel starts counting.", 'Set up usage counts');
    return;
  }
  if (!adminKey) {
    $('us-line').textContent = '';
    usageEmpty('Add your admin key under Team settings → Usage counts to see the numbers here. It stays in this browser.', 'Add the admin key');
    return;
  }
  let from = $('us-from').value;
  let to = $('us-to').value;
  if (from > to) [from, to] = [to, from];
  const key = `${usage.url}|${from}|${to}`;
  if (!force && usageReport.key === key && usageReport.events && Date.now() - usageReport.loadedAt < 60000) {
    renderUsageReport();
    return;
  }
  $('us-empty').hidden = true;
  $('us-line').textContent = 'Loading…';
  $('us-refresh').disabled = true;
  try {
    // The function keeps UTC dates; ask for a day either side, then keep this computer's days
    const before = dayFromIso(from);
    before.setDate(before.getDate() - 1);
    const after = dayFromIso(to);
    after.setDate(after.getDate() + 1);
    const events = await readUsage(usage.url, adminKey, before.toISOString().slice(0, 10), after.toISOString().slice(0, 10));
    Object.assign(usageReport, {
      from, to, key, loadedAt: Date.now(), truncated: !!events.truncated,
      events: events.filter((event) => {
        const at = new Date(event.at);
        if (Number.isNaN(at.getTime())) return false;
        const day = isoDay(at);
        return day >= from && day <= to;
      }),
    });
    log(`Usage: ${plural(usageReport.events.length, 'count')} from ${from} to ${to}`);
    renderUsageReport();
  } catch (error) {
    $('us-line').textContent = '';
    usageEmpty(`Couldn't load the numbers: ${error.userMessage || error.message}.`);
  } finally {
    $('us-refresh').disabled = false;
  }
}

/** Adds up the counts: totals, per rep, per template, per day and per version. */
function usageSummary(events) {
  const totals = { emails: 0, emailEvents: 0, saves: 0, pdfs: 0, files: 0, pages: 0, accessorials: 0, separate: 0, rateCons: 0, rateConPages: 0 };
  const reps = new Map();
  const templates = new Map();
  const days = new Map();
  const versions = new Map();
  const row = (map, key, start) => {
    if (!map.has(key)) map.set(key, { ...start });
    return map.get(key);
  };
  for (const event of events) {
    const at = new Date(event.at);
    const day = isoDay(at);
    const rep = row(reps, event.rep || 'Unknown', { name: event.rep || 'Unknown', emails: 0, saves: 0, files: 0, pages: 0, rateCons: 0, last: 0 });
    rep.last = Math.max(rep.last, at.getTime());
    const perDay = row(days, day, { emails: 0, saves: 0 });
    if (event.version) {
      const version = row(versions, `${event.version}${event.beta ? ' beta' : ''}`, { name: `${event.version}${event.beta ? ' (beta)' : ''}`, reps: new Set(), last: 0 });
      version.reps.add(event.rep || 'Unknown');
      version.last = Math.max(version.last, at.getTime());
    }
    if (event.action === 'email') {
      const emails = Number(event.emails) || 1;
      totals.emails += emails;
      totals.emailEvents += 1;
      totals.files += Number(event.files) || 0;
      totals.pages += Number(event.pages) || 0;
      totals.pdfs += Number(event.pdfs) || 0;
      if (event.accessorials) totals.accessorials += 1;
      if (event.separate) totals.separate += 1;
      rep.emails += emails;
      rep.files += Number(event.files) || 0;
      rep.pages += Number(event.pages) || 0;
      perDay.emails += emails;
      const name = event.template || 'Blank email';
      const t = row(templates, event.templateId || name, { name, emails: 0, accessorials: 0, separate: 0, files: 0, pages: 0 });
      t.name = name; // the latest name, if it was renamed
      t.emails += emails;
      t.files += Number(event.files) || 0;
      t.pages += Number(event.pages) || 0;
      if (event.accessorials) t.accessorials += 1;
      if (event.separate) t.separate += 1;
    } else if (event.action === 'save') {
      totals.saves += 1;
      totals.files += Number(event.files) || 0;
      totals.pages += Number(event.pages) || 0;
      totals.pdfs += Number(event.pdfs) || 0;
      rep.saves += 1;
      rep.files += Number(event.files) || 0;
      rep.pages += Number(event.pages) || 0;
      perDay.saves += 1;
    } else if (event.action === 'rateCon') {
      totals.rateCons += 1;
      totals.rateConPages += Number(event.pages) || 0;
      rep.rateCons += 1;
    }
  }
  return { totals, reps: [...reps.values()], templates: [...templates.values()], days, versions: [...versions.values()] };
}

function fmt(n) {
  return Number(n || 0).toLocaleString();
}

function hoursText(minutes) {
  if (minutes < 60) return `${Math.round(minutes)} min`;
  const hours = minutes / 60;
  return `${hours < 10 ? hours.toFixed(1).replace(/\.0$/, '') : fmt(Math.round(hours))} hr`;
}

function renderUsageReport() {
  const events = usageReport.events || [];
  const from = dayFromIso(usageReport.from);
  const to = dayFromIso(usageReport.to);
  const span = `${from.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: from.getFullYear() === to.getFullYear() ? undefined : 'numeric' })} – ${to.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}`;
  const loaded = new Date(usageReport.loadedAt).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  $('us-line').textContent = `${span} · ${plural(events.length, 'count')}${usageReport.truncated ? ' (the first 100,000 only: pick fewer days to see them all)' : ''} · updated ${loaded}`;
  $('us-csv').disabled = !events.length;
  if (!events.length) {
    usageEmpty('Nothing was counted in these dates. Counts show up once your reps have the published settings and save a PDF or create an email.');
    return;
  }
  $('us-empty').hidden = true;
  $('us-report').hidden = false;
  const { totals, reps, templates, days, versions } = usageSummary(events);
  const minutes = (totals.emailEvents + totals.saves) * usageMinutes();
  const figures = [
    ['Emails created', fmt(totals.emails), `${fmt(totals.accessorials)} with accessorials`],
    ['PDFs saved', fmt(totals.saves), 'with Save PDF'],
    ['Files combined', fmt(totals.files), `into ${fmt(totals.pdfs)} PDFs`],
    ['Pages', fmt(totals.pages), 'in those PDFs'],
    ['Rate cons caught', fmt(totals.rateCons), `${plural(totals.rateConPages, 'page')} left out`],
    ['Reps', fmt(reps.length), 'using Restack'],
    ['Time saved', hoursText(minutes), `at ${usageMinutes()} min each`],
  ];
  const list = $('us-figures');
  list.textContent = '';
  for (const [label, value, note] of figures) {
    const item = el('div', 'us-fig');
    item.append(el('dt', '', label), el('dd', 'us-val', value), el('dd', 'us-sub', note));
    list.append(item);
  }
  renderUsageChart(days);
  const dateText = (ms) => (ms ? new Date(ms).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : '');
  fillUsageTable($('us-reps'), ['Rep', 'Emails', 'PDFs saved', 'Files combined', 'Pages', 'Rate cons caught', 'Last used'],
    reps.sort((a, b) => (b.emails + b.saves) - (a.emails + a.saves) || a.name.localeCompare(b.name))
      .map((r) => [r.name, fmt(r.emails), fmt(r.saves), fmt(r.files), fmt(r.pages), fmt(r.rateCons), dateText(r.last)]));
  fillUsageTable($('us-templates'), ['Template', 'Emails', 'With accessorials', 'Separate accessorial email', 'Files combined', 'Pages'],
    templates.sort((a, b) => b.emails - a.emails || a.name.localeCompare(b.name))
      .map((t) => [t.name, fmt(t.emails), fmt(t.accessorials), fmt(t.separate), fmt(t.files), fmt(t.pages)]));
  fillUsageTable($('us-versions'), ['Version', 'Reps', 'Last used'],
    versions.sort((a, b) => b.last - a.last).map((v) => [v.name, fmt(v.reps.size), dateText(v.last)]));
}

function fillUsageTable(table, heads, rows) {
  table.textContent = '';
  const head = document.createElement('thead');
  const tr = document.createElement('tr');
  heads.forEach((text, i) => {
    const th = el('th', i ? 'num' : '', text);
    th.scope = 'col';
    tr.append(th);
  });
  head.append(tr);
  const body = document.createElement('tbody');
  for (const values of rows) {
    const row = document.createElement('tr');
    values.forEach((value, i) => {
      const cell = el(i ? 'td' : 'th', i ? 'num' : '', value);
      if (!i) cell.scope = 'row';
      row.append(cell);
    });
    body.append(row);
  }
  table.append(head, body);
}

/** Emails and saved PDFs per day, as bars (an SVG that stretches to the width). */
function renderUsageChart(days) {
  const box = $('us-chart');
  box.textContent = '';
  const from = dayFromIso(usageReport.from);
  const to = dayFromIso(usageReport.to);
  const list = [];
  for (const at = new Date(from); at <= to && list.length < 400; at.setDate(at.getDate() + 1)) {
    const key = isoDay(at);
    list.push({ key, date: new Date(at), ...(days.get(key) || { emails: 0, saves: 0 }) });
  }
  const top = Math.max(1, ...list.map((d) => d.emails + d.saves));
  const step = top <= 5 ? 1 : top <= 10 ? 2 : top <= 25 ? 5 : top <= 50 ? 10 : Math.ceil(top / 5 / 10) * 10;
  const max = Math.ceil(top / step) * step;
  const W = 960;
  const H = 220;
  const left = 34;
  const bottom = 26;
  const plotW = W - left - 6;
  const plotH = H - bottom - 10;
  const slot = plotW / list.length;
  const barW = Math.max(2, Math.min(28, slot * 0.68));
  const y = (v) => 10 + plotH - (v / max) * plotH;
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Emails and saved PDFs per day" preserveAspectRatio="none">`;
  for (let v = 0; v <= max; v += step) {
    svg += `<line class="grid" x1="${left}" x2="${W - 6}" y1="${y(v).toFixed(1)}" y2="${y(v).toFixed(1)}"/><text class="tick" x="${left - 6}" y="${(y(v) + 4).toFixed(1)}" text-anchor="end">${v}</text>`;
  }
  const every = Math.ceil(list.length / 12);
  list.forEach((d, i) => {
    const x = left + slot * i + (slot - barW) / 2;
    const title = `${d.date.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' })}: ${plural(d.emails, 'email')}, ${plural(d.saves, 'saved PDF')}`;
    svg += `<g><title>${title}</title>`;
    if (d.saves) svg += `<rect class="bar-save" x="${x.toFixed(1)}" y="${y(d.saves).toFixed(1)}" width="${barW.toFixed(1)}" height="${(y(0) - y(d.saves)).toFixed(1)}" rx="1.5"/>`;
    if (d.emails) svg += `<rect class="bar-email" x="${x.toFixed(1)}" y="${y(d.saves + d.emails).toFixed(1)}" width="${barW.toFixed(1)}" height="${(y(d.saves) - y(d.saves + d.emails)).toFixed(1)}" rx="1.5"/>`;
    svg += '</g>';
    if (i % every === 0) svg += `<text class="tick" x="${(left + slot * i + slot / 2).toFixed(1)}" y="${H - 8}" text-anchor="middle">${d.date.toLocaleDateString(undefined, { month: 'numeric', day: 'numeric' })}</text>`;
  });
  svg += '</svg>';
  box.innerHTML = svg;
}

function csvCell(value) {
  let text = String(value == null ? '' : value);
  if (/^[=+\-@\t\r]/.test(text) && typeof value === 'string') text = `'${text}`; // no formulas when it's opened in Excel
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function downloadUsageCsv() {
  const events = usageReport.events || [];
  if (!events.length) return;
  const heads = ['date', 'time', 'rep', 'action', 'template', 'emails', 'accessorials', 'separate', 'pdfs', 'files', 'pages', 'split', 'multiStop', 'azure', 'version', 'beta'];
  const lines = [heads.join(',')];
  for (const event of [...events].sort((a, b) => String(a.at).localeCompare(String(b.at)))) {
    const at = new Date(event.at);
    lines.push([isoDay(at), at.toTimeString().slice(0, 5), event.rep, event.action, event.template, event.emails, event.accessorials, event.separate, event.pdfs, event.files, event.pages, event.split, event.multiStop, event.azure, event.version, event.beta].map(csvCell).join(','));
  }
  downloadText(`${lines.join('\r\n')}\r\n`, `restack-usage-${usageReport.from}-to-${usageReport.to}.csv`, 'text/csv');
  log(`Usage: CSV downloaded (${plural(events.length, 'count')})`);
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

/** The emails Create email opened and their files, from the hand-over store (one, or two). */
function pendingEmails(pending) {
  if (!pending) return [];
  const raw = Array.isArray(pending.emails) && pending.emails.length
    ? pending.emails
    : [{ subject: pending.subject, files: Array.isArray(pending.files) ? pending.files : [{ name: pending.name, bytes: pending.bytes, pages: pending.pages }] }];
  return raw
    .map((e) => ({
      subject: String((e && e.subject) || ''),
      files: (e && Array.isArray(e.files) ? e.files : [])
        .filter((f) => f && f.bytes && f.name)
        .map((f) => ({ ...f, bytes: f.bytes instanceof Uint8Array ? f.bytes : new Uint8Array(f.bytes) })),
    }))
    .filter((e) => e.files.length);
}

/** Once an email has its files: takes them out of the store (the store goes when both are done). */
async function forgetPending(subject) {
  try {
    const pending = await idbGet(PENDING_KEY);
    const left = pendingEmails(pending).filter((e) => e.subject !== subject);
    if (!left.length) await idbDelete(PENDING_KEY);
    else await idbSet(PENDING_KEY, { ...pending, emails: left, files: left[0].files, subject: left[0].subject });
  } catch (error) {
    // harmless: it expires on its own
  }
}

async function initCompose() {
  showView('compose-view');
  // No Settings in a new email's panel: it would save Restack's settings from the copy this panel
  // loaded, which can be older than changes made in the main window since (like a template).
  const settings = document.querySelector('.foot-settings');
  if (settings) settings.hidden = true;
  setSummary('Attach a combined PDF to this email');
  let pending = null;
  try {
    pending = await idbGet(PENDING_KEY);
  } catch (error) {
    log(`Compose: couldn't check for a waiting PDF (${error.name}: ${error.message})`);
  }
  const text = $('compose-text');
  const button = $('btn-attach');
  const waiting = pendingEmails(pending);
  if (!waiting.length || Date.now() - pending.created > PENDING_MAX_AGE_MS) {
    text.textContent = 'No combined PDF is waiting. Open Restack on a thread, check the files you want, then choose Email.';
    button.hidden = true;
    log('Compose: no PDF waiting');
    return;
  }

  // Which email this is: the one with the same subject
  let subject = '';
  try {
    const field = Office.context.mailbox.item.subject;
    if (field && typeof field.getAsync === 'function') subject = String((await call((cb) => field.getAsync(cb))) || '');
  } catch (error) {
    subject = '';
  }
  const match = waiting.find((e) => e.subject && subject.trim() === e.subject.trim());
  const card = button.parentElement;
  const picks = [];

  const attach = async (entry) => {
    const names = entry.files.map((f) => f.name).join(', ');
    for (const b of [button, ...picks]) b.disabled = true;
    let done = 0;
    try {
      for (const file of entry.files) {
        text.textContent = `Attaching ${file.name}…`;
        await call((cb) => Office.context.mailbox.item.addFileAttachmentFromBase64Async(bytesToBase64(file.bytes), file.name, { isInline: false }, cb));
        done += 1;
      }
      await forgetPending(entry.subject);
      log(`Compose: attached ${names} ✔`);
      text.textContent = `Attached ${names}.`;
      button.hidden = true;
      for (const b of picks) b.remove();
    } catch (error) {
      log(`Compose: attaching failed after ${done} of ${entry.files.length} (${error.message})`);
      text.textContent = `Couldn't attach ${entry.files[done] ? entry.files[done].name : 'the PDF'}: ${error.message}`;
      for (const b of [button, ...picks]) b.disabled = false;
      if (!picks.length) button.hidden = false;
    }
  };

  if (match || waiting.length === 1) {
    const entry = match || waiting[0];
    const names = entry.files.map((f) => f.name).join(', ');
    button.addEventListener('click', () => attach(entry));
    if (pending.viaLink) {
      // Outlook attached these from the Azure link already; offer the button only as a backup.
      text.textContent = `This email should already have ${names} attached. If ${entry.files.length > 1 ? "they're" : "it's"} missing, click ${word('attachPdf')}.`;
      button.hidden = false;
      log('Compose: already attached through Azure; button offered as a backup');
      return;
    }
    if (match) {
      // The email the Email button just opened: attach without another click.
      log('Compose: this is the email Restack opened, attaching right away');
      await attach(entry);
      return;
    }
    text.textContent = `Ready to attach: ${names} (${humanSize(entry.files.reduce((sum, f) => sum + f.bytes.length, 0))}).`;
    button.hidden = false;
    log('Compose: PDF waiting');
    return;
  }

  // Two emails were opened, and this one's subject has changed since: pick whose files go on it
  text.textContent = pending.viaLink
    ? 'These emails should already have their files attached. If this one is missing them, pick which:'
    : 'Which files go on this email?';
  button.hidden = true;
  for (const entry of waiting) {
    const pick = el('button', 'btn', `${entry.subject}: ${entry.files.map((f) => f.name).join(', ')}`);
    pick.type = 'button';
    pick.addEventListener('click', () => attach(entry));
    picks.push(pick);
    card.append(pick);
  }
  log(`Compose: ${waiting.length} emails waiting; this one's subject matches neither`);
}

// --------------------------------------------------------------------------- start-up

let readUiWired = false;

function wireReadUi() {
  readUiWired = true;
  $('check-all').addEventListener('change', (event) => {
    const on = event.target.checked;
    for (const file of selectableFiles()) {
      const was = state.checked.has(file.id);
      if (on) state.checked.add(file.id);
      else state.checked.delete(file.id);
      if (on && !was && state.multiStop) assignStop(file); // in list order: the next free stop each
    }
    state.built.clear();
    normalizeOrder();
    renderList(); // places, stop pickers and lumper tags come and go with the ticks
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
  for (const kind of Object.keys(ACC_FIELDS)) {
    $(`acc-${kind}`).addEventListener('change', (event) => onAccessorialPick(kind, event.target.checked));
    $(`${kind}-at`).addEventListener('change', syncAccessorialBoxes);
  }
  for (const id of ['lum-total', 'oth-total']) $(id).addEventListener('change', onMoneyChange);
  $('load-form').addEventListener('input', () => { // (ticks fire this too: the hint follows the accessorials)
    $('acc-error').textContent = '';
    updateLoadHint();
  });
  $('tpl-back').addEventListener('click', () => (editing ? cancelEdit() : closeTemplates()));
  $('tpl-form').addEventListener('submit', (event) => onTemplateSubmit(event).catch(fail));
  $('tpl-cancel').addEventListener('click', cancelEdit);
  $('tpl-lumper-sep').addEventListener('change', syncLumperNameRow);
  $('tpl-acc-sep').addEventListener('change', syncAccRow);
  $('q-add').addEventListener('click', addQuestion);
  $('q-intro').addEventListener('input', (event) => { qDraft.intro = event.target.value; });
  $('q-list').addEventListener('keydown', (event) => { // Enter in a question box doesn't save the template (or add a line)
    if (event.key === 'Enter' && event.target.matches('input, textarea')) event.preventDefault();
  });
  $('tip-keep').addEventListener('click', () => setKeepAdding(true));
  $('keep-off').addEventListener('click', () => setKeepAdding(false));
  $('keep-restart').addEventListener('click', () => keepStartOver().catch(fail));
  $('keep-add-anyway').addEventListener('click', () => keepAddAnyway().catch(fail));
  $('keep-ask-restart').addEventListener('click', () => keepStartOver().catch(fail));
  $('quiz-back').addEventListener('click', quizBack);
  $('quiz-cancel').addEventListener('click', cancelQuiz);
  $('quiz').addEventListener('keydown', onQuizKeys);
  wireHowto();
  $('tpl-delete').addEventListener('click', () => onTemplateDelete().catch(fail));
  $('tpl-duplicate').addEventListener('click', onTemplateDuplicate);
  $('team-new').addEventListener('click', () => editTemplate('team', null));
  $('mine-new').addEventListener('click', () => editTemplate('mine', null));
  $('tpl-publish').addEventListener('click', () => onPublish().catch(fail));
  $('team-back').addEventListener('click', closeTeamSheet);
  $('rc-remove').addEventListener('click', () => onRcRemove().catch(fail));
  $('rw-save').addEventListener('click', () => onReceiptWordsSave($('rw-words').value.split(/[,;\n]+/)).catch(fail));
  $('rw-default').addEventListener('click', () => onReceiptWordsSave(DEFAULT_RECEIPT_WORDS).catch(fail));
  $('admin-form')?.addEventListener('submit', (event) => onAdminAdd(event).catch(fail));
  $('us-gen').addEventListener('click', () => {
    $('us-key').value = randomB64url(24);
    setStatusIn('us-status', 'New counting key made. Copy it into the Function App as RESTACK_USAGE_KEY.');
  });
  $('us-admin-gen').addEventListener('click', () => {
    $('us-admin').value = randomB64url(24);
    setStatusIn('us-status', 'New admin key made. Copy it into the Function App as RESTACK_USAGE_ADMIN_KEY.');
  });
  $('us-copy').addEventListener('click', () => copyFrom('us-key', 'us-status').catch(fail));
  $('us-admin-copy').addEventListener('click', () => copyFrom('us-admin', 'us-status').catch(fail));
  $('us-save').addEventListener('click', () => onUsageSave().catch(fail));
  $('us-test').addEventListener('click', () => onUsageTest().catch(fail));
  $('rc-leave-out').addEventListener('click', onRcLeaveOut);
  $('rc-keep').addEventListener('click', onRcKeep);
  $('pv-rc-sample').addEventListener('click', () => onUseRateConSample().catch(fail));
  $('az-gen').addEventListener('click', onAzureGenerate);
  $('az-copy').addEventListener('click', () => onAzureCopy().catch(fail));
  $('az-save').addEventListener('click', () => onAzureSave().catch(fail));
  $('az-test').addEventListener('click', () => onAzureTest().catch(fail));
  $('name-form').addEventListener('submit', onNameSubmit);
  $('file-name').addEventListener('input', updateNameHint);
  $('name-cancel').addEventListener('click', onNameCancel);
  $('pv-close').addEventListener('click', closePreview);
  $('pv-order-reset').addEventListener('click', resetPageOrder);
  $('pv-prev').addEventListener('click', () => stepPreview(-1));
  $('pv-next').addEventListener('click', () => stepPreview(1));
  $('pv-include').addEventListener('change', (event) => {
    const file = state.files[state.previewIndex];
    if (file) onRowToggle(file, event.target.checked);
  });
  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape' || CONSOLE) return;
    if (!$('quiz').hidden) cancelQuiz();
    else if (!$('howto').hidden) closeHowto(true);
    else if (!$('preview').hidden) closePreview();
    else if (!$('team-sheet').hidden) closeTeamSheet();
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
  const debounced = debounce(() => loadThread().catch(fail), 350);
  const reload = () => {
    noteSelection().catch(() => {}); // Keep adding: this click counts even if another comes before the list catches up
    debounced();
  };
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

// --------------------------------------------------------------------------- Settings
//
// The gear across from Test report opens Settings: Manage templates, Team (the team admin's own
// settings, folded up), Appearance, Text size and View (each a row that drops down its choices), and
// two switches. Each person's choices are saved in Outlook's roaming settings, so they follow them to
// other computers. mode.js puts last time's light or dark and text size on the page before anything
// shows, so the panel doesn't flash white in dark mode or jump in size.

const APPEARANCE_KEY = 'appearance';
const APPEARANCE_WORD = { light: 'light', dark: 'dark', match: 'matchOutlook' }; // their WORDS
let appearance = 'match';
let appearanceReady = false; // the saved choice is in (picks before that would be overwritten)

const TEXT_ZOOM = { small: 1, medium: 1.12, large: 1.25 }; // Small is the panel's usual size
const TEXT_WORD = { small: 'textSmall', medium: 'textMedium', large: 'textLarge' };
const VIEWS = ['large', 'medium', 'small', 'list', 'details', 'tiles'];
const VIEW_WORD = { large: 'viewLarge', medium: 'viewMedium', small: 'viewSmall', list: 'viewList', details: 'viewDetails', tiles: 'viewTiles' };

/** Each person's choices (roaming settings), with what they start as. */
const prefs = {
  textSize: 'small',
  view: 'medium',
  tagReceipts: false, // files named like receipts start out tagged Receipt (the team admin sets the words)
  keepAdding: false,  // each email clicked adds its files to the list, for people who don't use conversations
  pins: [],           // pinned templates (ids), in the order they want them
};

function loadPrefs() {
  const size = readRoaming('textSize');
  prefs.textSize = TEXT_ZOOM[size] ? size : 'small';
  const view = readRoaming('view');
  prefs.view = VIEWS.includes(view) ? view : 'medium';
  prefs.tagReceipts = readRoaming('tagReceipts') === true;
  prefs.keepAdding = readRoaming('keepAdding') === true;
  const pins = readRoaming('pins');
  prefs.pins = Array.isArray(pins) ? pins.map(String).filter((id) => /^[A-Za-z0-9_-]{1,40}$/.test(id)).slice(0, 50) : [];
  applyTextSize(prefs.textSize);
}

function savePref(name, value) {
  prefs[name] = value;
  writeRoaming(name, value).catch((error) => log(`Settings: couldn't save ${name} (${error.message})`));
}

/** How bright a #rrggbb color is, from 0 (black) to 1 (white), or null if it isn't one. */
function luminance(color) {
  let hex = String(color || '').trim().replace(/^#/, '');
  if (hex.length === 3) hex = hex.replace(/./g, '$&$&');
  if (!/^[0-9a-f]{6}$/i.test(hex)) return null;
  const [r, g, b] = [0, 2, 4].map((i) => {
    const v = parseInt(hex.slice(i, i + 2), 16) / 255;
    return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Whether Outlook is dark (by its background color), or the computer is when Outlook doesn't say. */
function outlookIsDark(theme) {
  let colors = theme || null;
  if (!colors) {
    try {
      colors = typeof Office !== 'undefined' && Office.context ? Office.context.officeTheme : null;
    } catch (error) {
      colors = null;
    }
  }
  const light = colors ? luminance(colors.bodyBackgroundColor) : null;
  if (light !== null) return light < 0.35;
  return !!(window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches);
}

/** Puts the panel in light or dark for the choice, and remembers which for next time (mode.js). */
function applyMode(theme) {
  const mode = appearance === 'dark' || (appearance === 'match' && outlookIsDark(theme)) ? 'dark' : 'light';
  document.documentElement.dataset.mode = mode;
  writeSetting('mode', mode);
  markAppearance();
  return mode;
}

/** The same, fading the colors over instead of switching at once. */
function fadeMode(theme) {
  const root = document.documentElement;
  const before = root.dataset.mode;
  root.classList.add('mode-fade');
  const mode = applyMode(theme);
  clearTimeout(fadeMode.timer);
  fadeMode.timer = setTimeout(() => root.classList.remove('mode-fade'), 450);
  return before !== mode;
}

/** Saves the choice for this mailbox (roaming settings). */
function saveAppearance() {
  writeRoaming(APPEARANCE_KEY, appearance).catch((error) => log(`Appearance: couldn't save (${error.message})`));
}

/** A pick in Settings: the panel fades over to it, the list folds up, and Settings stays open. */
function pickAppearance(choice) {
  if (!appearanceReady || choice === appearance || !APPEARANCE_WORD[choice]) return;
  appearance = choice;
  markAppearance();
  fadeMode();
  saveAppearance();
  log(`Appearance: ${word(APPEARANCE_WORD[choice])} (${document.documentElement.dataset.mode})`);
}

// ---- Text size: the whole panel bigger (list, buttons, forms, preview), saved per person

/** How much the panel is zoomed for the text size (1 at Small). */
function pageZoom() {
  return Number(document.documentElement.style.zoom) || 1;
}

/** The panel's size in its own (zoomed) pixels, for CSS in place of vw and vh, which don't follow the zoom. */
function measurePage() {
  const zoom = pageZoom();
  const root = document.documentElement.style;
  root.setProperty('--page-w', `${Math.round(window.innerWidth / zoom)}px`);
  root.setProperty('--page-h', `${Math.round(window.innerHeight / zoom)}px`);
}

function applyTextSize(size) {
  const zoom = TEXT_ZOOM[size] || 1;
  document.documentElement.style.zoom = zoom === 1 ? '' : String(zoom);
  writeSetting('size', size === 'small' ? null : size); // for mode.js, next time
  measurePage();
  markChoice('size', size);
}

function pickTextSize(size) {
  if (!TEXT_ZOOM[size] || size === prefs.textSize) return;
  savePref('textSize', size);
  applyTextSize(size);
  log(`Text size: ${word(TEXT_WORD[size])}`);
}

// ---- View: how the list shows files (like File Explorer's View menu), saved per person

function pickView(view) {
  if (!VIEWS.includes(view) || view === prefs.view) return;
  savePref('view', view);
  markChoice('view', view);
  log(`View: ${word(VIEW_WORD[view])}`);
  if (readUiWired) renderList();
}

// ---- the drop-down rows and the Team group

const SETTINGS_LISTS = ['team-menu', 'appearance-list', 'size-list', 'view-list'];

/** The row in Settings that opens a list. */
function listRow(list) {
  return document.querySelector(`#settings-pop [aria-controls="${list.id}"]`);
}

/** The list that's dropped down in Settings, or null. */
function openSettingsList() {
  return SETTINGS_LISTS.map((id) => $(id)).find((list) => list && !list.hidden) || null;
}

/** Drops down (or folds up) one of Settings' lists; any other folds up. Optionally focuses its current choice, as a drop-down does. */
function showSettingsList(list, open, focusChoice) {
  for (const id of SETTINGS_LISTS) {
    const other = $(id);
    if (!other) continue;
    const on = other === list && open;
    other.hidden = !on;
    const row = listRow(other);
    if (row) row.setAttribute('aria-expanded', String(on));
  }
  if (!open || !focusChoice) return;
  const target = list.querySelector('[aria-selected="true"]') || [...list.querySelectorAll('button')].find((b) => !b.disabled && !b.hidden);
  if (target) target.focus();
}

/** Marks a drop-down's choice: checked in its list, and shown (its icon and name) on its row. */
function markChoice(key, value) {
  const list = $(`${key}-list`);
  if (!list) return;
  let current = null;
  for (const option of list.querySelectorAll('[role="option"]')) {
    const on = option.dataset.value === value;
    option.setAttribute('aria-selected', String(on));
    option.tabIndex = on ? 0 : -1;
    if (on) current = option;
  }
  if (!current) return;
  const icon = $(`set-${key}-icon`);
  const glyph = current.querySelector('.opt-icon');
  if (icon && glyph) icon.replaceChildren(...[...glyph.childNodes].map((node) => node.cloneNode(true)));
  const now = $(`set-${key}-now`);
  if (now) now.textContent = current.querySelector('.opt-name').textContent;
}

function markAppearance() {
  markChoice('appearance', appearance);
}

function markSwitch(id, on) {
  const button = $(id);
  if (button) button.setAttribute('aria-checked', String(!!on));
}

/** Opens Settings above its button. The admin's Team group and the list's settings show only where they apply. */
function openSettings() {
  const pop = $('settings-pop');
  const templates = $('set-templates');
  templates.hidden = !(readUiWired && canMakeOwn());
  if (!templates.hidden) {
    templates.querySelector('.mi-name').textContent = word(teamTemplates().length || state.myTemplates.length ? 'manageTemplates' : 'createTemplate');
  }
  const admin = readUiWired && isAdmin();
  $('set-team-group').hidden = !admin;
  if (admin) renderTeamMenu();
  teamNote('');
  $('set-sep-look').hidden = templates.hidden && !admin;
  for (const id of ['set-size', 'set-view', 'set-sep-prefs', 'set-tag-receipts', 'set-keep-adding']) $(id).hidden = !readUiWired;
  markAppearance();
  markChoice('size', prefs.textSize);
  markChoice('view', prefs.view);
  markSwitch('set-tag-receipts', prefs.tagReceipts);
  markSwitch('set-keep-adding', prefs.keepAdding);
  showSettingsList(null, false); // compact each time: the choices drop down from their rows
  pop.hidden = false;
  $('btn-settings').setAttribute('aria-expanded', 'true');
  const first = [templates, $('set-team'), $('set-appearance')].find((b) => b && !b.hidden && b.offsetParent !== null);
  (first || $('set-appearance')).focus();
}

function closeSettings(returnFocus) {
  const pop = $('settings-pop');
  if (!pop || pop.hidden) return;
  pop.hidden = true;
  showSettingsList(null, false);
  $('btn-settings').setAttribute('aria-expanded', 'false');
  if (returnFocus) $('btn-settings').focus();
}

/** A short note under the Team group (after Publish, Share install file and so on). */
function teamNote(text, tone) {
  if (CONSOLE) {
    consoleNote(text, tone);
    return;
  }
  const note = $('team-menu-note');
  if (!note) return;
  note.textContent = text || '';
  note.hidden = !text;
  note.className = `set-note${tone ? ` ${tone}` : ''}`;
}

/** The Team group's states: unpublished changes, and each setting at a glance. */
function renderTeamMenu() {
  if (CONSOLE) {
    renderConsolePublish();
    return;
  }
  if (!$('set-team-group') || !isAdmin()) return;
  const pending = !!team.draft && !team.justPublished;
  const now = $('set-team-now');
  now.textContent = pending ? 'Changes to publish' : '';
  now.classList.toggle('pending', pending);
  const state = $('team-publish-state');
  state.textContent = pending ? 'Changes to publish' : (team.justPublished ? 'Downloaded' : 'Up to date');
  state.classList.toggle('on', pending);
  $('team-publish').disabled = !team.draft;
  const settings = teamSettings();
  markSwitch('team-rep-templates', settings && settings.repTemplates);
  markSwitch('team-sent-with', settings && settings.sentWith);
  const admins = $('team-admins-state');
  if (admins) admins.textContent = String(teamAdmins().length);
  $('team-rc-state').textContent = teamRateCon() ? 'On' : 'Off';
  $('team-azure-state').textContent = teamAzure() ? 'On' : 'Off';
  $('team-receipts-state').textContent = receiptWords().join(', ') || 'None';
  $('team-usage-state').textContent = teamUsage() ? 'On' : 'Off';
}

/** Arrow keys, Home and End move between a list's buttons. */
function onSettingsListKeys(event) {
  const list = event.currentTarget;
  const items = [...list.querySelectorAll('button')].filter((b) => !b.disabled && !b.hidden);
  if (!items.length) return;
  const at = Math.max(0, items.indexOf(document.activeElement));
  const next = {
    ArrowDown: items[(at + 1) % items.length],
    ArrowUp: items[(at + items.length - 1) % items.length],
    Home: items[0],
    End: items[items.length - 1],
  }[event.key];
  if (!next) return;
  event.preventDefault();
  next.focus();
}

function wireSettings() {
  const box = document.querySelector('.foot-settings');
  const pop = $('settings-pop');
  const lists = SETTINGS_LISTS.map((id) => $(id));
  if (!box || !pop || !$('btn-settings') || !$('set-templates') || lists.some((list) => !list || !listRow(list))) {
    log('Settings: not on this page (an older copy of taskpane.html)');
    return;
  }
  $('btn-settings').addEventListener('click', () => (pop.hidden ? openSettings() : closeSettings(true)));
  $('set-templates').addEventListener('click', () => {
    closeSettings(false);
    templatesOpener = $('btn-settings');
    openTemplates();
  });
  // Each row drops its list down (click, Enter, Space or the arrow keys); a pick folds it back up
  for (const list of lists) {
    const row = listRow(list);
    row.addEventListener('click', () => {
      showSettingsList(list, list.hidden, true);
      if (list.id === 'team-menu') teamNote('');
    });
    row.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
      event.preventDefault();
      showSettingsList(list, true, true);
    });
    list.addEventListener('keydown', onSettingsListKeys);
  }
  const picks = { 'appearance-list': pickAppearance, 'size-list': pickTextSize, 'view-list': pickView };
  for (const [id, pick] of Object.entries(picks)) {
    const list = $(id);
    list.addEventListener('click', (event) => {
      const option = event.target.closest('[role="option"]');
      if (!option) return;
      listRow(list).focus(); // first, so focus isn't lost when the list folds up
      showSettingsList(list, false);
      pick(option.dataset.value);
    });
  }
  // The switches
  $('set-tag-receipts').addEventListener('click', () => setTagReceipts(!prefs.tagReceipts));
  $('set-keep-adding').addEventListener('click', () => setKeepAdding(!prefs.keepAdding));
  // Team (the admin's)
  $('team-publish').addEventListener('click', () => onPublish().catch((error) => teamNote(`Couldn't publish: ${error.message}`, 'error')));
  $('team-share').addEventListener('click', () => onShareInstall().catch((error) => teamNote(error.message, 'error')));
  $('team-share-beta').addEventListener('click', () => onShareBeta().catch((error) => teamNote(error.message, 'error')));
  $('team-rep-templates').addEventListener('click', () => {
    const on = $('team-rep-templates').getAttribute('aria-checked') !== 'true';
    onRepTemplatesToggle(on).catch((error) => teamNote(error.message, 'error'));
  });
  $('team-sent-with')?.addEventListener('click', () => {
    const on = $('team-sent-with').getAttribute('aria-checked') !== 'true';
    onSentWithToggle(on).catch((error) => teamNote(error.message, 'error'));
  });
  for (const [id, which] of [['team-open-admins', 'admins'], ['team-open-rc', 'rc'], ['team-open-azure', 'azure'], ['team-open-receipts', 'receipts'], ['team-open-usage', 'usage']]) {
    $(id)?.addEventListener('click', () => openTeamSheet(which));
  }
  $('team-console').addEventListener('click', openConsole);
  box.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape' || pop.hidden) return;
    event.stopPropagation(); // just Settings, not the screen behind it
    const open = openSettingsList();
    if (open) { // the dropped-down list first, then Settings
      listRow(open).focus();
      showSettingsList(open, false);
    } else {
      closeSettings(true);
    }
  });
  box.addEventListener('focusout', (event) => {
    if (!pop.hidden && event.relatedTarget && !box.contains(event.relatedTarget)) closeSettings(false);
  });
  document.addEventListener('pointerdown', (event) => {
    if (!pop.hidden && !event.target.closest('.foot-settings')) closeSettings(false);
  });
  window.addEventListener('blur', () => closeSettings(false)); // a click in Outlook, outside the panel
  window.addEventListener('resize', measurePage);
  markAppearance();
  markChoice('size', prefs.textSize);
  markChoice('view', prefs.view);
  if (window.matchMedia) {
    const media = matchMedia('(prefers-color-scheme: dark)');
    const follow = () => { if (appearance === 'match') fadeMode(); };
    if (media.addEventListener) media.addEventListener('change', follow);
    else if (media.addListener) media.addListener(follow);
  }
}

/** Once Outlook is ready: the saved choice, and Outlook's theme (followed while the panel is open). */
function loadAppearance() {
  const saved = readRoaming(APPEARANCE_KEY);
  appearance = ['light', 'dark', 'match'].includes(saved) ? saved : 'match';
  appearanceReady = true;
  const mode = applyMode();
  log(`Appearance: ${word(APPEARANCE_WORD[appearance])} (${mode})`);
  try {
    const mailbox = Office.context.mailbox;
    if (Office.EventType.OfficeThemeChanged && mailbox && typeof mailbox.addHandlerAsync === 'function') {
      mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, (event) => {
        if (appearance === 'match' && fadeMode(event && event.officeTheme)) log(`Appearance: Outlook changed to ${document.documentElement.dataset.mode}`);
      }, (result) => {
        if (result && result.status === Office.AsyncResultStatus.Failed) log("Appearance: can't follow Outlook's theme while open (it will next time)");
      });
    }
  } catch (error) {
    // older Outlook: the panel follows Outlook's theme the next time it opens
  }
}

/**
 * Loads the panel again straight from the site (the extra ?fresh= in the address gets past any copy
 * Outlook kept). At most twice in a row, counted in the address (?tries=), so a site that's halfway
 * through publishing can't keep the panel reloading: after that it carries on with what it has.
 */
function reloadFromSite(target, why) {
  let url;
  try {
    url = new URL(window.location.href);
  } catch (error) {
    return false;
  }
  const tries = Number(url.searchParams.get('tries')) || 0;
  if (tries >= 2) {
    log(`Start: ${why}, even after loading again from the site`);
    return false;
  }
  url.searchParams.set('fresh', target);
  url.searchParams.set('tries', String(tries + 1));
  try {
    window.location.replace(url.href);
  } catch (error) {
    return false;
  }
  setTimeout(startNow, 5000); // only if the reload didn't happen
  return true;
}

/**
 * After an update, Outlook can show a copy of taskpane.html it kept from before, next to the new
 * taskpane.js. The page says which version it is; when that isn't this script's, load it again from the site.
 */
function reloadIfOldPage() {
  const meta = document.querySelector('meta[name="restack-version"]');
  const page = meta ? meta.getAttribute('content') : '';
  if (page === VERSION) return false;
  return reloadFromSite(VERSION, `the page is ${page || 'an older one'}, the script ${VERSION}`);
}

/**
 * New versions show up as soon as they're uploaded (#30). Outlook can open the panel from a copy of the page
 * and script it kept (GitHub Pages lets it keep them for up to 10 minutes), so before starting, the panel asks
 * the site which version is up, and loads that one when it isn't this one. It waits for an answer at most
 * SITE_CHECK_MS; offline or slow, the panel starts as it is. The beta asks its own folder.
 */
const SITE_CHECK_MS = 1500;

async function siteVersion() {
  const stop = typeof AbortController === 'function' ? new AbortController() : null;
  const timer = setTimeout(() => { if (stop) stop.abort(); }, SITE_CHECK_MS);
  try {
    // (no-cache: Outlook asks the site whether its copy is still current, a quick "yes" when it is)
    const response = await fetch(new URL('taskpane.html', window.location.href).href, { cache: 'no-cache', signal: stop ? stop.signal : undefined });
    if (!response.ok) return '';
    const match = (await response.text()).match(/<meta\s+name="restack-version"\s+content="([^"]+)"/);
    return match ? match[1] : '';
  } catch (error) {
    return ''; // offline, too slow, or blocked: start with what's here
  } finally {
    clearTimeout(timer);
  }
}

async function start() {
  if (reloadIfOldPage()) return;
  const site = await siteVersion();
  if (site && site !== VERSION && reloadFromSite(site, `the site has ${site}, this is ${VERSION}`)) return;
  startNow();
}

function startNow() {
  if (startNow.done) return;
  startNow.done = true;
  applyWords();
  measurePage();
  const badge = $('beta-badge');
  if (badge) badge.hidden = !BETA;
  try {
    wireSettings();
  } catch (error) {
    log(`Settings: couldn't set up (${error.message})`);
  }
  $('btn-copy-report').addEventListener('click', () => copyReport());
  window.addEventListener('error', (event) => log(`Script error: ${event.message}`));
  window.addEventListener('unhandledrejection', (event) => log(`Unhandled: ${event.reason && (event.reason.message || event.reason)}`));
  if (CONSOLE) {
    collectEnvironment();
    startConsole().catch(fail);
    return;
  }
  if (typeof Office === 'undefined') {
    appearanceReady = true; // nothing saved to wait for
    showMessage("Outlook's add-in library didn't load. Check the internet connection, then close and reopen this panel.");
    setSummary('');
    return;
  }
  Office.onReady(async (info) => {
    try {
      collectEnvironment();
      loadAppearance();
      loadPrefs();
      if (!info || info.host !== Office.HostType.Outlook) {
        setSummary('');
        showMessage('Open this from Outlook: select an email, then click Restack on the ribbon.');
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
      wireReadUi();
      registerHandlers();
      renderKeepBar();
      loadThread().catch(fail);
      flushUsage().catch(() => {}); // counts left from last time
    } catch (error) {
      fail(error);
    }
  });
}

start().catch(fail);
