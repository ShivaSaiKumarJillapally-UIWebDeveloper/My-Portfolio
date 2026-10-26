// Backend for Edla Meghana website "Request a Call" form.
// Paste this whole file into the Apps Script editor.
// Do not upload this file to Vercel.

// ---------- SETTINGS ----------
const SPREADSHEET_ID = '1S4mHkKAAe95J3sq4KVwNVdttZCT5FCtCY5bXt6c_9Kc';
const SHEET_NAME = '';
// SHEET_NAME: leave empty to use the first tab.
const TIMEZONE = 'Asia/Kolkata';
const NOTIFY_EMAIL = 'myfutureworks1@gmail.com';
// NOTIFY_EMAIL: set to empty quotes to turn emails off.
const HEADERS = [
  'Enquiry ID',
  'Submission Date and Time',
  'Customer Name',
  'Mobile Number',
  'Project Requirement',
  'Enquiry Status',
  'Follow-up Date',
  'Notes'
];
const SECONDS_BETWEEN_SAME_NUMBER = 120;
const MAX_PER_HOUR = 60;

const MSG_NAME = 'Please enter your name.';
const MSG_PHONE =
  'Please enter a valid 10-digit Indian mobile number.';
const MSG_REQ =
  'Please tell us a little about your project.';
const MSG_DUP =
  'We already received your enquiry. ' +
  'We will contact you shortly.';
const MSG_BUSY =
  'We are receiving many requests right now. ' +
  'Please try again later.';
const MSG_ERR =
  'Sorry, something went wrong. ' +
  'Please try again in a moment.';

// ---------- ENTRY POINTS ----------
function doGet() {
  return json_({ ok: true, message: 'Service is running.' });
}

function doPost(e) {
  const raw = e ? (e.postData ? e.postData.contents : '') : '';
  if (!raw) return fail_('Invalid request.');
  if (gt_(raw.length, 5000)) return fail_('Invalid request.');
  let d;
  try {
    d = JSON.parse(raw);
  } catch (err) {
    return fail_('Invalid request.');
  }
  return handle_(d);
}

function handle_(d) {
  const lock = LockService.getScriptLock();
  try {
    // Honeypot: bots fill this hidden field.
    if (d.website) return json_({ ok: true });

    const name = clean_(d.name, 80);
    const requirement = clean_(d.requirement, 1000);
    const mobile = normalisePhone_(d.mobile);
    const token = String(d.token || '').slice(0, 60);

    if (lt_(name.length, 2)) return fail_(MSG_NAME);
    if (!mobile) return fail_(MSG_PHONE);
    if (lt_(requirement.length, 5)) return fail_(MSG_REQ);

    lock.waitLock(15000);
    const cache = CacheService.getScriptCache();

    // Duplicate protection
    if (token) {
      if (cache.get('t_' + token)) return json_({ ok: true });
    }
    if (cache.get('p_' + mobile)) return fail_(MSG_DUP);

    // Hourly cap
    const hourKey = 'h_' + Utilities.formatDate(
      new Date(), TIMEZONE, 'yyyyMMddHH');
    const count = Number(cache.get(hourKey) || 0);
    if (!lt_(count, MAX_PER_HOUR)) return fail_(MSG_BUSY);

    const sheet = getSheet_();
    ensureHeaders_(sheet);

    const now = new Date();
    const id = 'ENQ-' +
      Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd') + '-' +
      Utilities.getUuid().slice(0, 6).toUpperCase();
    const when = Utilities.formatDate(
      now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, 8);
    const fmt = ['@', '@', '@', '@', '@', '@', '@', '@'];
    range.setNumberFormats([fmt]);
    range.setValues([
      [id, when, name, mobile, requirement, 'New', '', '']
    ]);
    SpreadsheetApp.flush();

    if (token) cache.put('t_' + token, '1', 600);
    cache.put('p_' + mobile, '1', SECONDS_BETWEEN_SAME_NUMBER);
    cache.put(hourKey, String(count + 1), 3600);

    notify_(id, name, mobile, requirement, when);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return fail_(MSG_ERR);
  } finally {
    try {
      lock.releaseLock();
    } catch (err2) {
      console.error(err2);
    }
  }
}

// ---------- HELPERS ----------
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (SHEET_NAME) return ss.getSheetByName(SHEET_NAME);
  return ss.getSheets()[0];
}

// Adds headers only if the sheet is completely empty.
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    const r = sheet.getRange(1, 1, 1, HEADERS.length);
    r.setValues([HEADERS]);
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function normalisePhone_(v) {
  let s = String(v || '').replace(/[\s\-()]/g, '');
  s = s.replace(/^\+?91/, '').replace(/^0/, '');
  if (/^[6-9]\d{9}$/.test(s)) return '+91' + s;
  return '';
}

function clean_(v, max) {
  return String(v || '')
    .replace(/[\u0000-\u001F\u007F]+/g, ' ')
    .trim()
    .slice(0, max);
}

function notify_(id, name, mobile, requirement, when) {
  if (!NOTIFY_EMAIL) return;
  try {
    const body =
      'ID: ' + id + '\n' +
      'Time: ' + when + '\n' +
      'Name: ' + name + '\n' +
      'Mobile: ' + mobile + '\n\n' +
      'Requirement:\n' + requirement;
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New website enquiry: ' + name,
      body: body
    });
  } catch (err) {
    console.error(err);
  }
}

// Comparison helpers, written without angle brackets.
function lt_(a, b) {
  return Math.sign(a - b) === -1;
}
function gt_(a, b) {
  return Math.sign(a - b) === 1;
}

function json_(o) {
  const out = ContentService.createTextOutput(JSON.stringify(o));
  return out.setMimeType(ContentService.MimeType.JSON);
}
function fail_(m) {
  return json_({ ok: false, message: m });
}

// Run once from the editor to give permissions.
function authoriseOnce() {
  getSheet_();
  if (NOTIFY_EMAIL) MailApp.getRemainingDailyQuota();
}

// Run from the editor to test saving without the website.
function testSave() {
  const out = handle_({
    name: 'Test Customer',
    mobile: '9876543210',
    requirement: 'Test enquiry from the editor',
    token: 'test' + Date.now()
  });
  Logger.log(out.getContent());
}
