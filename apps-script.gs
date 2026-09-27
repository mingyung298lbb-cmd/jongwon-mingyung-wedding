/**
 * Wedding RSVP + Guest Comments backend for Google Apps Script.
 * Spreadsheet stays private. The web app only exposes:
 *  - POST action=rsvp       -> create/update one RSVP row by normalized name+phone key
 *  - POST action=comment    -> append a comment with 공개여부=FALSE
 *  - GET  action=comments   -> return only comments whose 공개여부 is TRUE
 *
 * 1) Create a private Google Sheet.
 * 2) Open Extensions > Apps Script and paste this file.
 * 3) Run setupWeddingSheets() once.
 * 4) Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 5) Copy the /exec URL into WEDDING_API_URL in index.html.
 */

const RSVP_SHEET = 'RSVP';
const COMMENT_SHEET = '댓글';

function setupWeddingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let rsvp = ss.getSheetByName(RSVP_SHEET);
  if (!rsvp) rsvp = ss.insertSheet(RSVP_SHEET);
  if (rsvp.getLastRow() === 0) {
    rsvp.appendRow([
      '고유키',
      '구분',
      '성함',
      '연락처',
      '참석여부',
      '참석인원',
      '최초제출',
      '최종수정'
    ]);
    rsvp.setFrozenRows(1);
  }

  let comments = ss.getSheetByName(COMMENT_SHEET);
  if (!comments) comments = ss.insertSheet(COMMENT_SHEET);
  if (comments.getLastRow() === 0) {
    comments.appendRow([
      'ID',
      '이름',
      '댓글',
      '제출시간',
      '공개여부'
    ]);
    comments.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    setupWeddingSheets();
    const p = e.parameter || {};
    const action = String(p.action || '');

    if (action === 'rsvp') {
      saveRsvp_(p);
      return text_('ok');
    }

    if (action === 'comment') {
      saveComment_(p);
      return text_('ok');
    }

    return text_('invalid action');
  } catch (err) {
    console.error(err);
    return text_('error');
  }
}

function doGet(e) {
  const p = e.parameter || {};
  if (String(p.action || '') !== 'comments') {
    return text_('ok');
  }

  const prefix = safeCallback_(String(p.prefix || ''));
  if (!prefix) return text_('invalid callback');

  setupWeddingSheets();
  const comments = getPublicComments_();
  return ContentService
    .createTextOutput(prefix + '(' + JSON.stringify({ comments }) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function saveRsvp_(p) {
  const side = clean_(p.side, 10);
  const name = clean_(p.name, 20);
  const phone = normalizePhone_(p.phone);
  const attendance = clean_(p.attendance, 10);
  const guests = attendance === '참석' ? clean_(p.guests || '1', 4) : '0';

  if (!['신랑측', '신부측'].includes(side)) throw new Error('invalid side');
  if (!name) throw new Error('missing name');
  if (phone.length < 10) throw new Error('invalid phone');
  if (!['참석', '불참'].includes(attendance)) throw new Error('invalid attendance');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RSVP_SHEET);
  const key = makeKey_(name, phone);
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const lastRow = sheet.getLastRow();
    let foundRow = 0;

    if (lastRow >= 2) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === key) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow) {
      const firstSubmitted = sheet.getRange(foundRow, 7).getValue() || now;
      sheet.getRange(foundRow, 1, 1, 8).setValues([[
        key, side, name, phone, attendance, guests, firstSubmitted, now
      ]]);
    } else {
      sheet.appendRow([key, side, name, phone, attendance, guests, now, now]);
    }
  } finally {
    lock.releaseLock();
  }
}

function saveComment_(p) {
  const name = clean_(p.name, 20);
  const message = clean_(p.message, 200);

  if (!name) throw new Error('missing name');
  if (!message) throw new Error('missing message');

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(COMMENT_SHEET);

  const id = Utilities.getUuid();
  sheet.appendRow([id, name, message, new Date(), false]);
}

function getPublicComments_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(COMMENT_SHEET);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';

  return values
    .filter(row => row[4] === true || String(row[4]).toUpperCase() === 'TRUE')
    .slice(-30)
    .reverse()
    .map(row => ({
      name: String(row[1] || ''),
      message: String(row[2] || ''),
      date: row[3] instanceof Date
        ? Utilities.formatDate(row[3], tz, 'yyyy.MM.dd')
        : ''
    }));
}

function makeKey_(name, phone) {
  return String(name).replace(/\s+/g, '').toLowerCase() + '|' + phone;
}

function normalizePhone_(value) {
  return String(value || '').replace(/\D/g, '').slice(0, 11);
}

function clean_(value, maxLen) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .trim()
    .slice(0, maxLen);
}

function safeCallback_(name) {
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(name) ? name : '';
}

function text_(value) {
  return ContentService
    .createTextOutput(String(value))
    .setMimeType(ContentService.MimeType.TEXT);
}
