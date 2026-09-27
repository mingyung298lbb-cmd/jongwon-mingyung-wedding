const SPREADSHEET_ID = "여기에_구글시트_ID";
const RSVP_SHEET = "RSVP";
const COMMENT_SHEET = "댓글";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    if (data.action === "rsvp") return saveRsvp_(data);
    if (data.action === "comment") return saveComment_(data);
    return json_({ ok:false, error:"invalid_action" });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function doGet(e) {
  const action = (e.parameter.action || "").toLowerCase();
  if (action !== "comments") return json_({ ok:false, error:"invalid_action" });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(COMMENT_SHEET);
  if (!sh || sh.getLastRow() < 4) return json_({ ok:true, comments:[] });

  const values = sh.getRange(4, 1, sh.getLastRow() - 3, 5).getValues();
  const tz = Session.getScriptTimeZone() || "Asia/Seoul";
  const comments = values
    .filter(r => String(r[3]).trim() === "공개")
    .map(r => ({
      name: String(r[1] || "").trim(),
      message: String(r[2] || "").trim(),
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy.MM.dd") : String(r[0] || "")
    }))
    .filter(x => x.name && x.message)
    .slice(-50)
    .reverse();

  return json_({ ok:true, comments });
}

function saveRsvp_(data) {
  const side = clean_(data.side, 20);
  const name = clean_(data.name, 30);
  const phone = normalizePhone_(data.phone);
  const attendance = clean_(data.attendance, 10);
  const guestCount = attendance === "참석" ? Math.max(1, Math.min(4, Number(data.guestCount) || 1)) : 0;

  if (!["신랑측","신부측"].includes(side)) return json_({ok:false,error:"invalid_side"});
  if (!name || phone.length < 10) return json_({ok:false,error:"invalid_identity"});
  if (!["참석","불참"].includes(attendance)) return json_({ok:false,error:"invalid_attendance"});

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(RSVP_SHEET);
  if (!sh) return json_({ok:false,error:"missing_rsvp_sheet"});

  const key = name.toLowerCase().replace(/\s+/g,"") + "|" + phone;
  const lastRow = sh.getLastRow();
  let targetRow = 0;

  if (lastRow >= 4) {
    const keys = sh.getRange(4, 7, lastRow - 3, 1).getValues().flat();
    const idx = keys.findIndex(v => String(v) === key);
    if (idx >= 0) targetRow = idx + 4;
  }

  const row = [[new Date(), side, name, formatPhone_(phone), attendance, guestCount, key, ""]];
  if (targetRow) {
    sh.getRange(targetRow, 1, 1, 8).setValues(row);
    return json_({ok:true,updated:true});
  }

  sh.getRange(Math.max(4, lastRow + 1), 1, 1, 8).setValues(row);
  return json_({ok:true,updated:false});
}

function saveComment_(data) {
  const name = clean_(data.name, 30);
  const message = clean_(data.message, 200);
  if (!name || !message) return json_({ok:false,error:"invalid_comment"});

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(COMMENT_SHEET);
  if (!sh) return json_({ok:false,error:"missing_comment_sheet"});

  const row = Math.max(4, sh.getLastRow() + 1);
  sh.getRange(row, 1, 1, 5).setValues([[new Date(), name, message, "숨김", ""]]);
  return json_({ok:true});
}

function clean_(value, maxLen) {
  return String(value == null ? "" : value).replace(/[<>]/g,"").trim().slice(0, maxLen);
}

function normalizePhone_(value) {
  return String(value || "").replace(/\D/g,"").slice(0, 11);
}

function formatPhone_(digits) {
  if (digits.length === 11) return digits.replace(/(\d{3})(\d{4})(\d{4})/,"$1-$2-$3");
  if (digits.length === 10) return digits.replace(/(\d{3})(\d{3})(\d{4})/,"$1-$2-$3");
  return digits;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
