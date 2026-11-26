/**
 * קליטת לידים מדף הנחיתה של המכינה לתוך Google Sheet.
 *
 * הגיליון: https://docs.google.com/spreadsheets/d/1ElYjhIf-gdHATlduU29En3YlAatmh_-nWMsCgm__WHI/edit
 *
 * התקנה:
 * 1. פותחים את הגיליון > Extensions > Apps Script.
 * 2. מדביקים את הקוד הזה במקום הקוד הקיים ושומרים.
 * 3. Deploy > New deployment > Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. מעתיקים את ה-URL שמסתיים ב-/exec ומדביקים אותו ב-ENDPOINT בקובץ index.html.
 * 5. (רשות) ממלאים NOTIFY_EMAIL כדי לקבל מייל על כל ליד חדש.
 */

const SPREADSHEET_ID = '1ElYjhIf-gdHATlduU29En3YlAatmh_-nWMsCgm__WHI';
const NOTIFY_EMAIL = ''; // לדוגמה: 'Mechina@savion.huji.ac.il'

const HEADERS = [
  'תאריך', 'שם פרטי', 'שם משפחה', 'טלפון', 'אימייל', 'הסכמה לדיוור',
  'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'referrer', 'עמוד'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      clean_(d.firstName), clean_(d.lastName),
      "'" + clean_(d.phone),          // keep the leading 0
      clean_(d.email), clean_(d.consent),
      clean_(d.utm_source), clean_(d.utm_medium), clean_(d.utm_campaign), clean_(d.utm_content),
      clean_(d.referrer), clean_(d.page)
    ]);
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'ליד חדש מדף הנחיתה של המכינה',
        d.firstName + ' ' + d.lastName + '\n' + d.phone + '\n' + d.email +
        '\nמקור: ' + (d.utm_source || 'ישיר'));
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

// Strip leading =,+,-,@ so a submitted value can't run as a spreadsheet formula
function clean_(v) {
  return String(v == null ? '' : v).slice(0, 300).replace(/^[=+\-@]+/, '');
}
