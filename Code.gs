/*  Lazer fizikasi va Elektr-magnetizm virtual laboratoriyalari: natijalarni yig‘ish serveri (bitta server ikkalasiga xizmat qiladi)
    Google Apps Script. O‘rnatish tartibi SOZLASH.md faylida.           */

const PASSWORD = 'UrDU-2026';      // <-- o‘qituvchi parolini shu yerda o‘zgartiring
const SHEET_NAME = 'Natijalar';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', 'vaqt', 'talaba', 'variant', 'ish', 'ish nomi', 'tekshiruv (to‘g‘ri/jami)', 'baho', 'izoh', 'checks_json', 'journals_json', 'qurilma', 'fan']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) { return out_({ ok: true, info: 'Lazer-lab natijalar serveri ishlayapti.' }); }

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = sheet_();
    if (d.action === 'submit') {
      if (!d.name || String(d.name).trim().length < 5) return out_({ ok: false, error: 'Talaba ismi kiritilmagan.' });
      const id = Utilities.getUuid().slice(0, 8).toUpperCase();
      const ch = d.checks || {};
      const n = Object.keys(ch).length, okN = Object.keys(ch).filter(k => ch[k].ok).length;
      sh.appendRow([id, d.ts || new Date().toISOString(), String(d.name).trim(), d.variant, d.lab, d.title || '',
        n ? okN + '/' + n : '', '', '', JSON.stringify(ch), JSON.stringify(d.journals || []), d.ua || '', d.course || 'lazer']);
      return out_({ ok: true, id: id });
    }
    if (d.pw !== PASSWORD) return out_({ ok: false, error: 'Parol noto‘g‘ri.' });
    if (d.action === 'list') {
      const v = sh.getDataRange().getValues(); v.shift();
      const rows = v.filter(r => r[0]).map(r => ({ id: r[0], ts: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), name: r[2], variant: r[3], lab: r[4], title: r[5],
        grade: String(r[7] || ''), comment: String(r[8] || ''), checks: safe_(r[9], {}), journals: safe_(r[10], []), course: String(r[12] || 'lazer') }));
      return out_({ ok: true, rows: rows });
    }
    if (d.action === 'grade') {
      const v = sh.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(d.id)) {
        sh.getRange(i + 1, 8).setValue(d.grade || ''); sh.getRange(i + 1, 9).setValue(d.comment || '');
        return out_({ ok: true });
      }
      return out_({ ok: false, error: 'Yozuv topilmadi.' });
    }
    return out_({ ok: false, error: 'Noma’lum amal.' });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
function safe_(s, d) { try { return JSON.parse(s) } catch (e) { return d } }
