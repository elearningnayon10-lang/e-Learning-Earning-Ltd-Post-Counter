// Team passcode: use English letters/numbers only, e.g. ELEL2026team
const PASS = 'CHANGE_ME';

const H = { posts: ['t', 'ch', 'd', 'c', 'note'], courses: ['id', 'g', 'n'] };

function sh_(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, 1, H[n].length).setNumberFormat('@').setValues([H[n]]);
  }
  return s;
}
function put_(n, list) {
  if (!list.length) return;
  const s = sh_(n), k = H[n].length;
  const vals = list.map(o => H[n].map(c => String(o[c] == null ? '' : o[c])));
  s.getRange(s.getLastRow() + 1, 1, vals.length, k).setNumberFormat('@').setValues(vals);
}
function get_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {}; h.forEach((c, i) => o[c] = String(r[i])); return o;
  });
}
function del_(n, id) {
  const s = sh_(n), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) s.deleteRow(i + 1);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.p !== PASS) return out_({ error: 'pass' });
  return out_({ posts: get_('posts'), cats: get_('courses') });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.p !== PASS) return out_({ error: 'pass' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.op === 'addPost') put_('posts', [b.r]);
    else if (b.op === 'delPost') del_('posts', b.t);
    else if (b.op === 'addCat') put_('courses', [b.r]);
    else if (b.op === 'delCat') del_('courses', b.id);
    else if (b.op === 'renCat') {
      const s = sh_('courses'), v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(b.id)) s.getRange(i + 1, 3).setValue(b.n);
    } else if (b.op === 'bulk') {
      if (get_('courses').length === 0 && get_('posts').length === 0) {
        put_('courses', b.cats);
        put_('posts', b.posts);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: 1 });
}
