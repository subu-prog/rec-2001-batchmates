var SHEET_ID = '1PbhzM_Ddo7aYfXIt_yhs32AdrHE6GnodNgBxX-72wU0';
var TAB = 'Entries';
var MAX_COMMENT = 5000;


function keyOk_(e) {
  var p = (e && e.parameter) || {};
  var want = PropertiesService.getScriptProperties().getProperty('PAGE_KEY');
  if (!want) return false;
  return !!p.key && p.key === want;
}

function out_(obj, e) {
  var json = JSON.stringify(obj);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
}

function readEntries_() {
  var sh = getSheet_();
  var lastRow = sh.getLastRow();
  var entries = [];
  if (lastRow >= 2) {
    var nCols = Math.max(sh.getLastColumn(), 5);
    var values = sh.getRange(2, 1, lastRow - 1, nCols).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[2] && !r[3]) continue;
      entries.push({ id: String(r[4] || ''), name: String(r[0]), roll: String(r[1]), branch: String(r[2]), comments: String(r[3]) });
    }
  }
  return entries;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var tok = String(p.tok || '');
  var ok = !!tok && !!CacheService.getScriptCache().get('tok_' + tok);
  if (!ok) return out_({ error: 'unauthorized' }, e);
  return out_({ entries: readEntries_() }, e);
}

function doPost(e) {
  if (!keyOk_(e)) return jsonOut_({ ok: false, error: 'unauthorized' });
  var p = (e && e.parameter) || {};
  var action = p.action || 'add';
  if (action === 'auth') {
    var tok = String(p.tok || '').slice(0, 100);
    if (tok) CacheService.getScriptCache().put('tok_' + tok, '1', 1800);
    return jsonOut_({ ok: true });
  }
  var name = String(p.name || '').slice(0, 200);
  var roll = String(p.roll || '').slice(0, 50);
  var branch = String(p.branch || '').slice(0, 100);
  var comments = String(p.comments || '').slice(0, MAX_COMMENT);
  var sh = getSheet_();
  if (action === 'edit') {
    var id = String(p.id || '');
    if (!id) return jsonOut_({ ok: false, error: 'missing id' });
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return jsonOut_({ ok: false, error: 'not found' });
    var ids = sh.getRange(2, 5, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        sh.getRange(i + 2, 1).setValue(name);
        sh.getRange(i + 2, 2).setNumberFormat('@').setValue(roll);
        sh.getRange(i + 2, 3).setValue(branch);
        sh.getRange(i + 2, 4).setValue(comments);
        return jsonOut_({ ok: true, id: id });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }
  if (!name && !roll && !branch && !comments) return jsonOut_({ ok: false, error: 'empty row' });
  var newId = Utilities.getUuid();
  sh.appendRow([name, '', branch, comments, newId]);
  var row = sh.getLastRow();
  sh.getRange(row, 2).setNumberFormat('@').setValue(roll);
  return jsonOut_({ ok: true, id: newId });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
