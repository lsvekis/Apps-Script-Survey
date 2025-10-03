/******************************************************
 * Survey Web App (wide format: one submission = one row)
 ******************************************************/
var SPREADSHEET_ID = '1GcIzA'; // set if this project is NOT bound to the Sheet

function getSS_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Spreadsheet not found. Set SPREADSHEET_ID.');
  return ss;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Survey')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Return questions as a JSON string (avoids structured-clone issues) */
function loadQuestions() {
  try {
    var ss = getSS_();
    var qSh = ss.getSheetByName('Questions');
    if (!qSh) throw new Error('Questions sheet not found.');

    var lastRow = qSh.getLastRow(), lastCol = qSh.getLastColumn();
    if (lastRow < 2) return JSON.stringify({ ok:false, error:'No questions in sheet.' });

    var headers = qSh.getRange(1,1,1,lastCol).getValues()[0];
    var ix = {}; for (var i=0;i<headers.length;i++) ix[headers[i]] = i;
    var data = qSh.getRange(2,1,lastRow-1,lastCol).getValues();

    var items = [];
    data.forEach(function(row){
      var code = String(row[ix.code]||'').trim();
      if (!code) return;
      items.push({
        code: code,
        section: String(row[ix.section]||'').trim(),
        order: Number(row[ix.order]||0),
        type: String(row[ix.type]||'').toLowerCase(),
        text: String(row[ix.text]||'').trim(),
        required: String(row[ix.required]||'').toLowerCase()==='true',
        options: String(row[ix.options]||'').split('|').map(function(s){return s.trim();}).filter(Boolean),
        scale_min: Number(row[ix.scale_min]||0),
        scale_max: Number(row[ix.scale_max]||0)
      });
    });

    items.sort(function(a,b){
      return a.section===b.section ? a.order - b.order : String(a.section).localeCompare(String(b.section));
    });

    // Ensure Responses has at least timestamp/submission_id
    withRetry_(function(){
      var rSh = ss.getSheetByName('Responses') || ss.insertSheet('Responses');
      if (rSh.getLastColumn() < 2) {
        rSh.clearContents();
        rSh.getRange(1,1,1,2).setValues([['timestamp','submission_id']]);
        rSh.setFrozenRows(1);
      }
    });

    return JSON.stringify({ ok:true, questions:items });
  } catch (e) {
    return JSON.stringify({ ok:false, error: e.message || String(e) });
  }
}

/** Append one row; add columns if new codes appear; lock + retry for safety */
function submitRow(payload) {
  try {
    if (!payload || !payload.answers) return JSON.stringify({ ok:false, error:'Invalid payload' });

    var ss = getSS_();
    var rSh = ss.getSheetByName('Responses');
    if (!rSh) throw new Error('Responses sheet not found.');

    var lastCol = Math.max(2, rSh.getLastColumn());
    var headers = rSh.getRange(1,1,1,lastCol).getValues()[0];
    var map = {}; for (var i=0;i<headers.length;i++) map[headers[i]] = i;

    var row = new Array(headers.length);
    row[map.timestamp] = new Date();
    var sid = Utilities.getUuid();
    row[map.submission_id] = sid;

    var answers = payload.answers;
    Object.keys(answers).forEach(function(code){
      if (!map.hasOwnProperty(code)) {
        // Add new column if Questions changed after Responses header built
        var newCol = rSh.getLastColumn()+1;
        rSh.getRange(1,newCol).setValue(code);
        headers.push(code);
        map[code] = headers.length-1;
        row.push('');
      }
      var v = answers[code];
      if (Array.isArray(v)) v = v.join('; ');
      row[map[code]] = v;
    });

    withRetry_(function(){
      var lock = LockService.getDocumentLock();
      lock.waitLock(8000);
      try {
        var lr = rSh.getLastRow();
        rSh.getRange(lr+1,1,1,headers.length).setValues([row]);
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
    });

    return JSON.stringify({ ok:true, submissionId: sid });
  } catch (e) {
    return JSON.stringify({ ok:false, error: e.message || String(e) });
  }
}

/** Retry wrapper for transient storage/concurrency hiccups */
function withRetry_(fn) {
  var attempts = 5, lastErr;
  for (var i=0;i<attempts;i++) {
    try { return fn(); }
    catch(e) {
      var msg = String(e.message||e);
      if (msg.indexOf('FAILED_PRECONDITION')>-1 || msg.indexOf('Internal error')>-1 || msg.indexOf('Rate Limit')>-1) {
        lastErr = e;
        Utilities.sleep(200*(i+1)*(i+1)); // 200ms, 800ms, 1800ms, ...
        continue;
      }
      throw e;
    }
  }
  throw lastErr || new Error('Unknown storage error after retries.');
}
