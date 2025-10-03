/******************************************************
 * Setup: Create tabs and seed 10 standard questions
 * Run seedStandardSurvey() once (then customize Questions).
 ******************************************************/
function seedStandardSurvey() {
  var ss = SpreadsheetApp.getActive();

  var qSh = ensureSheetWithHeaders_(ss, 'Questions',
    ['code','section','order','type','text','required','options','scale_min','scale_max']);
  var rSh = ensureSheetWithHeaders_(ss, 'Responses', ['timestamp','submission_id']);

  // Clear Questions below the header
  clearBelowHeader_(qSh);

  // 10 simple, mixed-type questions (feel free to edit later)
  var rows = [
    ['Q1','About you',1,'single','What is your age range?',true,'Under 18|18–24|25–34|35–44|45–54|55–64|65+','', ''],
    ['Q2','About you',2,'single','What is your primary role?',true,'Student|Professional|Manager|Executive|Other','', ''],
    ['Q3','About you',3,'multi','Which platforms do you use regularly? (Select all)',false,'Web|iOS|Android|Desktop','', ''],

    ['Q4','Experience',1,'likert','Rate your overall satisfaction',true,'',1,5],
    ['Q5','Experience',2,'likert','How easy was it to complete tasks?',false,'',1,5],

    ['Q6','Feedback',1,'short','One thing we did well was…',false,'','', ''],
    ['Q7','Feedback',2,'long','If we could improve one thing, it would be…',false,'','', ''],

    ['Q8','Preferences',1,'single','Would you recommend us to a friend?',true,'Yes|No|Not sure','', ''],
    ['Q9','Preferences',2,'multi','What features interest you most? (Select all)',false,'Speed|Analytics|Integrations|Mobile access|Security','', ''],
    ['Q10','Contact',1,'short','(Optional) Your email if you want follow-up',false,'','', '']
  ];

  if (rows.length) qSh.getRange(2,1,rows.length,rows[0].length).setValues(rows);

  // Build Responses header: timestamp | submission_id | codes…
  buildResponsesHeaderFromQuestions_();
  SpreadsheetApp.getUi().alert('Standard survey seeded. You can customize Questions now.');
}

/************ helpers (Setup) ************/
function ensureSheetWithHeaders_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function clearBelowHeader_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr > 1) sh.getRange(2,1,lr-1,lc).clearContent();
}
function buildResponsesHeaderFromQuestions_() {
  var ss = SpreadsheetApp.getActive();
  var qSh = ss.getSheetByName('Questions');
  var rSh = ss.getSheetByName('Responses');

  var q = qSh.getDataRange().getValues();
  if (q.length < 2) return;

  // Collect codes, sorted by section->order
  var body = q.slice(1).map(function(r){
    return { code:r[0], section:r[1], order:Number(r[2]||0) };
  }).filter(function(x){ return x.code; });

  body.sort(function(a,b){
    return a.section === b.section ? a.order - b.order : String(a.section).localeCompare(String(b.section));
  });

  var headers = ['timestamp','submission_id'];
  for (var i=0;i<body.length;i++) headers.push(body[i].code);

  rSh.clearContents();
  rSh.getRange(1,1,1,headers.length).setValues([headers]);
  rSh.setFrozenRows(1);
}
