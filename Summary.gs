/***************************************************************
 * Summary & Charts (optional but recommended)
 ***************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Survey')
    .addItem('Build Summary & Charts', 'buildSurveySummary')
    .addToUi();
}

function buildSurveySummary() {
  var ss = SpreadsheetApp.getActive();
  var qSh = ss.getSheetByName('Questions');
  var rSh = ss.getSheetByName('Responses');
  if (!qSh || !rSh) { SpreadsheetApp.getUi().alert('Questions or Responses sheet not found.'); return; }

  var sSh = ensureSheet_(ss, 'Summary'); clearSheet_(sSh);
  var vSh = ensureSheet_(ss, 'Verbatims'); clearSheet_(vSh);

  // Questions
  var qlr = qSh.getLastRow(), qlc = qSh.getLastColumn();
  var qHeaders = qSh.getRange(1,1,1,qlc).getValues()[0];
  var qIdx = index_(qHeaders);
  var qRows = qSh.getRange(2,1,Math.max(0, qlr-1),qlc).getValues();

  // Responses
  var rlr = rSh.getLastRow(), rlc = rSh.getLastColumn();
  if (rlr < 2) { sSh.getRange(1,1).setValue('No responses yet.'); return; }
  var rHeaders = rSh.getRange(1,1,1,rlc).getValues()[0];
  var rIdx = index_(rHeaders);
  var rRows = rSh.getRange(2,1,Math.max(0, rlr-1),rlc).getValues();

  var questions = [], byCode = {};
  qRows.forEach(function(r){
    var code = String(r[qIdx.code]||'').trim(); if (!code) return;
    var meta = {
      code: code,
      section: String(r[qIdx.section]||'').trim(),
      order: Number(r[qIdx.order]||0),
      type: String(r[qIdx.type]||'').toLowerCase().trim(),
      text: String(r[qIdx.text]||'').trim(),
      options: parseOptions_(String(r[qIdx.options]||'')),
      min: Number(r[qIdx.scale_min]||0),
      max: Number(r[qIdx.scale_max]||0)
    };
    questions.push(meta); byCode[code]=meta;
  });
  questions.sort(function(a,b){ return a.section===b.section ? a.order - b.order : String(a.section).localeCompare(String(b.section)); });

  // Header
  sSh.getRange(1,1,3,2).setValues([
    ['Survey Summary',''],
    ['Generated at', new Date()],
    ['Total submissions', rRows.length]
  ]);
  sSh.getRange(1,1,1,2).setFontWeight('bold').setFontSize(14);
  sSh.setFrozenRows(4);
  sSh.setColumnWidth(1, 380);
  sSh.setColumnWidth(2, 140);
  sSh.setColumnWidth(3, 140);
  for (var c=4;c<=12;c++) sSh.setColumnWidth(c, 110);

  // Verbatims header
  vSh.getRange(1,1,1,5).setValues([['submission_id','timestamp','code','question','answer']]).setFontWeight('bold');
  vSh.setFrozenRows(1);

  var row = 5, currentSection = null;

  questions.forEach(function(q){
    if (q.section !== currentSection) {
      sSh.getRange(row,1).setValue(q.section).setFontWeight('bold').setFontSize(12);
      sSh.getRange(row,1,1,12).setBackground('#f3f4f6'); row++;
      currentSection = q.section;
    }

    sSh.getRange(row,1).setValue(q.text + '  [' + q.code + ']').setFontWeight('bold'); row++;

    if (!rIdx.hasOwnProperty(q.code)) { sSh.getRange(row,1).setValue('No column in Responses for '+q.code).setFontColor('#b91c1c'); row+=2; return; }
    var col = rIdx[q.code];

    var values = rRows.map(function(r){ return r[col]; }).filter(function(v){ return v!=='' && v!=null; });

    if (q.type==='single' || q.type==='multi') {
      var tally = {}; q.options.forEach(function(opt){ tally[opt]=0; });
      var others = {};
      values.forEach(function(v){
        var parts = (q.type==='multi') ? String(v).split(';').map(function(s){return s.trim();}).filter(Boolean) : [String(v).trim()];
        parts.forEach(function(choice){
          if (tally.hasOwnProperty(choice)) tally[choice] += 1;
          else { if (!others[choice]) others[choice]=0; others[choice]+=1; }
        });
      });
      var table = [['Option','Count']];
      q.options.forEach(function(opt){ table.push([opt, tally[opt]||0]); });
      Object.keys(others).forEach(function(k){ table.push(['(Other) '+k, others[k]]); });

      var h = table.length;
      sSh.getRange(row,1,h,2).setValues(table);
      sSh.getRange(row,1,1,2).setFontWeight('bold').setBackground('#f8fafc');

      var chart = sSh.newChart().asBarChart()
        .setPosition(row,4,0,0)
        .addRange(sSh.getRange(row,1,h,2))
        .setOption('title', q.text)
        .setOption('legend',{position:'none'})
        .setOption('height', Math.max(220, 28*h))
        .build();
      sSh.insertChart(chart);
      row += h + 2;

    } else if (q.type==='likert') {
      var min=q.min||0, max=q.max||10, counts={}, total=0, sum=0;
      for (var s=min;s<=max;s++) counts[s]=0;
      values.forEach(function(v){
        var n=Number(v);
        if (!isNaN(n) && n>=min && n<=max){ counts[n]++; total++; sum+=n; }
      });
      var avg = total ? sum/total : 0;
      var dist = [['Score','Count']]; for (var s2=min;s2<=max;s2++) dist.push([s2, counts[s2]||0]);

      sSh.getRange(row,1,dist.length,2).setValues(dist);
      sSh.getRange(row,1,1,2).setFontWeight('bold').setBackground('#f8fafc');

      var lchart = sSh.newChart().asColumnChart()
        .setPosition(row,4,0,0)
        .addRange(sSh.getRange(row,1,dist.length,2))
        .setOption('title', q.text + ' (Avg: ' + avg.toFixed(2) + ')')
        .setOption('legend',{position:'none'})
        .setOption('height', 260)
        .build();
      sSh.insertChart(lchart);

      var stats = [['Metric','Value'], ['Responses', total], ['Average', avg]];
      sSh.getRange(row,3,stats.length,2).setValues(stats);
      sSh.getRange(row,3,1,2).setFontWeight('bold').setBackground('#f8fafc');

      row += Math.max(dist.length, stats.length) + 2;

    } else if (q.type==='short' || q.type==='long') {
      var count = 0, vOut=[];
      for (var ri=0; ri<rRows.length; ri++){
        var ans = rRows[ri][col];
        if (ans!=='' && ans!=null){
          count++;
          vOut.push([rRows[ri][rIdx.submission_id]||'', rRows[ri][rIdx.timestamp]||'', q.code, q.text, ans]);
        }
      }
      if (vOut.length) vSh.getRange(vSh.getLastRow()+1,1,vOut.length,5).setValues(vOut);
      var t = [['Responses (non-empty)','Count'], ['Total', count]];
      sSh.getRange(row,1,t.length,2).setValues(t);
      sSh.getRange(row,1,1,2).setFontWeight('bold').setBackground('#f8fafc');
      row += t.length + 2;

    } else {
      sSh.getRange(row,1).setValue('Unsupported type: '+q.type); row += 2;
    }
  });

  sSh.autoResizeColumns(1,3); vSh.autoResizeColumns(1,5);
  SpreadsheetApp.getUi().alert('Summary built. Open the "Summary" and "Verbatims" sheets.');
}

/*** small helpers ***/
function ensureSheet_(ss, name){ return ss.getSheetByName(name) || ss.insertSheet(name); }
function clearSheet_(sh){ sh.getCharts().forEach(function(c){ sh.removeChart(c); }); sh.clear(); }
function index_(headers){ var o={}; headers.forEach(function(h,i){ o[String(h)]=i; }); return o; }
function parseOptions_(s){ return String(s||'').split('|').map(function(x){return x.trim();}).filter(Boolean); }
