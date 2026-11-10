// Bean Buzz Sync v3 — App <-> Google Sheet bridge.
// POST (app se): Sales, Kharcha, Stock tabs app ke data se refresh hote hain.
// GET ?data=1   : Sheet se pura app-data JSON wapas — sheet master hai,
//                 app khulte hi latest data khinch leta hai (auto update).
// _meta tab     : updatedAt, partner on/off config, settlements (chhupa hua tab).
var TZ = 'Asia/Kolkata';

function convDate(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) {
    var p = v.split('-');
    return new Date(+p[0], (+p[1]) - 1, +p[2]);
  }
  return v;
}
function fmtD(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return '' + v;
}
function slug(s) { return ('' + s).toLowerCase().replace(/[^a-z0-9]+/g, ''); }
function metaSheet(ss) {
  var sh = ss.getSheetByName('_meta');
  if (!sh) { sh = ss.insertSheet('_meta'); sh.hideSheet(); }
  return sh;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Sales tab ---
    var sSh = ss.getSheetByName('Sales');
    if (sSh.getLastRow() > 1) sSh.getRange(2, 1, sSh.getLastRow() - 1, 5).clearContent();
    if (data.sales && data.sales.length) {
      var sRows = data.sales.map(function (s, i) {
        return [convDate(s.date), ('' + s.date).substring(0, 7), s.amount, s.note || '', '=SUM($C$2:C' + (i + 2) + ')'];
      });
      sSh.getRange(2, 1, sRows.length, 5).setValues(sRows);
    }

    // --- Kharcha tab ---
    var kSh = ss.getSheetByName('Kharcha');
    if (kSh.getLastRow() > 1) kSh.getRange(2, 1, kSh.getLastRow() - 1, 7).clearContent();
    if (data.costs && data.costs.length) {
      var kRows = data.costs.map(function (c) {
        return [convDate(c.date), ('' + c.date).substring(0, 7), c.kindLabel, c.label, c.qty === '' ? '' : c.qty, c.amount, c.note || ''];
      });
      kSh.getRange(2, 1, kRows.length, 7).setValues(kRows);
    }

    // --- Stock tab (item master) ---
    if (data.items && data.items.length) {
      var stSh = ss.getSheetByName('Stock');
      var existing = {};
      var last = stSh.getLastRow();
      if (last > 1) {
        var vals = stSh.getRange(2, 1, last - 1, 8).getValues();
        for (var i = 0; i < vals.length; i++) existing[vals[i][0]] = vals[i];
      }
      var rows = data.items.map(function (it) {
        var ex = existing[it.name];
        var closing = (data.closings && data.closings[it.id] !== undefined) ? data.closings[it.id] : (ex ? ex[7] : '');
        return [it.name, it.unit, it.rate, it.open, '', '', '', closing];
      });
      if (last > 1) stSh.getRange(2, 1, last - 1, 8).clearContent();
      stSh.getRange(2, 1, rows.length, 8).setValues(rows);
    }

    // --- _meta: updatedAt + settlements ---
    var m = metaSheet(ss);
    m.getRange('A1').setValue('updatedAt');
    if (data.meta && data.meta.updatedAt) m.getRange('A2').setValue(data.meta.updatedAt);
    if (data.settlements) m.getRange('C2').setValue(JSON.stringify(data.settlements));
    if (m.getRange('B1').getValue() === '') { m.getRange('B1').setValue('partner'); m.getRange('B2').setValue(0); }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, sales: (data.sales || []).length, costs: (data.costs || []).length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '' + err })).setMimeType(ContentService.MimeType.JSON);
  }
}

function buildData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Items from Stock tab
  var stSh = ss.getSheetByName('Stock');
  var items = [], byName = {}, closingById = {};
  if (stSh && stSh.getLastRow() > 1) {
    var vals = stSh.getRange(2, 1, stSh.getLastRow() - 1, 8).getValues();
    for (var i = 0; i < vals.length; i++) {
      var nm = '' + vals[i][0]; if (!nm) continue;
      var id = slug(nm);
      items.push({ id: id, name: nm, unit: '' + vals[i][1], rate: +vals[i][2] || 0, open: +vals[i][3] || 0, low: 0 });
      byName[nm] = id;
      if (vals[i][7] !== '' && vals[i][7] !== null) closingById[id] = +vals[i][7] || 0;
    }
  }
  var dashMonth = '2026-11';
  var dSh = ss.getSheetByName('Dashboard');
  if (dSh) { var mv = dSh.getRange('B4').getValue(); if (mv) dashMonth = ('' + mv).substring(0, 7); }
  var closings = {};
  var ck = Object.keys(closingById);
  if (ck.length) { closings[dashMonth] = closingById; }

  // Sales
  var sales = [];
  var sSh = ss.getSheetByName('Sales');
  if (sSh && sSh.getLastRow() > 1) {
    var sv = sSh.getRange(2, 1, sSh.getLastRow() - 1, 5).getValues();
    for (var j = 0; j < sv.length; j++) {
      if (!sv[j][0]) continue;
      sales.push({ id: 's' + j, date: fmtD(sv[j][0]), amount: +sv[j][2] || 0, note: '' + (sv[j][3] || '') });
    }
  }

  // Kharcha -> costs (item name se id match)
  var kindMap = { 'Stock Purchase': 'stock', 'Stock': 'stock', 'Fresh (roz ka)': 'fresh', 'Fresh': 'fresh', 'Operating': 'op', 'Wastage': 'waste' };
  var costs = [];
  var kSh = ss.getSheetByName('Kharcha');
  if (kSh && kSh.getLastRow() > 1) {
    var kv = kSh.getRange(2, 1, kSh.getLastRow() - 1, 7).getValues();
    for (var k = 0; k < kv.length; k++) {
      if (!kv[k][0]) continue;
      var label = '' + kv[k][3];
      var kind = kindMap['' + kv[k][2]] || 'op';
      var itemId = byName[label] || null;
      costs.push({
        id: 'c' + k, date: fmtD(kv[k][0]), kind: kind,
        itemId: (kind === 'stock' || kind === 'waste') ? itemId : null,
        name: (kind === 'stock' || kind === 'waste') ? '' : label,
        cat: kind === 'op' ? label : '',
        qty: kv[k][4] === '' ? '' : +kv[k][4] || 0,
        amount: +kv[k][5] || 0, note: '' + (kv[k][6] || '')
      });
    }
  }

  var m = metaSheet(ss);
  var upd = m.getRange('A2').getValue();
  var partner = m.getRange('B2').getValue();
  var sett = [];
  var sj = m.getRange('C2').getValue();
  if (sj) { try { sett = JSON.parse(sj); } catch (e) { sett = []; } }

  return {
    ok: true,
    db: { items: items, sales: sales, costs: costs, closings: closings, settlements: sett },
    meta: { updatedAt: +upd || 0 },
    config: { partner: ('' + partner) === '1' }
  };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.data) {
    return ContentService.createTextOutput(JSON.stringify(buildData())).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Bean Buzz Sync LIVE');
}
