/**
 * Heat Register — Google Sheets backend (Google Apps Script).
 *
 * Bind this script to the Google Sheet that should hold your data
 * (Sheet → Extensions → Apps Script), run setup() once, then deploy as a Web app.
 * The web app is the only door into the sheet; every request must carry the
 * access token that setup() creates.
 */

const TABS = {
  Heats: ['HeatNo', 'Date', 'Furnace', 'Product', 'Duration_hr', 'Input_kg', 'Output_kg', 'Yield_pct',
          'Materials_Cost', 'Chemicals_Cost', 'Labour', 'Depreciation', 'Overhead', 'Total_Cost',
          'Cost_per_kg', 'Grade', 'Notes', 'ID'],
  HeatItems: ['HeatNo', 'Kind', 'Name', 'Qty', 'Rate', 'Cost'],
  StockReceipts: ['ID', 'Date', 'Material', 'Qty', 'Unit', 'Rate', 'Value'],
  Assets: ['ID', 'Name', 'Cost', 'LifeType', 'LifeValue', 'Salvage', 'Dep_Rate'],
  Meta: ['Key', 'Value']
};

/* ---------- pure conversion helpers (no Sheets calls) ---------- */

function num_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

function normDate_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

// App data object -> { tabName: [rows] } (rows exclude the header)
function dbToTables_(db) {
  var t = { Heats: [], HeatItems: [], StockReceipts: [], Assets: [], Meta: [] };

  (db.heats || []).forEach(function (h) {
    var mats = h.materials || [], chems = h.chemicals || [];
    var inputQty = 0, matCost = 0, chemCost = 0;
    mats.forEach(function (m) {
      inputQty += num_(m.qty);
      matCost += num_(m.qty) * num_(m.rate);
      t.HeatItems.push([h.heatNo, 'material', m.name, num_(m.qty), num_(m.rate), num_(m.qty) * num_(m.rate)]);
    });
    chems.forEach(function (c) {
      chemCost += num_(c.qty) * num_(c.rate);
      t.HeatItems.push([h.heatNo, 'chemical', c.name, num_(c.qty), num_(c.rate), num_(c.qty) * num_(c.rate)]);
    });
    var out = num_(h.outputQty);
    var total = matCost + chemCost + num_(h.labour) + num_(h.depreciation) + num_(h.overhead);
    t.Heats.push([
      h.heatNo, h.date || '', h.furnace || '', h.product || '',
      h.duration === '' || h.duration === undefined ? '' : num_(h.duration),
      inputQty,
      h.outputQty === '' || h.outputQty === undefined ? '' : out,
      inputQty > 0 ? Math.round(out / inputQty * 10000) / 100 : 0,
      matCost, chemCost, num_(h.labour), num_(h.depreciation), num_(h.overhead),
      total, out > 0 ? Math.round(total / out * 100) / 100 : 0,
      h.grade || '', h.notes || '', h.id
    ]);
  });

  (db.materialsIn || []).forEach(function (m) {
    t.StockReceipts.push([m.id, m.date || '', m.name, num_(m.qty), m.unit || 'kg', num_(m.rate), num_(m.qty) * num_(m.rate)]);
  });

  (db.assets || []).forEach(function (a) {
    var life = num_(a.lifeValue) || 1;
    t.Assets.push([a.id, a.name, num_(a.cost), a.lifeType || 'months', num_(a.lifeValue), num_(a.salvage),
                   Math.max(0, (num_(a.cost) - num_(a.salvage)) / life)]);
  });

  t.Meta.push(['nextHeat', db.nextHeat || 1]);
  return t;
}

// { tabName: [rows] } -> app data object
function tablesToDb_(t, tz) {
  var heats = [];
  var items = (t.HeatItems || []).filter(function (r) { return r[0] !== ''; });

  (t.Heats || []).filter(function (r) { return r[0] !== '' && r[17] !== ''; }).forEach(function (r) {
    var no = num_(r[0]);
    var pick = function (kind) {
      return items.filter(function (i) { return num_(i[0]) === no && i[1] === kind; })
        .map(function (i) { return { name: String(i[2]), qty: num_(i[3]), rate: num_(i[4]) }; });
    };
    heats.push({
      id: String(r[17]), heatNo: no, date: normDate_(r[1], tz), furnace: String(r[2]), product: String(r[3]),
      duration: r[4] === '' ? '' : num_(r[4]),
      materials: pick('material'), chemicals: pick('chemical'),
      labour: num_(r[10]), depreciation: num_(r[11]), overhead: num_(r[12]),
      outputQty: r[6] === '' ? '' : num_(r[6]),
      grade: String(r[15]), notes: String(r[16])
    });
  });

  var materialsIn = (t.StockReceipts || []).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: String(r[0]), date: normDate_(r[1], tz), name: String(r[2]), qty: num_(r[3]), unit: String(r[4] || 'kg'), rate: num_(r[5]) };
  });

  var assets = (t.Assets || []).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), cost: num_(r[2]), lifeType: String(r[3] || 'months'), lifeValue: num_(r[4]), salvage: num_(r[5]) };
  });

  var nextHeat = 1;
  (t.Meta || []).forEach(function (r) { if (r[0] === 'nextHeat') nextHeat = num_(r[1]) || 1; });
  heats.forEach(function (h) { if (h.heatNo >= nextHeat) nextHeat = h.heatNo + 1; });

  return { heats: heats, materialsIn: materialsIn, assets: assets, nextHeat: nextHeat, nextMat: 1, nextAsset: 1 };
}

/* ---------- Sheets I/O ---------- */

function ensureTabs_(ss) {
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    var header = TABS[name];
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // keep dates as plain text so Sheets does not convert them
    if (name === 'Heats') sh.getRange(2, 2, 2000, 1).setNumberFormat('@');
    if (name === 'StockReceipts') sh.getRange(2, 2, 2000, 1).setNumberFormat('@');
  });
}

function readTables_(ss) {
  var out = {};
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    var width = TABS[name].length;
    var last = sh ? sh.getLastRow() : 1;
    out[name] = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  });
  return out;
}

function writeTables_(ss, t) {
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    var width = TABS[name].length;
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
    var rows = t[name] || [];
    if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
  });
}

/* ---------- web app endpoints ---------- */

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return respond_({ ok: true, message: 'Heat Register backend is running. Use the app to talk to it.' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token) return respond_({ ok: false, error: 'Backend not set up yet: run setup() once in Apps Script.' });
    if (body.token !== token) return respond_({ ok: false, error: 'Wrong access token.' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();

    if (body.action === 'ping') return respond_({ ok: true, sheetUrl: ss.getUrl() });

    if (body.action === 'load') {
      return respond_({ ok: true, sheetUrl: ss.getUrl(), db: tablesToDb_(readTables_(ss), tz) });
    }

    if (body.action === 'save') {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        ensureTabs_(ss);
        writeTables_(ss, dbToTables_(body.db || {}));
      } finally {
        lock.releaseLock();
      }
      return respond_({ ok: true, savedAt: new Date().toISOString() });
    }

    return respond_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return respond_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * Run this ONCE from the Apps Script editor (select setup → Run).
 * Creates the tabs and an access token, and writes the token to the "Setup" tab.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTabs_(ss);
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('TOKEN', token);
  }
  var sh = ss.getSheetByName('Setup') || ss.insertSheet('Setup');
  sh.clear();
  sh.getRange(1, 1, 4, 2).setValues([
    ['Heat Register backend', 'Ready'],
    ['Access token (paste into the app → Settings)', token],
    ['Next step', 'Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → copy the Web app URL'],
    ['Warning', 'Anyone with the token can read and overwrite this data. Do not share it. To change it, delete the TOKEN script property and run setup() again.']
  ]);
  sh.setColumnWidth(1, 330);
  sh.setColumnWidth(2, 640);
  Logger.log('Setup complete. Token is in the Setup tab.');
}
