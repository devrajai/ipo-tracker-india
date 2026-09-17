/**
 * ============================================================================
 *  INDIA IPO TRACKER - Google Apps Script
 * ============================================================================
 *  WHAT IT DOES
 *    - Fetches current Mainboard + SME IPO data from Chittorgarh (free, no API
 *      key needed).
 *    - Sorts every IPO into the right tab: Upcoming / Open / Closed / Listed.
 *    - When an IPO's subscription closes, it moves automatically to "Closed".
 *    - When its listing window passes, it moves automatically to "Listed".
 *    - For Open IPOs it tries to fetch price band, lot size, issue size,
 *      subscription (RII/QIB/NII) and GMP from each IPO's detail page.
 *    - Every run appends a daily GMP snapshot to the "GMP Log" tab.
 *
 *  HOW TO INSTALL (one time, ~5 minutes)
 *    1. Open the "IPO Tracker - India" spreadsheet.
 *    2. Extensions > Apps Script. Delete any code, paste this ENTIRE file.
 *    3. Click the Save icon.
 *    4. In the toolbar dropdown choose "updateAll" > click Run.
 *    5. Authorize: Review permissions > your account > Advanced > Allow.
 *    6. Run "installTrigger" once. Done - it now updates every day at ~8 AM.
 *    7. Optional: in Apps Script left sidebar > Project Settings, set the
 *       time zone to (GMT+05:30) Kolkata so the trigger fires in the morning.
 *
 *  IF A SITE CHANGES ITS LAYOUT: the script logs errors instead of crashing.
 *    Ask in the chat for an updated version and re-paste it.
 * ============================================================================ */

var CONFIG = {
  SHEET_OPEN: 'Open',
  SHEET_UPCOMING: 'Upcoming',
  SHEET_CLOSED: 'Closed',
  SHEET_LISTED: 'Listed',
  SHEET_GMP: 'GMP Log',
  SHEET_DASH: 'Dashboard',

  // Main sources (server-rendered HTML - verified 2026-09-17)
  SOURCES: [
    'https://www.chittorgarh.com/',                                // Homepage (mainboard + SME tables)
    'https://www.chittorgarh.com/ipo/ipo_dashboard.asp',          // Mainboard dashboard
    'https://www.chittorgarh.com/ipo/ipo_dashboard.asp?a=sme'     // SME dashboard
  ],
  // Performance-tracker pages: every LISTED IPO of the year (fills the Listed tab)
  LISTED_SOURCES: [
    'https://www.chittorgarh.com/ipo/ipo_perf_tracker.asp',
    'https://www.chittorgarh.com/ipo/ipo_perf_tracker.asp?exchange=sme'
  ],
  BASE_URL: 'https://www.chittorgarh.com',

  // Fetch per-IPO detail pages (price band, lot size, GMP, subscription)?
  FETCH_DETAILS: true,
  MAX_DETAIL_FETCHES: 30,

  // An IPO that closed more than this many days ago is assumed listed.
  LISTED_AFTER_DAYS: 14,

  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
              '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
};

// Column order MUST match row 1 of the Open / Upcoming / Closed / Listed tabs.
var COLS = ['name','sector','openDate','closeDate','listingDate','priceLow',
            'priceHigh','lotSize','lotValue','issueSize','rii','qib','nii',
            'overallSub','gmp','estGain','de','roe','revGrowth',
            'listingPrice','actualGain','source','updatedAt','detailUrl'];

var DIAG = []; // per-run diagnostics, written to Dashboard column D

function updateAll() {
  var started = new Date();
  var status = [];
  try {
    var ipos = collectIpos_();                       // scrape + classify
    status.push(ipos.length + ' IPOs found');
    syncSheets_(ipos);                               // upsert / move / archive
    backfillListed_();                              // move listed IPOs to Listed tab
    if (CONFIG.FETCH_DETAILS) fetchDetails_();       // enrich Open tab + GMP log
    status.push('sheets synced');
  } catch (err) {
    status.push('ERROR: ' + err);
  }
  writeRunStatus_(started, status.join(' | '));
}

/** Convenience alias. */
function runNow() { updateAll(); }

/** Creates the daily trigger. Run this once. */
function installTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('updateAll').timeBased()
           .everyDays(1).atHour(8).create();
}

/** Removes all triggers of this project. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
}

/* ============================ SCRAPING ================================== */

/**
 * Fetches the dashboard pages, extracts IPO name + date range + detail link,
 * classifies each IPO by today's date.
 * Returns [{name, open, close, href, board, cat}]
 */
function collectIpos_() {
  var out = [];
  var today = startOfDay_(new Date());
  var seen = {};
  DIAG = [];

  CONFIG.SOURCES.forEach(function (url) {
    var board = /a=sme/i.test(url) ? 'SME' : 'Mainboard';
    var html;
    try {
      html = fetchHtml_(url);
    } catch (e) {
      DIAG.push(board + ' [' + url + '] FETCH FAILED: ' + (e && e.message ? e.message : e));
      return;
    }
    var rows = parseAllRows_(html);
    var found = 0;

    rows.forEach(function (r) {
      try {
        // Find the cell whose link points to an /ipo/ detail page
        var nameCell = null;
        for (var i = 0; i < r.length; i++) {
          if (r[i].href && r[i].href.indexOf('/ipo/') !== -1 &&
              r[i].href.indexOf('ipo_dashboard') === -1) { nameCell = r[i]; break; }
        }
        if (!nameCell) return;

        // Find a date-range text anywhere in the row ("16 - 18 Sep")
        var dateText = '';
        for (var j = 0; j < r.length; j++) {
          if (r[j].text && /\d{1,2}\s*[-\u2013]\s*\d{1,2}\s+[A-Za-z]{3,9}/.test(r[j].text)) {
            dateText = r[j].text; break;
          }
        }
        if (!dateText) return;

        var dates = parseDateRange_(dateText);
        if (!dates) return;

        var rawName = nameCell.text
            .replace(/(\d{1,2}\s*[-\u2013]\s*\d{1,2}\s+[A-Za-z]{3,9}).*$/, ' ')  // date glued to name
            .replace(/&/gi, '&')                 // decode entity
            .replace(/\s+/g, ' ')                    // normalize whitespace
            .replace(/\s*\(?(IPO|FPO)\)?\s*$/i, '')
            .replace(/[\s\u00a0]+(O|P|CT|LT)\s*$/i, '')  // trailing status letters
            .trim();
        if (!rawName || rawName.length < 3 || /^no records/i.test(rawName)) return;

        var key = rawName.toLowerCase();
        if (seen[key]) return;                      // dedupe across boards

        var href = nameCell.href;
        if (href && href.indexOf('http') !== 0) href = CONFIG.BASE_URL + href;

        seen[key] = true;
        out.push({
          name: rawName,
          open: dates.open,
          close: dates.close,
          href: href,
          board: board,
          cat: classify_(dates, today)
        });
        found++;
      } catch (rowErr) {
        // one bad row must never kill the whole run
        log_('Skipped a bad row: ' + rowErr);
      }
    });

    DIAG.push(board + ' [' + url + '] OK: ' + html.length + ' chars, ' +
               rows.length + ' rows scanned, ' + found + ' IPO rows matched');
  });

  if (!out.length) throw new Error('No IPO rows found. ' + DIAG.join(' | '));
  return out;
}

/** open  / close / listed classification from dates. */
function classify_(dates, today) {
  if (today < dates.open) return 'Upcoming';
  if (today <= dates.close) return 'Open';
  var daysSince = (today - dates.close) / 86400000;
  return daysSince <= CONFIG.LISTED_AFTER_DAYS ? 'Closed' : 'Listed';
}

/** GET with full browser headers, cookie bootstrap and retry. */
function fetchHtml_(url) {
  var headers = {
    'User-Agent': CONFIG.USER_AGENT,
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-IN,en;q=0.9',
    'Referer': 'https://www.google.com/',
    'Cache-Control': 'no-cache',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'cross-site',
    'Sec-Fetch-User': '?1'
  };

  // Cookie bootstrap: visit the homepage first, carry its cookies forward
  try {
    var home = UrlFetchApp.fetch(CONFIG.BASE_URL + '/', {
      headers: { 'User-Agent': CONFIG.USER_AGENT,
                 'Accept': 'text/html,*/*;q=0.8',
                 'Accept-Language': 'en-IN,en;q=0.9' },
      followRedirects: true, muteHttpExceptions: true
    });
    var sc = home.getAllHeaders()['Set-Cookie'];
    if (sc) {
      var jar = [];
      (Array.isArray(sc) ? sc : [sc]).forEach(function (c) {
        jar.push(String(c).split(';')[0]);
      });
      headers['Cookie'] = jar.join('; ');
    }
  } catch (e) { /* proceed without cookies */ }

  var params = { headers: headers, followRedirects: true, muteHttpExceptions: true };
  var res;
  for (var i = 0; i < 2; i++) {
    res = UrlFetchApp.fetch(url, params);
    var body = res.getContentText();
    if (res.getResponseCode() === 200 && body && body.length > 500) return body;
    Utilities.sleep(2000);
  }
  throw new Error('HTTP ' + res.getResponseCode() + ', body ' + (body ? body.length : 0) + ' chars');
}

/* ------------------------- HTML parsing helpers ------------------------- */

/** Minimal regex-based <table> parser -> [{rows: [[{text, href}]]}] */
function parseHtmlTables_(html) {
  var tables = [];
  var tRe = /<table[^>]*>([\s\S]*?)<\/table>/gi, m;
  while ((m = tRe.exec(html)) !== null) {
    var rows = [];
    var rRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi, r;
    while ((r = rRe.exec(m[1])) !== null) {
      var cells = [];
      var cRe = /<(td|th)[^>]*>([\s\S]*?)<\/\1>/gi, c;
      while ((c = cRe.exec(r[1])) !== null) {
        var h = c[2].match(/<a[^>]+href=["']([^"']+)["']/i);
        cells.push({
          text: cleanText_(c[2]),
          href: h ? h[1] : null
        });
      }
      if (cells.length) rows.push(cells);
    }
    if (rows.length) tables.push({ rows: rows });
  }
  return tables;
}

/** Scans ALL <tr> rows in the document, ignoring table nesting. */
function parseAllRows_(html) {
  var rows = [];
  var rRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi, r;
  while ((r = rRe.exec(html)) !== null) {
    var cells = [];
    var cRe = /<(td|th)[^>]*>([\s\S]*?)<\/\1>/gi, c;
    while ((c = cRe.exec(r[1])) !== null) {
      var h = c[2].match(/<a[^>]+href=["']([^"']+)["']/i);
      cells.push({ text: cleanText_(c[2]), href: h ? h[1] : null });
    }
    if (cells.length) rows.push(cells);
  }
  return rows;
}

function cleanText_(s) {
  return s.replace(/<[^>]+>/g, ' ')
          .replace(/&nbsp;/gi, ' ')
          .replace(/&/gi, '<').replace(/&/gi, '>')
          .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(d); })
          .replace(/&/gi, '&')
          .replace(/\s+/g, ' ').trim();
}

/** "16 - 18 Sep" / "01-03 Sep" -> {open: Date, close: Date} (year inferred). */
function parseDateRange_(txt) {
  if (!txt) return null;
  var m = txt.match(/(\d{1,2})\s*[-\u2013]\s*(\d{1,2})\s+([A-Za-z]{3,9})/);
  var now = new Date();
  if (m) {
    var month = monthIndex_(m[3]);
    if (month === -1) return null;
    var year = now.getFullYear();
    var close = new Date(year, month, parseInt(m[2], 10));
    // If that close date is far in the past it belongs to next year.
    if (close.getTime() < now.getTime() - 45 * 86400000) {
      year += 1;
      close = new Date(year, month, parseInt(m[2], 10));
    }
    var open = new Date(year, month, parseInt(m[1], 10));
    return { open: open, close: close };
  }
  // Single date fallback: "18 Sep 2026"
  m = txt.match(/(\d{1,2})\s+([A-Za-z]{3,9})\s*(\d{4})?/);
  if (m) {
    var mo = monthIndex_(m[2]);
    if (mo === -1) return null;
    var yr = m[3] ? parseInt(m[3], 10) : now.getFullYear();
    var d = new Date(yr, mo, parseInt(m[1], 10));
    return { open: d, close: d };
  }
  return null;
}

function monthIndex_(s) {
  var months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var i = months.indexOf(s.slice(0, 3).toLowerCase());
  return i;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

/* ============================ SHEET SYNC ================================ */

function syncSheets_(ipos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = {
    Open: ss.getSheetByName(CONFIG.SHEET_OPEN),
    Upcoming: ss.getSheetByName(CONFIG.SHEET_UPCOMING),
    Closed: ss.getSheetByName(CONFIG.SHEET_CLOSED),
    Listed: ss.getSheetByName(CONFIG.SHEET_LISTED)
  };

  // Index existing rows by name -> {sheet, rowIndex, values}
  var index = {};
  Object.keys(sheets).forEach(function (cat) {
    var sh = sheets[cat];
    var last = sh.getLastRow();
    if (last < 2) return;
    var names = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      var n = String(names[i][0]).trim().toLowerCase();
      if (n) index[n] = { sheet: sh, cat: cat, rowIndex: i + 2 };
    }
  });

  var appendCount = {};
  Object.keys(sheets).forEach(function (c) { appendCount[c] = 0; });

  ipos.forEach(function (ipo) {
    var key = ipo.name.toLowerCase();
    var existing = index[key];

    // Fresh data known so far (dates + source + detail link)
    var data = {
      openDate: ipo.open, closeDate: ipo.close,
      source: 'Chittorgarh (' + ipo.board + ')',
      updatedAt: new Date(),
      detailUrl: ipo.href || ''
    };

    if (existing) {
      // Same category -> update dates/source in place.
      if (existing.cat === ipo.cat) {
        var sh = existing.sheet;
        var vals = sh.getRange(existing.rowIndex, 1, 1, COLS.length).getValues()[0];
        vals[COLS.indexOf('openDate')] = data.openDate;
        vals[COLS.indexOf('closeDate')] = data.closeDate;
        vals[COLS.indexOf('source')] = data.source;
        vals[COLS.indexOf('updatedAt')] = data.updatedAt;
        if (data.detailUrl && !vals[COLS.indexOf('detailUrl')]) {
          vals[COLS.indexOf('detailUrl')] = data.detailUrl;
        }
        sh.getRange(existing.rowIndex, 1, 1, COLS.length).setValues([vals]);
      } else {
        // Category changed (Open->Closed, Closed->Listed, Upcoming->Open...)
        moveRow_(existing.sheet, existing.rowIndex, sheets[ipo.cat], data, ipo.name);
      }
    } else {
      // Brand new IPO -> append to its category tab.
      var row = newRow_(ipo.name, data);
      var target = sheets[ipo.cat];
      target.appendRow(row);
      appendCount[ipo.cat] += 1;
    }
  });

  // If an "Open" IPO stopped appearing on the dashboard, close-date logic
  // inside next runs will retire it; nothing else needed here.
}

function moveRow_(fromSheet, rowIndex, toSheet, newData, name) {
  var vals = fromSheet.getRange(rowIndex, 1, 1, COLS.length).getValues()[0];
  vals[COLS.indexOf('openDate')] = newData.openDate;
  vals[COLS.indexOf('closeDate')] = newData.closeDate;
  vals[COLS.indexOf('source')] = newData.source;
  vals[COLS.indexOf('updatedAt')] = newData.updatedAt;
  toSheet.appendRow(vals);
  fromSheet.deleteRow(rowIndex);
  log_('Moved "' + name + '" -> ' + toSheet.getName());
}

function newRow_(name, data) {
  var row = [];
  for (var i = 0; i < COLS.length; i++) row.push('');
  row[COLS.indexOf('name')] = name;
  Object.keys(data).forEach(function (k) {
    var idx = COLS.indexOf(k);
    if (idx >= 0) row[idx] = data[k];
  });
  return row;
}

/* ===================== PER-IPO DETAIL ENRICHMENT ======================== */

/**
 * For every IPO currently in the Open tab, fetch its detail page and extract
 * price band, lot size, issue size, listing date, subscription and GMP.
 * Also appends a daily GMP snapshot to the GMP Log tab.
 */
function fetchDetails_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.SHEET_OPEN);
  var gmpSheet = ss.getSheetByName(CONFIG.SHEET_GMP);
  var last = sh.getLastRow();
  if (last < 2) return;
  var n = last - 1;
  if (n > CONFIG.MAX_DETAIL_FETCHES) n = CONFIG.MAX_DETAIL_FETCHES;
  var values = sh.getRange(2, 1, n, COLS.length).getValues();
  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (var i = 0; i < n; i++) {
    var row = values[i];
    var name = String(row[0]);
    // Use the real detail URL captured from the dashboard link; fall back to a guess
    var url = String(row[COLS.indexOf('detailUrl')] || '').trim();
    if (!url) url = detailUrl_(name);
    if (!url) continue;

    try {
      var html = fetchHtml_(url);
      applyDetail_(row, html);

      // Daily GMP snapshot (skip if already logged today for this IPO)
      var gmp = row[COLS.indexOf('gmp')];
      if (gmp !== '' && !gmpLoggedToday_(gmpSheet, todayKey, name)) {
        gmpSheet.appendRow([new Date(), name, gmp,
                            row[COLS.indexOf('priceHigh')], '', '']);
      }
    } catch (e) { /* best effort - leave row unchanged */ }

    // small politeness delay
    Utilities.sleep(400);
  }
  sh.getRange(2, 1, n, COLS.length).setValues(values);
}

/** Chittorgarh detail pages live at /ipo/<slug>/; slug derived from name. */
function detailUrl_(name) {
  return CONFIG.BASE_URL + '/ipo/' + name.toLowerCase()
      .replace(/[^a-z0-9]+/g, '-') + '-ipo/';
}

function applyDetail_(row, html) {
  var text = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');

  var m;
  m = text.match(/price\s*band[^0-9]{0,40}([\d,]+)\s*(?:to|[-\u2013])\s*([\d,]+)/i);
  if (m) {
    row[COLS.indexOf('priceLow')] = num_(m[1]);
    row[COLS.indexOf('priceHigh')] = num_(m[2]);
    var lotM = text.match(/lot\s*size[^0-9]{0,40}([\d,]+)/i);
    if (lotM) {
      var lot = num_(lotM[1]);
      row[COLS.indexOf('lotSize')] = lot;
      if (lot) row[COLS.indexOf('lotValue')] = lot * num_(m[2]);
    }
  }
  m = text.match(/issue\s*size[^0-9]{0,60}([\d.,]+)\s*(?:crore|cr\b)/i);
  if (m) row[COLS.indexOf('issueSize')] = parseFloat(m[1].replace(/,/g, ''));

  m = text.match(/listing\s*date[^A-Za-z0-9]{0,30}(\d{1,2}\s+\w+\s*,?\s*\d{4}|\w+\s+\d{1,2},?\s*\d{4})/i);
  if (m) row[COLS.indexOf('listingDate')] = new Date(m[1]);

  m = text.match(/GMP[^0-9+\-]{0,50}([+\-]?\d[\d.]*)/i);
  if (m) row[COLS.indexOf('gmp')] = parseFloat(m[1]);

  m = text.match(/(?:Retail|RII)[^0-9]{0,20}([\d.]+)\s*x/i);
  if (m) row[COLS.indexOf('rii')] = parseFloat(m[1]);
  m = text.match(/QIB[^0-9]{0,20}([\d.]+)\s*x/i);
  if (m) row[COLS.indexOf('qib')] = parseFloat(m[1]);
  m = text.match(/(?:NII|Non\s*Institutional)[^0-9]{0,20}([\d.]+)\s*x/i);
  if (m) row[COLS.indexOf('nii')] = parseFloat(m[1]);
  m = text.match(/(?:total|overall)[^0-9]{0,20}subscription[^0-9]{0,20}([\d.]+)\s*x/i);
  if (m) row[COLS.indexOf('overallSub')] = parseFloat(m[1]);

  // Estimated listing gain from GMP (if both available)
  var gmp = row[COLS.indexOf('gmp')], high = row[COLS.indexOf('priceHigh')];
  if (gmp !== '' && high) {
    row[COLS.indexOf('estGain')] = Math.round((gmp / high) * 1000) / 10;
  }
  row[COLS.indexOf('updatedAt')] = new Date();
}

function gmpLoggedToday_(gmpSheet, todayKey, name) {
  var last = gmpSheet.getLastRow();
  if (last < 2) return false;
  var tz = Session.getScriptTimeZone();
  var dates = gmpSheet.getRange(Math.max(2, last - 100), 1, Math.min(99, last - 1), 2).getValues();
  for (var i = dates.length - 1; i >= 0; i--) {
    var d = dates[i][0] instanceof Date ? Utilities.formatDate(dates[i][0], tz, 'yyyy-MM-dd') : '';
    if (d === todayKey && String(dates[i][1]).toLowerCase() === String(name).toLowerCase()) return true;
  }
  return false;
}

function num_(s) { return parseFloat(String(s).replace(/,/g, '')); }

/* ===================== LISTED-BACKFILL =================================== */

/**
 * Backfills the Listed tab from the performance-tracker pages (all listed
 * IPOs of the year). IPOs found in other tabs are moved to Listed.
 */
function backfillListed_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDefs = [
    ['Open', CONFIG.SHEET_OPEN], ['Upcoming', CONFIG.SHEET_UPCOMING],
    ['Closed', CONFIG.SHEET_CLOSED], ['Listed', CONFIG.SHEET_LISTED]
  ];
  var sheets = {}, nameIndex = {};
  sheetDefs.forEach(function (d) {
    var sh = ss.getSheetByName(d[1]);
    sheets[d[0]] = sh;
    var last = sh.getLastRow();
    if (last < 2) return;
    var names = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      var n = String(names[i][0]).trim().toLowerCase();
      if (n && !(n in nameIndex)) nameIndex[n] = d[0];
    }
  });

  var BLOCK = ['company','issuer','issuer company','company name','name','sr no','no.','sno'];

  CONFIG.LISTED_SOURCES.forEach(function (url) {
    var board = /exchange=sme/i.test(url) ? 'SME' : 'Mainboard';
    var html;
    try {
      html = fetchHtml_(url);
    } catch (e) {
      DIAG.push('Listed ' + board + ' [' + url + '] FETCH FAILED: ' + (e && e.message ? e.message : e));
      return;
    }
    var rows = parseAllRows_(html);
    var toMove = [], toAdd = [], samples = 0;

    rows.forEach(function (r) {
      // Diagnostics: capture a few sample rows so the structure is visible
      if (samples < 3 && r.length >= 4) {
        var txt = r.map(function (c) { return String(c.text).substring(0, 22); }).join(' | ');
        if (txt.replace(/[|\s]/g, '').length > 5) {
          DIAG.push('PT ' + board + ' sample: [' + txt + ']');
          samples++;
        }
      }

      try {
        if (r.length < 4) return;                       // full data rows have many columns

        // must contain a date (range "16 - 18 Sep" or single "01 Sep 2026")
        var hasDate = false;
        for (var d2 = 0; d2 < r.length; d2++) {
          var t2 = String(r[d2].text);
          if (/\d{1,2}\s*[-\u2013]\s*\d{1,2}\s+[A-Za-z]{3,9}/.test(t2) ||
              /\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}/.test(t2)) { hasDate = true; break; }
        }
        if (!hasDate) return;

        // must contain at least 2 numeric cells (prices/gains) - not nav/sidebar rows
        var numCells = 0;
        for (var d3 = 0; d3 < r.length; d3++) {
          if (/^[\u20b9$]?\s*[\d,]+(\.\d+)?\s*%?$/.test(String(r[d3].text).trim())) numCells++;
        }
        if (numCells < 2) return;

        // Name cell: prefer /ipo/ link, then any non-nav link, then first letter cell
        var nameCell = null, fallbackCell = null;
        for (var i = 0; i < r.length; i++) {
          var h = r[i].href;
          if (h && String(h).indexOf('/ipo/') !== -1 && String(h).indexOf('ipo_perf') === -1) {
            nameCell = r[i]; break;
          }
          if (!fallbackCell && h && !isNavHref_(h)) fallbackCell = r[i];
        }
        if (!nameCell) nameCell = fallbackCell;
        if (!nameCell) {
          for (var i3 = 0; i3 < r.length; i3++) {
            if (/[A-Za-z]{3}/.test(r[i3].text) && !/^\d/.test(String(r[i3].text).trim())) {
              nameCell = r[i3]; break;
            }
          }
        }
        if (!nameCell) return;

        var rawName = nameCell.text
            .replace(/(\d{1,2}\s*[-\u2013]\s*\d{1,2}\s+[A-Za-z]{3,9}).*$/, ' ')
            .replace(/&/gi, '&')
            .replace(/\s+/g, ' ')
            .replace(/\s*\(?(IPO|FPO)\)?\s*$/i, '')
            .replace(/[\s\u00a0]+(O|P|CT|LT)\s*$/i, '')
            .trim();
        if (!rawName || rawName.length < 3) return;
        if (BLOCK.indexOf(rawName.toLowerCase()) !== -1) return;

        var key = rawName.toLowerCase();
        var where = nameIndex[key];
        if (where === 'Listed') return;                 // already archived

        var href = nameCell.href;
        if (href && href.indexOf('http') !== 0) href = CONFIG.BASE_URL + href;
        var entry = { name: rawName, key: key, href: href, fromCat: where || null,
                      listingDate: listDateFromRow_(r) };

        if (where) { toMove.push(entry); } else { toAdd.push(entry); }
        nameIndex[key] = 'Listed';                       // avoid duplicates within run
      } catch (e) { /* skip bad rows */ }
    });

    // Execute moves
    toMove.forEach(function (mv) {
      try {
        if (!mv.fromCat) return;
        var sh = sheets[mv.fromCat];
        var last = sh.getLastRow();
        if (last < 2) return;
        var names = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < names.length; i++) {
          if (String(names[i][0]).trim().toLowerCase() === mv.key) {
            var vals = sh.getRange(i + 2, 1, 1, COLS.length).getValues()[0];
            sh.deleteRow(i + 2);
            if (!vals[COLS.indexOf('listingDate')]) vals[COLS.indexOf('listingDate')] = mv.listingDate;
            vals[COLS.indexOf('source')] = 'Chittorgarh (' + board + ') - listed';
            vals[COLS.indexOf('updatedAt')] = new Date();
            if (!vals[COLS.indexOf('detailUrl')] && mv.href) vals[COLS.indexOf('detailUrl')] = mv.href;
            sheets.Listed.appendRow(vals);
            return;
          }
        }
      } catch (e) { /* skip */ }
    });

    // Execute adds
    toAdd.forEach(function (a) {
      try {
        sheets.Listed.appendRow(newRow_(a.name, {
          source: 'Chittorgarh (' + board + ') - listed',
          listingDate: a.listingDate,
          updatedAt: new Date(),
          detailUrl: a.href || ''
        }));
      } catch (e) { /* skip */ }
    });

    DIAG.push('Listed ' + board + ' [' + url + '] OK: ' + rows.length +
               ' rows scanned, ' + toMove.length + ' moved, ' + toAdd.length + ' added');
  });
}

function isNavHref_(h) {
  var l = String(h || '').toLowerCase();
  if (!l) return true;
  return l.charAt(0) === '#' || l.indexOf('javascript:') === 0 || l.indexOf('mailto:') === 0 ||
         l.indexOf('ipo_perf') !== -1 || l.indexOf('ipo_dashboard') !== -1 ||
         l.indexOf('/broker') !== -1 || l.indexOf('open-account') !== -1 ||
         l.indexOf('.css') !== -1 || l.indexOf('.js') !== -1 ||
         l.indexOf('facebook') !== -1 || l.indexOf('twitter') !== -1 || l.indexOf('whatsapp') !== -1;
}


/** Finds a single listing date like "01 Sep 2026" anywhere in a row. */
function listDateFromRow_(r) {
  for (var i = 0; i < r.length; i++) {
    var m = String(r[i].text).match(/(\d{1,2})\s+([A-Za-z]{3,9}),?\s+(\d{4})/);
    if (m) {
      var mo = monthIndex_(m[2]);
      if (mo !== -1) return new Date(parseInt(m[3], 10), mo, parseInt(m[1], 10));
    }
  }
  return '';
}

/* ============================== LOGGING ================================= */

function log_(msg) {
  try { Logger.log(msg); } catch (e) {}
}

function writeRunStatus_(started, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName(CONFIG.SHEET_DASH);
    dash.getRange('A24').setValue('Last automated run');
    dash.getRange('B24').setValue(
      Utilities.formatDate(started, Session.getScriptTimeZone(),
                          'dd/MM/yyyy HH:mm') + ' - ' + status);
    // Diagnostics area (Dashboard column D)
    dash.getRange('D23').setValue('DIAGNOSTICS');
    dash.getRange('D24:D40').clearContent();
    if (DIAG && DIAG.length) {
      var d = DIAG.map(function (s) { return [String(s).substring(0, 200)]; });
      dash.getRange(24, 4, Math.min(d.length, 17), 1).setValues(d);
    } else {
      dash.getRange('D24').setValue('no diagnostics');
    }
  } catch (e) {}
}
