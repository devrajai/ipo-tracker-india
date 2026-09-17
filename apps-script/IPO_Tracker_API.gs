/**
 * ============================================================================
 *  INDIA IPO TRACKER - API & Link Health (add-on file)
 * ============================================================================
 *  Add this as a SECOND file in the same Apps Script project as
 *  IPO_Tracker_AppsScript.gs (Extensions > Apps Script > + next to Files).
 *
 *  WHAT IT ADDS
 *    1. doGet()          - a JSON API that the public website can read:
 *                          GET <web-app-url>            -> everything
 *                          GET <web-app-url>?tab=open  -> one tab only
 *                          Deploy: Deploy > New deployment > Web app >
 *                          Execute as: Me, Who has access: Anyone. Then put
 *                          the /exec URL into config.js (FEED_URL) of the
 *                          website. GMP % is COMPUTED here, never stored.
 *    2. resolveIpoLinks()- scrapes each Open/Upcoming IPO's detail page for
 *                          DIRECT RHP/DRHP pdf links (issuer/registrar hosted,
 *                          bypassing NSE/BSE gateway pages) and the registrar
 *                          allotment-status link. Writes them into extra
 *                          columns it adds itself: Allotment URL, RHP URL,
 *                          DRHP URL.
 *    3. checkLinks()     - the self-healing monitor: checks every stored link
 *                          (30-min trigger). On 404/410 it re-runs the resolver
 *                          and replaces the dead link. Statuses are logged to
 *                          a "Link Health" tab with a timestamped history.
 *    4. installLinkTrigger() - run ONCE to create the every-30-minutes
 *                          time-driven trigger for checkLinks().
 *
 *  QUOTAS (free Google account, approximate)
 *    UrlFetchApp: ~20,000 calls/day. checkLinks + resolver stay well below it
 *    at a 30-minute cadence with typical IPO counts. Do not lower the interval
 *    below 10 minutes.
 * ============================================================================
 */

var API = {
  TABS: ['Open', 'Upcoming', 'Closed', 'Listed'],
  COL_ALLOTMENT: 'Allotment URL',
  COL_RHP: 'RHP URL',
  COL_DRHP: 'DRHP URL',
  COL_LINK_STATUS: 'Link Status',
  COL_LAST_CHECKED: 'Last Checked',
  HEALTH_TAB: 'Link Health',
  FETCH_TIMEOUT: 20,          // seconds per page fetch
  LINK_TRIGGER_MINUTES: 30    // how often checkLinks() runs
};

/* ---------------------------------------------------------------------------
 * 1. JSON API  (the website's "Google Sheet feed")
 * ------------------------------------------------------------------------- */

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var only = params.tab || null;
  var payload = {
    ok: true,
    updated: new Date().toISOString(),
    stats: {
      open: countRows_('Open'),
      upcoming: countRows_('Upcoming'),
      closed: countRows_('Closed'),
      listed: countRows_('Listed')
    }
  };

  for (var i = 0; i < API.TABS.length; i++) {
    var tab = API.TABS[i];
    if (only && tab.toLowerCase() !== String(only).toLowerCase()) continue;
    payload[tab.toLowerCase()] = readTab_(tab);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Reads one tab into an array of objects keyed by camelCase names. */
function readTab_(tabName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(String);
  var idx = {};                                    // header -> column index
  for (var c = 0; c < headers.length; c++) idx[headers[c]] = c;

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var name = values[r][0];
    if (!name || !String(name).trim()) continue;   // skip blank rows
    rows.push({
      name: String(name),
      sector: cell_(values[r], idx['Sector']),
      openDate: cell_(values[r], idx['Open Date']),
      closeDate: cell_(values[r], idx['Close Date']),
      listingDate: cell_(values[r], idx['Listing Date']),
      priceLow: num_(cell_(values[r], idx['Price Band Low (Rs)'])),
      priceHigh: num_(cell_(values[r], idx['Price Band High (Rs)'])),
      lotSize: num_(cell_(values[r], idx['Lot Size (Shares)'])),
      lotValue: num_(cell_(values[r], idx['Lot Value (Rs)'])),
      issueSizeCr: num_(cell_(values[r], idx['Issue Size (Rs Cr)'])),
      rii: num_(cell_(values[r], idx['RII Sub (x)'])),
      qib: num_(cell_(values[r], idx['QIB Sub (x)'])),
      nii: num_(cell_(values[r], idx['NII Sub (x)'])),
      overallSub: num_(cell_(values[r], idx['Overall Sub (x)'])),
      gmp: num_(cell_(values[r], idx['GMP (Rs)'])),
      // GMP % is always computed fresh - it can never go stale in storage.
      gmpPct: gmpPct_(cell_(values[r], idx['GMP (Rs)']),
                      cell_(values[r], idx['Price Band Low (Rs)'])),
      de: cell_(values[r], idx['Debt to Equity']),
      roe: cell_(values[r], idx['ROE (%)']),
      revGrowth: cell_(values[r], idx['Revenue Growth (%)']),
      listingPrice: num_(cell_(values[r], idx['Listing Price (Rs)'])),
      actualGain: cell_(values[r], idx['Actual Listing Gain (%)']),
      detailUrl: cell_(values[r], idx['Detail URL']),
      allotmentUrl: cell_(values[r], idx[API.COL_ALLOTMENT]),
      rhpUrl: cell_(values[r], idx[API.COL_RHP]),
      drhpUrl: cell_(values[r], idx[API.COL_DRHP]),
      linkStatus: cell_(values[r], idx[API.COL_LINK_STATUS]),
      lastChecked: cell_(values[r], idx[API.COL_LAST_CHECKED])
    });
  }
  return rows;
}

function gmpPct_(gmpVal, priceLowVal) {
  var g = parseFloat(gmpVal), p = parseFloat(priceLowVal);
  if (!isFinite(g) || !isFinite(p) || p <= 0) return null;
  return Math.round((g / p) * 1000) / 10;          // one decimal
}

function cell_(row, colIdx) {
  if (colIdx === undefined || colIdx < 0) return '';
  var v = row[colIdx];
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(v).trim();
}

function num_(v) {
  var n = parseFloat(v);
  return isFinite(n) ? n : null;
}

function countRows_(tabName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return 0;
  var names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  var n = 0;
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0] || '').trim()) n++;
  }
  return n;
}

/* ---------------------------------------------------------------------------
 * 2. Link resolver - direct RHP/DRHP + allotment links
 *    Priority: issuer/registrar-hosted PDF first, NSE/BSE gateway only as
 *    fallback (the gateway links are what we try to bypass).
 * ------------------------------------------------------------------------- */

function resolveIpoLinks() {
  ensureColumns_(['Open', 'Upcoming']);
  var ss = SpreadsheetApp.getActive();
  var resolved = 0;

  ['Open', 'Upcoming'].forEach(function (tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh) return;
    var values = sh.getDataRange().getValues();
    var headers = values[0].map(String);
    var col = function (h) { return headers.indexOf(h); };
    var cDetail = col('Detail URL'), cAllot = col(API.COL_ALLOTMENT),
        cRhp = col(API.COL_RHP), cDrhp = col(API.COL_DRHP),
        cStatus = col(API.COL_LINK_STATUS), cChecked = col(API.COL_LAST_CHECKED);
    if (cDetail < 0) return;

    for (var r = 1; r < values.length; r++) {
      var name = String(values[r][0] || '').trim();
      var detailUrl = String(values[r][cDetail] || '').trim();
      if (!name || !detailUrl) continue;

      var links = scrapeLinks_(detailUrl);
      if (links.rhp || links.drhp || links.allotment) resolved++;

      if (links.allotment) values[r][cAllot] = links.allotment;
      if (links.rhp) values[r][cRhp] = links.rhp;
      if (links.drhp) values[r][cDrhp] = links.drhp;
      values[r][cStatus] = links.any ? 'resolved ' + Utilities.formatDate(
        new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm') : 'no links found';
      values[r][cChecked] = new Date();
    }
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  });

  logHealth_('resolver run: ' + resolved + ' IPO(s) with resolvable links');
  return resolved;
}

/**
 * Fetches an IPO detail page and looks for:
 *  - RHP / DRHP: direct .pdf links (issuer or registrar hosted preferred,
 *    exchange gateways like nseindia.com/bseindia.com de-prioritised).
 *  - Allotment: registrar / allotment-status links.
 * Best-effort: returns whatever it finds, never throws.
 */
function scrapeLinks_(detailUrl) {
  var out = { rhp: '', drhp: '', allotment: '', any: false };
  try {
    var resp = UrlFetchApp.fetch(detailUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: false,
      timeout: API.FETCH_TIMEOUT * 1000
    });
    var code = resp.getResponseCode();
    if (code !== 200) { out.any = false; return out; }
    var html = resp.getContentText();

    var hrefs = html.match(/href\s*=\s*["']([^"']+)["']/ig) || [];
    var pdfs = [], others = [];
    hrefs.forEach(function (m) {
      var url = m.replace(/^href\s*=\s*["']/i, '').replace(/["']$/, '');
      var low = url.toLowerCase();
      if (low.indexOf('javascript:') === 0 || low.indexOf('#') === 0) return;
      if (low.slice(-4) === '.pdf') pdfs.push(url); else others.push(url);
    });

    var abs = function (u) {
      if (/^https?:\/\//i.test(u)) return u;
      if (u.indexOf('//') === 0) return 'https:' + u;
      if (u.charAt(0) === '/') {
        var m = detailUrl.match(/^(https?:\/\/[^\/]+)/i);
        return m ? m[1] + u : u;
      }
      return detailUrl.replace(/[^\/]*$/, '') + u;
    };

    // RHP/DRHP: prefer direct PDFs whose filename says rhp / red herring /
    // prospectus. De-prioritise exchange gateways.
    var isExchange = function (u) {
      return /nseindia\.com|bseindia\.com/i.test(u);
    };
    pdfs.forEach(function (u) {
      var full = abs(u), low = full.toLowerCase();
      if (/rhp|red[ _-]?herring/.test(low) && !out.rhp) {
        out.rhp = isExchange(full) ? out.rhp || full : full;
      }
    });
    pdfs.forEach(function (u) {
      var full = abs(u), low = full.toLowerCase();
      if (/drhp|draft[ _-]?red[ _-]?herring/.test(low) && !out.drhp) {
        out.drhp = isExchange(full) ? out.drhp || full : full;
      }
    });
    // Generic prospectus pdf if no explicit rhp tag was found.
    if (!out.rhp) {
      pdfs.forEach(function (u) {
        var full = abs(u), low = full.toLowerCase();
        if (/prospectus/.test(low) && !isExchange(full) && !out.rhp) out.rhp = full;
      });
    }

    // Allotment: registrar status-page links. Common registrars in India:
    others.concat(pdfs).forEach(function (u) {
      var full = abs(u), low = full.toLowerCase();
      if (!out.allotment && /allotment|bigshare|linkintime|kfintech|skyline|mas|unistart|registrar/.test(low)) {
        out.allotment = full;
      }
    });

    out.any = !!(out.rhp || out.drhp || out.allotment);
  } catch (err) {
    // network hiccups are fine - the next run retries
  }
  return out;
}

/* ---------------------------------------------------------------------------
 * 3. Self-healing link monitor
 *    IMPORTANT: a 404 on an allotment link usually means the registrar has
 *    not published yet - that is "pending", not "dead". The resolver always
 *    re-derives the link from the detail page, never guesses.
 * ------------------------------------------------------------------------- */

function checkLinks() {
  ensureColumns_(['Open', 'Upcoming']);
  var ss = SpreadsheetApp.getActive();
  var toFetch = [];      // {url}
  var targets = [];      // {sheet, row, colIdx, kind, url, ipo}

  ['Open', 'Upcoming'].forEach(function (tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh) return;
    var values = sh.getDataRange().getValues();
    var headers = values[0].map(String);
    [[API.COL_ALLOTMENT, 'allotment'], [API.COL_RHP, 'rhp'], [API.COL_DRHP, 'drhp']]
      .forEach(function (pair) {
        var c = headers.indexOf(pair[0]);
        if (c < 0) return;
        for (var r = 1; r < values.length; r++) {
          var url = String(values[r][c] || '').trim();
          if (/^https?:\/\//i.test(url)) {
            targets.push({ sheet: sh, row: r + 1, colIdx: c + 1, kind: pair[1], url: url, ipo: values[r][0] });
            toFetch.push({ url: url });
          }
        }
      });
  });

  if (!toFetch.length) { logHealth_('checkLinks: nothing to check'); return; }

  var responses = UrlFetchApp.fetchAll(toFetch.map(function (t) {
    return { url: t.url, muteHttpExceptions: true, followRedirects: true,
             validateHttpsCertificates: false, timeout: API.FETCH_TIMEOUT * 1000 };
  }));

  var dead = 0, ok = 0;
  for (var i = 0; i < responses.length; i++) {
    var code = responses[i].getResponseCode();
    var t = targets[i];
    if (code === 200 || code === 301 || code === 302) {
      ok++;
    } else if (code === 404 || code === 410) {
      dead++;
      healLink_(t);   // re-derive from the IPO's detail page
    }
    setStatus_(t, code);
  }
  logHealth_('checkLinks: ' + ok + ' ok, ' + dead + ' dead and re-resolved, of ' + responses.length + ' checked');
}

/** Re-derives one dead link by re-scraping the IPO's detail page. */
function healLink_(target) {
  try {
    var detail = getDetailUrl_(target.sheet, target.row);
    if (!detail) return;
    var links = scrapeLinks_(detail);
    var fresh = target.kind === 'rhp' ? links.rhp
              : target.kind === 'drhp' ? links.drhp : links.allotment;
    if (fresh && fresh !== target.url) {
      target.sheet.getRange(target.row, target.colIdx).setValue(fresh);
      logHealth_('healed ' + target.kind + ' link for ' + target.ipo + ' -> ' + fresh);
    } else {
      logHealth_('could not re-resolve ' + target.kind + ' for ' + target.ipo + ' (may not be published yet - will retry)');
    }
  } catch (err) {
    logHealth_('healLink_ error for ' + target.ipo + ': ' + err);
  }
}

function getDetailUrl_(sheet, row1based) {
  var values = sheet.getRange(row1based, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var c = headers.indexOf('Detail URL');
  if (c < 0) return '';
  return String(values[c] || '').trim();
}

function setStatus_(target, code) {
  try {
    var headers = target.sheet.getRange(1, 1, 1, target.sheet.getLastColumn()).getValues()[0].map(String);
    var cStatus = headers.indexOf(API.COL_LINK_STATUS) + 1;
    var cChecked = headers.indexOf(API.COL_LAST_CHECKED) + 1;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
    if (cStatus > 0) target.sheet.getRange(target.row, cStatus)
      .setValue(target.kind + ' ' + code + ' @ ' + now);
    if (cChecked > 0) target.sheet.getRange(target.row, cChecked).setValue(new Date());
  } catch (err) { /* non-fatal */ }
}

/* ---------------------------------------------------------------------------
 * 4. Setup helpers
 * ------------------------------------------------------------------------- */

/** Appends the extra columns this file needs, if they are not there yet. */
function ensureColumns_(tabNames) {
  var ss = SpreadsheetApp.getActive();
  tabNames.forEach(function (tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh) return;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) {
      return String(h || '').trim();
    });
    [API.COL_ALLOTMENT, API.COL_RHP, API.COL_DRHP, API.COL_LINK_STATUS, API.COL_LAST_CHECKED]
      .forEach(function (h) {
        if (headers.indexOf(h) < 0) {
          sh.insertColumnsAfter(sh.getLastColumn(), 1);
          var col = sh.getLastColumn();
          sh.getRange(1, col).setValue(h).setFontWeight('bold');
          headers.push(h);
        }
      });
  });
}

/** Run once: every-30-min trigger for the self-healing monitor. */
function installLinkTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkLinks') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkLinks').timeBased()
    .everyMinutes(API.LINK_TRIGGER_MINUTES).create();
  logHealth_('link-health trigger installed (every ' + API.LINK_TRIGGER_MINUTES + ' min)');
}

/** Appends a timestamped line to the Link Health tab (created on demand). */
function logHealth_(msg) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(API.HEALTH_TAB);
  if (!sh) {
    sh = ss.insertSheet(API.HEALTH_TAB);
    sh.getRange(1, 1, 1, 2).setValues([['Time', 'Event']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), msg]);
}
