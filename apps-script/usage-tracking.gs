/* ==========================================================================
   USAGE TRACKING for One Big Aftercare: the 3 changes to Code.gs
   (matches the Code.gs that uses CONFIG, keyOk_, adminOk_, json_, ensureTab_)
   ========================================================================== */

/* 1. In  var TAB = { ... }  add at the end of the list:
        , USAGE: 'Usage Log'                                                   */

/* 2. In  function doPostInner_(e)  right after the line
        var action = body.action;
      add these two lines:                                                     */

  if (action === 'track') return (keyOk_(body.key) || adminOk_(body.key)) ? json_(trackUsage_(body)) : json_({ ok: false, error: 'Wrong key.' });
  if (action === 'usage') return adminOk_(body.key) ? json_(usageData_()) : json_({ ok: false, error: 'Wrong key.' });

/* 3. Paste these functions anywhere (for example right after function text_):  */

var USAGE_HEAD = ['When', 'Person (device code)', 'View', 'Page', 'Chapter', 'Screen'];

// One row each time someone opens a page. No names or numbers: the person is a random code made by their browser.
function trackUsage_(b) {
  var clean = function (v) { return String(v == null ? '' : v).slice(0, 40); };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    var sh = ensureTab_(SpreadsheetApp.getActiveSpreadsheet(), TAB.USAGE, USAGE_HEAD);
    sh.appendRow([new Date(), clean(b.dev), clean(b.role), clean(b.page), clean(b.chapter), clean(b.screen)]);
  } finally { try { lock.releaseLock(); } catch (z) {} }
  return { ok: true };
}

// The last 5,000 rows for the Admin > Usage tab. Admin key only.
function usageData_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.USAGE);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var last = sh.getLastRow(), start = Math.max(2, last - 4999);
  var rows = sh.getRange(start, 1, last - start + 1, USAGE_HEAD.length).getValues().map(function (r) {
    return [r[0] instanceof Date ? r[0].toISOString() : String(r[0]), r[1], r[2], r[3], r[4], r[5]];
  });
  return { ok: true, rows: rows };
}

/* Then: Save > Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy. */
