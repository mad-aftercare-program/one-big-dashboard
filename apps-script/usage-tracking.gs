/* ==========================================================================
   USAGE TRACKING for One Big Aftercare  (Code.gs in One-Big-Aftercare-Master)
   Three small edits. Do NOT replace the whole file: your real keys live in it.
   ========================================================================== */

/* ---- EDIT 1 of 3 ----------------------------------------------------------
   In function doPostInner_(e), find this line:

     var viewOk = action === 'exportSheet' || action === 'postCheer';

   and replace it with:

     var viewOk = action === 'exportSheet' || action === 'postCheer' || action === 'track';
*/

/* ---- EDIT 2 of 3 ----------------------------------------------------------
   In the same function, find this line:

     if (action === 'ping') return json_({ ok: true });

   and paste these two lines directly BELOW it:

     if (action === 'track') return json_(trackUsage_(body));
     if (action === 'usage') return json_(usageData_());
*/

/* ---- EDIT 3 of 3 ----------------------------------------------------------
   Paste everything below at the very bottom of Code.gs.
*/

// The "Usage Log" tab in this master sheet. Created the first time it is needed.
function usageSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Usage Log');
  if (!sh) {
    sh = ss.insertSheet('Usage Log');
    sh.appendRow(['Timestamp', 'Person (device code)', 'View', 'Page', 'Chapter', 'Screen']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd mmm yyyy, hh:mm');
  }
  return sh;
}

// One row each time someone opens a page on the dashboard. No names, no phone numbers.
function trackUsage_(b) {
  var clean = function (v) { return String(v || '').slice(0, 40); };
  usageSheet_().appendRow([new Date(), clean(b.dev), clean(b.role), clean(b.page), clean(b.chapter), clean(b.screen)]);
  return { ok: true };
}

// For Admin > Usage on the dashboard (admin key only). Sends the latest 5,000 rows.
function usageData_() {
  var sh = usageSheet_(), last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var start = Math.max(2, last - 4999);
  var rows = sh.getRange(start, 1, last - start + 1, 6).getValues().map(function (r) {
    return [r[0] instanceof Date ? r[0].toISOString() : String(r[0]), r[1], r[2], r[3], r[4], r[5]];
  });
  return { ok: true, rows: rows };
}

/* ---- THEN ---------------------------------------------------------------
   Save. Deploy > Manage deployments > pencil (Edit) > Version: New version
   > Deploy. Keep the same deployment so the dashboard link does not change.
*/
