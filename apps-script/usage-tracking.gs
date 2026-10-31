/* ==========================================================================
   USAGE TRACKING for One Big Aftercare
   Paste this whole block at the BOTTOM of Code.gs, then make the 3-line
   change inside doPost described in STEP 2 below.
   It uses the ACCESS_KEY, ADMIN_KEY and MASTER_ID already set at the top
   of Code.gs.
   ========================================================================== */

// Creates the "Usage Log" tab in the master sheet the first time it is needed.
function usageSheet_() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  let sh = ss.getSheetByName('Usage Log');
  if (!sh) {
    sh = ss.insertSheet('Usage Log');
    sh.appendRow(['Timestamp', 'Person (device code)', 'View', 'Page', 'Chapter', 'Screen']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd mmm yyyy, hh:mm');
  }
  return sh;
}

function usageOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Called every time someone opens a page on the dashboard. Adds one row.
function trackUsage_(b) {
  if (b.key !== ACCESS_KEY && b.key !== ADMIN_KEY) return usageOut_({ ok: false, error: 'Wrong key' });
  const clean = v => String(v || '').slice(0, 40);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    usageSheet_().appendRow([new Date(), clean(b.dev), clean(b.role), clean(b.page), clean(b.chapter), clean(b.screen)]);
  } finally { lock.releaseLock(); }
  return usageOut_({ ok: true });
}

// Called by the Admin > Usage tab. Admin key only. Returns the last 5,000 rows.
function usageData_(b) {
  if (b.key !== ADMIN_KEY) return usageOut_({ ok: false, error: 'Admin key needed' });
  const sh = usageSheet_(), last = sh.getLastRow();
  if (last < 2) return usageOut_({ ok: true, rows: [] });
  const start = Math.max(2, last - 4999);
  const rows = sh.getRange(start, 1, last - start + 1, 6).getValues().map(r => [
    r[0] instanceof Date ? r[0].toISOString() : String(r[0]), r[1], r[2], r[3], r[4], r[5]
  ]);
  return usageOut_({ ok: true, rows: rows });
}

/* --------------------------------------------------------------------------
   STEP 2: inside your existing  function doPost(e) { ... }
   add these 3 lines as the VERY FIRST lines of the function:

     const _u = JSON.parse((e.postData && e.postData.contents) || '{}');
     if (_u.action === 'track') return trackUsage_(_u);
     if (_u.action === 'usage') return usageData_(_u);

   STEP 3: Save, then Deploy > Manage deployments > pencil (Edit) >
   Version: "New version" > Deploy. Keep the same deployment so the
   web app link does not change.
   -------------------------------------------------------------------------- */
