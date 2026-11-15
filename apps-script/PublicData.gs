/**
 * One Big Aftercare: public read-only data  (add-on file)
 * ---------------------------------------------------------------------------
 * Lets the dashboard load its numbers on ANY device, with no link or key to
 * paste. It adds one read-only request:  <web app link>?action=public
 *
 * Keep this file BELOW Code.gs in the Apps Script file list. It does not change
 * Code.gs. Delete this file (and deploy a new version) to switch it off.
 *
 * What the public request gives
 *   The same numbers the dashboard counts from, with every name and detail
 *   taken out first:
 *   - Youth codes become random-looking codes. Wingman names become "Wingman 1",
 *     "Wingman 2" and so on (per chapter).
 *   - No emails, phone numbers, PC full names, conversation notes, crisis
 *     details, sheet links, sync log or dashboard passwords.
 *   - Crises are sent as chapter and status only (for the open-crisis count).
 *   It can only read. Every save still needs the access key.
 *
 * The access key never leaves this script. The add-on finds it itself by
 * asking keyOk_() which saved value is the key.
 */

var doGetBeforePublic_ = doGet;
doGet = function (e) {
  var p = (e && e.parameter) || {};
  if (p.action !== 'public') return doGetBeforePublic_(e);
  try {
    var key = publicSourceKey_();
    if (!key) return publicJson_({ ok: false, error: 'The public view is not set up yet.' });
    var q = { action: 'data', key: key, as: 'frontline' };
    if (p.fresh) q.fresh = '1';
    var out = doGetBeforePublic_({ parameter: q, parameters: {} });
    var r = JSON.parse(out.getContent());
    if (!r || !r.ok || !r.data) return publicJson_({ ok: false, error: 'Could not load the numbers.' });
    return publicJson_({ ok: true, public: true, data: publicData_(r.data), builtAt: r.builtAt || '', src: r.src || '' });
  } catch (err) {
    return publicJson_({ ok: false, error: 'Could not load the numbers.' });
  }
};

function publicJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// The access key, found without anyone typing it: the first saved value that keyOk_ accepts.
function publicSourceKey_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('public_src_key_ok');
  var cands = [];
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG) Object.keys(CONFIG).forEach(function (k) { if (typeof CONFIG[k] === 'string') cands.push(CONFIG[k]); });
  } catch (z) {}
  var props = PropertiesService.getScriptProperties().getProperties();
  if (props.PUBLIC_SOURCE_KEY) cands.unshift(props.PUBLIC_SOURCE_KEY);
  Object.keys(props).forEach(function (k) { cands.push(props[k]); });
  if (hit && cands.indexOf(hit) >= 0) return hit;
  if (typeof keyOk_ !== 'function') return props.PUBLIC_SOURCE_KEY || '';
  var admin = typeof adminOk_ === 'function' ? adminOk_ : function () { return false; }, fallback = '';
  for (var i = 0; i < cands.length; i++) {
    var v = String(cands[i] || '');
    if (!v || v.length > 200) continue;
    var ok = false; try { ok = keyOk_(v); } catch (z) {}
    if (!ok) continue;
    var adm = false; try { adm = admin(v); } catch (z) {}
    if (!adm) { cache.put('public_src_key_ok', v, 21600); return v; }
    if (!fallback) fallback = v;
  }
  return fallback;
}

// A stable random-looking code for a youth code. The salt is made once and kept in this script.
function publicSalt_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('PUBLIC_SALT');
  if (!s) { s = Utilities.getUuid(); p.setProperty('PUBLIC_SALT', s); }
  return s;
}
function publicCode_(salt, v) {
  v = String(v == null ? '' : v).trim();
  if (!v) return '';
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + v), h = '';
  for (var i = 0; i < 6; i++) h += ((d[i] + 256) % 256).toString(16).padStart(2, '0');
  return 'Y-' + h;
}

function publicData_(d) {
  var salt = publicSalt_(), norm = function (x) { return String(x == null ? '' : x).toLowerCase().replace(/\s+/g, ' ').trim(); };
  var arr = function (x) { return Array.isArray(x) ? x : []; };

  // Wingmen: "Wingman 1", "Wingman 2" ... per chapter. Every spelling of a name maps to the same label.
  var label = {}, count = {}, wingmen = arr(d.wingmen).map(function (w) {
    var ch = w.chapter || '';
    count[ch] = (count[ch] || 0) + 1;
    var t = 'Wingman ' + count[ch];
    [w.name, w.full].concat(arr(w.alts)).forEach(function (n) { if (n) label[ch + '|' + norm(n)] = t; });
    return { id: t + ' (' + ch + ')', name: t, chapter: ch, status: w.status || '', youth: w.youth || 0, first: w.first || '', convos: w.convos || 0, hours: w.hours || 0, full: '', alts: [] };
  });
  var wmLabel = function (ch, n) { return n ? (label[ch + '|' + norm(n)] || 'Wingman (other)') : ''; };

  var youth = arr(d.youth).map(function (y) {
    return [publicCode_(salt, y[0]), y[1], wmLabel(y[1], y[2]), y[3] == null ? '' : y[3]];
  });
  var interactions = arr(d.interactions).map(function (i, n) {
    return { id: 'i' + n, youth: publicCode_(salt, i.youth), chapter: i.chapter, wingman: wmLabel(i.chapter, i.wingman), date: i.date, month: i.month, happened: i.happened };
  });
  var uploads = arr(d.uploads).map(function (u) {
    return { id: u.id, chapter: u.chapter, type: u.type, date: u.date, month: u.month,
      wingmenPresent: arr(u.wingmenPresent).map(function (n) { return wmLabel(u.chapter, n); }),
      wingmenRows: u.wingmenRows || 0, youthPresent: u.youthPresent || 0, youthRows: u.youthRows || 0, saved: true };
  });
  var crises = arr(d.crises).map(function (c) { return { chapter: c.chapter, status: c.status }; });

  var settings = null;
  if (d.settings && typeof d.settings === 'object') {
    settings = JSON.parse(JSON.stringify(d.settings));
    delete settings.pw; delete settings.adminHash;
  }
  return {
    asOf: d.asOf || '', chapters: arr(d.chapters), youth: youth, wingmen: wingmen, calendar: arr(d.calendar),
    interactions: interactions, uploads: uploads, crises: crises, crisisLog: [], legacyCrises: [], projects: [],
    texts: d.texts || {}, cheers: arr(d.cheers), rules: arr(d.rules), config: d.config || {}, ageCounts: d.ageCounts || null,
    settings: settings
  };
}
