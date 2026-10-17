/**
 * Supply Day Book — office backend (Google Apps Script, bound to the Sheet).
 *
 * Tabs are created automatically on first call:
 *   Supplymen : id | name | pin | role | active        (role: supplyman / admin)
 *   Bills     : date | supplyman_id | bill_no | party | amount | mobile | otp
 *   Entries   : one row per saved entry from the phones (idempotent on entry_id)
 *
 * Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone.
 * The phone app POSTs JSON {action, pin, ...} and gets JSON back.
 */

var TZ = "Asia/Kolkata";
var PHOTO_FOLDER = "Supply Day Book Photos";

var ENTRY_COLS = ["entry_id","date","supplyman_id","supplyman_name","bill_no","party","bill_amount",
  "status","mode","received","balance","promised","note","otp_status","photo_status","photo_url",
  "lat","lng","accuracy_m","gps_time","device_time","device","server_time","map_link"];

function doGet(e) {
  return out({ ok: true, service: "Supply Day Book", time: new Date().toISOString() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = {};
    try { body = JSON.parse(e.postData.contents || "{}"); } catch (err) { return out({ ok: false, code: "bad_json", error: "Bad request" }); }
    ensureSheets();
    var action = String(body.action || "");
    if (action === "login")   return out(login(body));
    var user = auth(body.pin);
    if (!user) return out({ ok: false, code: "bad_pin", error: "Wrong PIN" });
    if (action === "bills")   return out(bills(user, body));
    if (action === "sync")    return out(sync(user, body));
    if (action === "summary") return out(summary(user, body));
    return out({ ok: false, code: "unknown", error: "Unknown action" });
  } catch (err) {
    return out({ ok: false, code: "server", error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheets ---------- */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name) { return ss().getSheetByName(name); }

function ensureSheets() {
  var s = ss();
  if (!s.getSheetByName("Supplymen")) {
    var sm = s.insertSheet("Supplymen");
    sm.appendRow(["id","name","pin","role","active"]);
    sm.appendRow(["S1","Test Supplyman","1234","supplyman","yes"]);
    sm.appendRow(["ADMIN","Office","9999","admin","yes"]);
    sm.setFrozenRows(1);
  }
  if (!s.getSheetByName("Bills")) {
    var b = s.insertSheet("Bills");
    b.appendRow(["date","supplyman_id","bill_no","party","amount","mobile","otp"]);
    b.appendRow([todayStr(),"S1","10241","Sharma Medicos",4820,"9810012345",""]);
    b.appendRow([todayStr(),"S1","10242","Gupta Medical Hall",1265,"",""]);
    b.setFrozenRows(1);
  }
  if (!s.getSheetByName("Entries")) {
    var en = s.insertSheet("Entries");
    en.appendRow(ENTRY_COLS);
    en.setFrozenRows(1);
  }
  var first = s.getSheets()[0];
  if (first.getName() === "Sheet1" && first.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(first);
}

function rows(name) {
  var sh = sheet(name); if (!sh) return [];
  var v = sh.getDataRange().getValues(); if (v.length < 2) return [];
  var head = v[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var outRows = [];
  for (var i = 1; i < v.length; i++) {
    var r = {}, empty = true;
    for (var j = 0; j < head.length; j++) { r[head[j]] = v[i][j]; if (v[i][j] !== "" && v[i][j] != null) empty = false; }
    if (!empty) { r._row = i + 1; outRows.push(r); }
  }
  return outRows;
}

/* ---------- auth ---------- */
function auth(pin) {
  pin = String(pin || "").trim(); if (!pin) return null;
  var list = rows("Supplymen");
  for (var i = 0; i < list.length; i++) {
    var u = list[i];
    if (String(u.pin).trim() === pin && String(u.active || "yes").toLowerCase() !== "no")
      return { id: String(u.id), name: String(u.name), role: String(u.role || "supplyman").toLowerCase() };
  }
  return null;
}
function login(body) {
  var u = auth(body.pin);
  if (!u) return { ok: false, code: "bad_pin", error: "Wrong PIN" };
  return { ok: true, user: u };
}

/* ---------- bills ---------- */
function bills(user, body) {
  var d = normDate(body.date) || todayStr();
  var list = rows("Bills"), outB = [];
  for (var i = 0; i < list.length; i++) {
    var r = list[i];
    if (normDate(r.date) !== d) continue;
    if (user.role !== "admin" && String(r.supplyman_id).trim() !== user.id) continue;
    outB.push({ no: String(r.bill_no).trim(), party: String(r.party || "").trim(), amt: Number(r.amount) || 0,
      mobile: String(r.mobile || "").replace(/\D/g, ""), otp: String(r.otp || "").replace(/\D/g, "") });
  }
  return { ok: true, date: d, bills: outB };
}

/* ---------- sync ---------- */
function sync(user, body) {
  var entries = body.entries || [];
  var sh = sheet("Entries");
  var existing = indexByEntryId(sh);
  var photos = {}, saved = [];
  for (var i = 0; i < entries.length; i++) {
    var en = entries[i]; if (!en || !en.id) continue;
    if (user.role !== "admin" && String(en.supplymanId) !== user.id) continue; // a phone may only write its own supplyman's rows
    var photoUrl = en.photoUrl || "";
    if (en.photo && !photoUrl) {
      try { photoUrl = savePhoto(en); photos[en.id] = photoUrl; } catch (err) { photoUrl = ""; }
    }
    var mapLink = (en.lat && en.lng) ? "https://maps.google.com/?q=" + en.lat + "," + en.lng : "";
    var row = [en.id, en.date, en.supplymanId, en.supplymanName, en.billNo, en.party, Number(en.billAmount) || 0,
      en.status, en.mode || "", Number(en.received) || 0, Number(en.balance) || 0, en.promised || "", en.note || "",
      en.otpStatus || "", en.photoStatus || "", photoUrl, en.lat || "", en.lng || "", en.acc || "", en.gpsTime || "",
      en.at || "", en.device || "", new Date().toISOString(), mapLink];
    if (existing[en.id]) sh.getRange(existing[en.id], 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); existing[en.id] = sh.getLastRow(); }
    saved.push(en.id);
  }
  return { ok: true, saved: saved, photos: photos };
}

function indexByEntryId(sh) {
  var map = {}; var last = sh.getLastRow(); if (last < 2) return map;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0]) map[String(ids[i][0])] = i + 2;
  return map;
}

function savePhoto(en) {
  var folder = photoFolder();
  var m = /^data:(image\/\w+);base64,(.+)$/.exec(en.photo || "");
  if (!m) return "";
  var bytes = Utilities.base64Decode(m[2]);
  var name = [en.date, en.supplymanName, en.billNo, en.id].join("_").replace(/[^\w\-]+/g, "_") + ".jpg";
  var file = folder.createFile(Utilities.newBlob(bytes, m[1], name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return "https://drive.google.com/file/d/" + file.getId() + "/view";
}
function photoFolder() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

/* ---------- office summary ---------- */
function summary(user, body) {
  if (user.role !== "admin") return { ok: false, code: "forbidden", error: "Office login needed" };
  var d = normDate(body.date) || todayStr();
  var list = rows("Entries"), outE = [];
  for (var i = 0; i < list.length; i++) {
    var r = list[i]; if (normDate(r.date) !== d) continue;
    outE.push({ id: r.entry_id, date: d, supplymanId: r.supplyman_id, supplymanName: r.supplyman_name, billNo: r.bill_no, party: r.party,
      billAmount: r.bill_amount, status: r.status, mode: r.mode, received: r.received, balance: r.balance, promised: r.promised,
      note: r.note, otpStatus: r.otp_status, photoStatus: r.photo_status, photoUrl: r.photo_url, lat: r.lat, lng: r.lng, acc: r.accuracy_m,
      at: r.device_time instanceof Date ? r.device_time.toISOString() : String(r.device_time || "") });
  }
  return { ok: true, date: d, entries: outE };
}

/* ---------- dates ---------- */
function todayStr() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"); }
function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  var s = String(v || "").trim(); if (!s) return "";
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s); if (m) return m[1] + "-" + p(m[2]) + "-" + p(m[3]);
  m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(s); if (m) return m[3] + "-" + p(m[2]) + "-" + p(m[1]); // dd/mm/yyyy (Marg style)
  return s;
}
function p(n) { n = String(n); return n.length < 2 ? "0" + n : n; }
