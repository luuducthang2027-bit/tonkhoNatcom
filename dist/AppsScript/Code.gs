/* Tồn kho các chi nhánh – Web App dùng chung.
   Triển khai: Thực thi với tư cách "Tôi", ai có quyền truy cập: người trong tổ chức (hoặc danh sách cụ thể).
   Dữ liệu chung lưu NGAY TRONG GOOGLE SHEET:
   - Nếu script mở từ Google Sheet (Tiện ích mở rộng > Apps Script) thì dùng chính Sheet đó (vd "Tồn Biểu Đồ").
   - Nếu script độc lập thì tự tạo 1 Google Sheet tên "Tồn kho chi nhánh - Dữ liệu chung" trong Drive của người triển khai.
   Sheet "PhienChung": mỗi ngày chốt gồm nhiều dòng (dữ liệu nén cắt khúc) – đừng sửa tay. */
var SHEET = 'PhienChung', CHUNK = 45000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tồn kho các chi nhánh')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function ss_() {
  var s = null;
  try { s = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (s) return s;
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SSID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  s = SpreadsheetApp.create('Tồn kho chi nhánh - Dữ liệu chung');
  p.setProperty('SSID', s.getId());
  return s;
}
function sh_() {
  var s = ss_(), sh = s.getSheetByName(SHEET);
  if (!sh) {
    sh = s.insertSheet(SHEET);
    sh.getRange('A1:E1').setValues([['Ngày chốt', 'Cập nhật (ms)', 'Người lưu', 'Khúc', 'Dữ liệu nén – không sửa tay']]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@'); sh.getRange('E:E').setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}
function who_() { try { return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } }
function find_(sh, cut) {
  var n = sh.getLastRow(); if (n < 2) return { first: 0, count: 0 };
  var v = sh.getRange(2, 1, n - 1, 1).getValues(), first = 0, count = 0;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === cut) { if (!first) first = i + 2; count++; }
  return { first: first, count: count };
}
function saveSnap(cut, txt) {
  cut = String(cut);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cut)) throw new Error('Ngày chốt không hợp lệ');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sh_(), f = find_(sh, cut), at = new Date().getTime(), by = who_(), rows = [];
    if (f.count) sh.deleteRows(f.first, f.count);
    for (var i = 0, k = 0; i < txt.length; i += CHUNK, k++) rows.push([cut, String(at), by, String(k), txt.substr(i, CHUNK)]);
    var r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, rows.length, 5).setNumberFormat('@').setValues(rows);
    SpreadsheetApp.flush();
    return { id: cut, name: 'ton_kho_' + cut, at: at, by: by };
  } finally { lock.releaseLock(); }
}
function listSnaps() {
  var sh = sh_(), n = sh.getLastRow(), out = [], seen = {};
  if (n < 2) return out;
  var v = sh.getRange(2, 1, n - 1, 3).getValues();
  for (var i = 0; i < v.length; i++) {
    var c = String(v[i][0]);
    if (/^\d{4}-\d{2}-\d{2}$/.test(c) && !seen[c]) { seen[c] = 1; out.push({ id: c, name: 'ton_kho_' + c, at: Number(v[i][1]) || 0, by: String(v[i][2] || '') }); }
  }
  out.sort(function (a, b) { return a.id < b.id ? 1 : -1; });
  return out;
}
function loadSnap(id) {
  var sh = sh_(), f = find_(sh, String(id));
  if (!f.count) throw new Error('Không thấy phiên ' + id + ' trong Sheet');
  var v = sh.getRange(f.first, 4, f.count, 2).getValues();
  v.sort(function (a, b) { return Number(a[0]) - Number(b[0]); });
  return v.map(function (r) { return String(r[1]); }).join('');
}
/* Chạy thử 1 lần trong trình soạn thảo để cấp quyền Sheet/Drive. */
function authorize() { return ss_().getUrl(); }
