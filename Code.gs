/**
 * Tết triệu đô – Trao yêu thương
 * Backend Google Apps Script — dữ liệu lưu trong chính Google Sheet chứa script này.
 *
 * Cài đặt lần đầu: chọn hàm `setup` ở thanh công cụ → Run → cấp quyền.
 * Mật khẩu quản trị được tạo ngẫu nhiên và in ra ở "Execution log".
 * Đổi mật khẩu: Project Settings → Script Properties → ADMIN_PASSWORD.
 */

const TZ = 'Asia/Ho_Chi_Minh';

const T = {
  campaign: { name: 'Campaign', head: ['key', 'value', 'ghi_chu'] },
  offers: { name: 'Offers', head: ['id', 'kind', 'name', 'desc', 'price', 'unit', 'owner', 'stock', 'date'], num: ['price', 'stock'] },
  lots: { name: 'Lots', head: ['id', 'name', 'desc', 'donor', 'startPrice', 'step', 'endsAt'], num: ['startPrice', 'step'] },
  teams: { name: 'Teams', head: ['team'] },
  contribs: {
    name: 'Contributions',
    head: ['id', 'createdAt', 'name', 'team', 'contact', 'type', 'refId', 'refName', 'qty', 'amount', 'note', 'paid', 'paidAt', 'paidBy', 'payCode', 'undoToken'],
    num: ['qty', 'amount'], bool: ['paid'],
  },
  bids: { name: 'Bids', head: ['id', 'createdAt', 'lotId', 'name', 'team', 'contact', 'amount'], num: ['amount'] },
};

const KINDS = ['product', 'service', 'activity'];
const MIN_DONATION = 10000;
const MAX_AMOUNT = 2000000000;
const UNDO_MINUTES = 15;

/* ------------------------------------------------------------------ */
/* API entry — trang web trên GitHub Pages gọi vào đây                  */
/* Body (text/plain): {"fn": "<tên hàm>", "args": [...]}                 */
/* ------------------------------------------------------------------ */

const API = {
  getState: getState, addContribution: addContribution, undoContribution: undoContribution, placeBid: placeBid,
  adminLogin: adminLogin, adminGetState: adminGetState, adminSetPaid: adminSetPaid,
  adminDeleteContribution: adminDeleteContribution, adminRestoreContribution: adminRestoreContribution,
  adminRecordWin: adminRecordWin, adminSaveCampaign: adminSaveCampaign, adminUpsert: adminUpsert, adminDelete: adminDelete,
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API.hasOwnProperty(req.fn) ? API[req.fn] : null;
    if (!fn) throw new Error('Yêu cầu không hợp lệ.');
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Mở link Web app trên trình duyệt để kiểm tra API còn chạy */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'tet-trieu-do', time: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Setup & seed                                                         */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(T).forEach(function (k) {
    const t = T[k];
    let s = ss.getSheetByName(t.name);
    if (!s) s = ss.insertSheet(t.name);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold').setBackground('#FBF1D9');
      s.setFrozenRows(1);
      // Cột chữ để dạng văn bản thuần, tránh Sheets tự đổi ngày/giờ
      t.head.forEach(function (h, i) {
        const isNum = (t.num || []).indexOf(h) >= 0 || (t.bool || []).indexOf(h) >= 0;
        if (!isNum) s.getRange(1, i + 1, s.getMaxRows(), 1).setNumberFormat('@');
      });
    }
  });

  if (sheet_('campaign').getLastRow() < 2) {
    const rows = [
      ['title', 'Tết triệu đô – Trao yêu thương', 'Tên chương trình'],
      ['org', 'Nhà Triệu Đô', 'Đơn vị tổ chức'],
      ['period', '15/11/2026 – 15/01/2027', 'Thời gian (hiển thị)'],
      ['beneficiary', 'Nhà Triệu Đô giàu tình thương, càng chung tay càng đầy! Quỹ cùng góp sức hỗ trợ các bé có hoàn cảnh khó khăn mắc tim bẩm sinh tại Bệnh viện E. 💗', 'Mô tả dưới tên chương trình'],
      ['goal', 30000000, 'Mục tiêu quỹ (đ)'],
      ['bankName', 'Ngân hàng TMCP Bắc Á (Bac A Bank)', 'Tên ngân hàng (hiển thị)'],
      ['bankBin', '970409', 'Mã BIN ngân hàng cho VietQR (Bac A Bank = 970409). Để trống nếu không muốn hiện QR'],
      ['account', '051001060000377', 'Số tài khoản nhận'],
      ['holder', 'HOANG THI THUY LA', 'Chủ tài khoản (viết hoa không dấu)'],
      ['payPrefix', 'TETTRIEUDO', 'Tiền tố nội dung chuyển khoản'],
    ];
    sheet_('campaign').getRange(2, 1, rows.length, 3).setValues(rows);
  }

  if (sheet_('teams').getLastRow() < 2) {
    sheet_('teams').getRange(2, 1, 1, 1).setValues([['Văn phòng triệu đô']]);
  }

  if (sheet_('offers').getLastRow() < 2) {
    const o = [
      ['p1', 'product', 'Hộp quà Tết "Lộc Xuân"', 'Trà Shan tuyết, mứt gừng, hạt điều rang muối, thiệp viết tay.', 450000, 'hộp', 'Ban Tết Marketing', 120, ''],
      ['p2', 'product', 'Bánh chưng xanh gói tay', 'Cặp bánh 1 kg do các anh chị Vận hành tự gói, giao trước 25 Tết.', 220000, 'cặp', 'Team Vận hành', 200, ''],
      ['p3', 'product', 'Câu đối thư pháp', 'Thư pháp viết theo tên người nhận, khổ 30x90 cm.', 300000, 'bộ', 'CLB Thư pháp', 40, ''],
      ['p4', 'product', 'Túi canvas "Trao yêu thương"', 'Thiết kế độc quyền của team Creative, in lụa thủ công.', 150000, 'chiếc', 'Team Creative', 300, ''],
      ['s1', 'service', 'Chụp ảnh gia đình ngày Tết', 'Buổi chụp 60 phút tại văn phòng hoặc công viên, 20 ảnh chỉnh sửa.', 800000, 'buổi', 'CLB Nhiếp ảnh', 25, ''],
      ['s2', 'service', 'Viết CV và luyện phỏng vấn 1:1', '45 phút cùng chuyên viên Tuyển dụng, phù hợp cho người thân sắp đi làm.', 500000, 'buổi', 'Team Nhân sự', 30, ''],
      ['s3', 'service', 'Tổng vệ sinh nhà đón Tết', 'Nhóm tình nguyện 3 người hỗ trợ dọn dẹp 4 tiếng (nội thành).', 1200000, 'lượt', 'Đội Tình nguyện', 15, ''],
      ['a1', 'activity', 'Chạy bộ gây quỹ "1 km = 10.000đ"', 'Cộng dồn số km cả tuần, công ty đối ứng thêm 50%.', 10000, 'km', 'CLB Chạy bộ', '', '2026-12-13'],
      ['a2', 'activity', 'Hội chợ Tết tại sảnh văn phòng', 'Đăng ký gian hàng hoặc làm tình nguyện viên đứng quầy.', 200000, 'suất', 'Ban Tổ chức', '', '2027-01-10'],
      ['a3', 'activity', 'Gói quà và viết thiệp cho các em', 'Buổi chiều thứ Bảy, phí tham gia dùng mua vật phẩm.', 100000, 'người', 'Đội Tình nguyện', '', '2027-01-09'],
    ];
    sheet_('offers').getRange(2, 1, o.length, o[0].length).setValues(o);
  }

  if (sheet_('lots').getLastRow() < 2) {
    const l = [
      ['l1', 'Tranh sơn mài "Xuân về bản"', 'Khổ 60x80 cm, họa sĩ khách mời tặng chương trình.', 'Họa sĩ Lê Minh', 5000000, 500000, '2027-01-14T20:00:00+07:00'],
      ['l2', 'Bữa trưa cùng Ban Điều hành', 'Bữa trưa 90 phút cho 2 người, trò chuyện về nghề nghiệp.', 'Ban Điều hành', 3000000, 300000, '2027-01-13T17:00:00+07:00'],
      ['l3', 'Áo đấu có chữ ký đội tuyển', 'Áo thi đấu có chữ ký, kèm giấy chứng nhận.', 'Đối tác tài trợ', 4000000, 500000, '2027-01-15T12:00:00+07:00'],
      ['l4', 'Thêm 1 ngày phép năm 2027', 'Một ngày nghỉ phép có lương, áp dụng cho nhân viên chính thức.', 'Phòng Nhân sự', 1000000, 100000, '2027-01-15T17:00:00+07:00'],
    ];
    sheet_('lots').getRange(2, 1, l.length, l[0].length).setValues(l);
  }

  const props = PropertiesService.getScriptProperties();
  let pw = props.getProperty('ADMIN_PASSWORD');
  if (!pw) {
    pw = Utilities.getUuid().slice(0, 8);
    props.setProperty('ADMIN_PASSWORD', pw);
  }
  Logger.log('Cài đặt xong. Mật khẩu quản trị: ' + pw);
  return pw;
}

/* ------------------------------------------------------------------ */
/* Public API (gọi từ trang web)                                        */
/* ------------------------------------------------------------------ */

function getState() {
  return buildState_(false);
}

function addContribution(input) {
  return withLock_(function () {
    input = input || {};
    const who = cleanWho_(input);
    const type = String(input.type || '');
    let refId = '', refName = '', qty = 1, amount = 0;

    if (KINDS.indexOf(type) >= 0) {
      const offer = readTable_('offers').filter(function (o) { return o.id === input.refId && o.kind === type; })[0];
      if (!offer) throw new Error('Mục này không còn trong danh mục. Vui lòng tải lại trang.');
      qty = Math.floor(Number(input.qty));
      if (!(qty >= 1 && qty <= 1000)) throw new Error('Số lượng không hợp lệ.');
      if (offer.stock !== null) {
        const sold = soldMap_()[offer.id] || 0;
        const left = offer.stock - sold;
        if (qty > left) throw new Error(left > 0 ? 'Chỉ còn ' + left + ' ' + offer.unit + '.' : 'Mục này đã hết.');
      }
      amount = offer.price > 0 ? qty * offer.price : Math.floor(Number(input.amount));
      if (offer.price === 0 && !(amount >= MIN_DONATION)) throw new Error('Số tiền tối thiểu là 10.000 đ.');
      refId = offer.id; refName = offer.name;
    } else if (type === 'donation') {
      amount = Math.floor(Number(input.amount));
      if (!(amount >= MIN_DONATION)) throw new Error('Số tiền ủng hộ tối thiểu là 10.000 đ.');
      refName = 'Ủng hộ trực tiếp';
    } else {
      throw new Error('Hình thức đóng góp không hợp lệ.');
    }
    if (amount > MAX_AMOUNT) throw new Error('Số tiền vượt giới hạn cho phép.');

    const c = newContribution_(who, type, refId, refName, qty, amount, clean_(input.note, 300));
    append_('contribs', c);
    const out = publicContribution_(c);
    out.undoToken = c.undoToken;
    return out;
  });
}

/** Người vừa đăng ký được hoàn tác trong 15 phút, khi chưa được xác nhận thanh toán. */
function undoContribution(id, token) {
  return withLock_(function () {
    const c = readTable_('contribs').filter(function (x) { return x.id === id; })[0];
    if (!c) return true;
    if (!token || c.undoToken !== token) throw new Error('Không thể hoàn tác khoản này.');
    if (c.paid) throw new Error('Khoản này đã được xác nhận, vui lòng liên hệ Ban tổ chức.');
    if (Date.now() - new Date(c.createdAt).getTime() > UNDO_MINUTES * 60000) throw new Error('Đã quá thời gian hoàn tác, vui lòng liên hệ Ban tổ chức.');
    delete_('contribs', id);
    return true;
  });
}

function placeBid(input) {
  return withLock_(function () {
    input = input || {};
    const who = cleanWho_(input);
    const lot = readTable_('lots').filter(function (l) { return l.id === input.lotId; })[0];
    if (!lot) throw new Error('Vật phẩm không còn. Vui lòng tải lại trang.');
    if (Date.now() >= new Date(lot.endsAt).getTime()) throw new Error('Phiên đấu giá đã kết thúc.');
    const top = topBid_(lot.id);
    const min = top ? top.amount + lot.step : lot.startPrice;
    const amount = Math.floor(Number(input.amount));
    if (!(amount >= min)) throw new Error(top ? 'Đã có người trả cao hơn. Mức tối thiểu hiện tại là ' + vnd_(min) + '.' : 'Giá khởi điểm là ' + vnd_(min) + '.');
    if (amount > MAX_AMOUNT) throw new Error('Số tiền vượt giới hạn cho phép.');
    const b = { id: uid_(), createdAt: new Date().toISOString(), lotId: lot.id, name: who.name, team: who.team, contact: who.contact, amount: amount };
    append_('bids', b);
    return publicBid_(b);
  });
}

/* ------------------------------------------------------------------ */
/* Admin API (cần mật khẩu)                                             */
/* ------------------------------------------------------------------ */

function adminLogin(pw) {
  checkAdmin_(pw);
  return { sheetUrl: SpreadsheetApp.getActive().getUrl() };
}

function adminGetState(pw) {
  checkAdmin_(pw);
  return buildState_(true);
}

function adminSetPaid(pw, adminName, id, paid) {
  checkAdmin_(pw);
  return withLock_(function () {
    const c = readTable_('contribs').filter(function (x) { return x.id === id; })[0];
    if (!c) throw new Error('Không tìm thấy khoản đóng góp.');
    const prev = Object.assign({}, c);
    c.paid = !!paid;
    c.paidAt = paid ? new Date().toISOString() : '';
    c.paidBy = paid ? (clean_(adminName, 60) || 'Ban tổ chức') : '';
    update_('contribs', id, c);
    return { prev: prev, next: c };
  });
}

function adminDeleteContribution(pw, id) {
  checkAdmin_(pw);
  return withLock_(function () {
    const c = readTable_('contribs').filter(function (x) { return x.id === id; })[0];
    if (!c) throw new Error('Không tìm thấy khoản đóng góp.');
    delete_('contribs', id);
    return c;
  });
}

function adminRestoreContribution(pw, c) {
  checkAdmin_(pw);
  return withLock_(function () {
    if (findRow_('contribs', c.id) > 0) update_('contribs', c.id, c); else append_('contribs', c);
    return true;
  });
}

function adminRecordWin(pw, lotId) {
  checkAdmin_(pw);
  return withLock_(function () {
    const lot = readTable_('lots').filter(function (l) { return l.id === lotId; })[0];
    if (!lot) throw new Error('Không tìm thấy vật phẩm.');
    if (Date.now() < new Date(lot.endsAt).getTime()) throw new Error('Phiên đấu giá chưa kết thúc.');
    const exists = readTable_('contribs').some(function (c) { return c.type === 'auction' && c.refId === lotId; });
    if (exists) throw new Error('Đã ghi nhận khoản thanh toán cho vật phẩm này.');
    const top = topBid_(lotId);
    if (!top) throw new Error('Phiên không có người trả giá.');
    const c = newContribution_({ name: top.name, team: top.team, contact: top.contact }, 'auction', lot.id, 'Đấu giá: ' + lot.name, 1, top.amount, '');
    append_('contribs', c);
    return c;
  });
}

function adminSaveCampaign(pw, camp, teams) {
  checkAdmin_(pw);
  return withLock_(function () {
    const prev = { campaign: readCampaign_(), teams: readTable_('teams').map(function (t) { return t.team; }) };
    const s = sheet_('campaign');
    const keys = ['title', 'org', 'period', 'beneficiary', 'goal', 'bankName', 'bankBin', 'account', 'holder', 'payPrefix'];
    const vals = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 2).getValues() : [];
    keys.forEach(function (k) {
      if (camp[k] === undefined) return;
      let v = k === 'goal' ? Math.max(0, Math.floor(Number(camp[k]) || 0)) : clean_(camp[k], 600);
      if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
      let found = false;
      for (let i = 0; i < vals.length; i++) {
        if (vals[i][0] === k) { s.getRange(i + 2, 2).setValue(v); found = true; break; }
      }
      if (!found) s.appendRow([k, v, '']);
    });
    if (Array.isArray(teams)) {
      const list = teams.map(function (t) { return clean_(t, 80); }).filter(Boolean);
      const ts = sheet_('teams');
      if (ts.getLastRow() > 1) ts.getRange(2, 1, ts.getLastRow() - 1, 1).clearContent();
      if (list.length) ts.getRange(2, 1, list.length, 1).setValues(list.map(function (t) { return [t]; }));
    }
    return prev;
  });
}

/** table: 'offers' | 'lots'. Trả về bản cũ (null nếu thêm mới) để hoàn tác. */
function adminUpsert(pw, table, obj) {
  checkAdmin_(pw);
  if (table !== 'offers' && table !== 'lots') throw new Error('Bảng không hợp lệ.');
  return withLock_(function () {
    const o = table === 'offers' ? cleanOffer_(obj) : cleanLot_(obj);
    const prev = readTable_(table).filter(function (x) { return x.id === o.id; })[0] || null;
    if (prev) update_(table, o.id, o); else append_(table, o);
    return { prev: prev, next: o };
  });
}

function adminDelete(pw, table, id) {
  checkAdmin_(pw);
  if (table !== 'offers' && table !== 'lots') throw new Error('Bảng không hợp lệ.');
  return withLock_(function () {
    const prev = readTable_(table).filter(function (x) { return x.id === id; })[0];
    if (!prev) throw new Error('Không tìm thấy mục cần xóa.');
    delete_(table, id);
    return prev;
  });
}

/* ------------------------------------------------------------------ */
/* Helpers                                                              */
/* ------------------------------------------------------------------ */

function buildState_(full) {
  const contribs = readTable_('contribs');
  const bids = readTable_('bids');
  return {
    campaign: readCampaign_(),
    teams: readTable_('teams').map(function (t) { return t.team; }).filter(Boolean),
    offers: readTable_('offers'),
    lots: readTable_('lots'),
    contributions: full ? contribs.map(function (c) { const o = Object.assign({}, c); delete o.undoToken; return o; }) : contribs.map(publicContribution_),
    bids: full ? bids : bids.map(publicBid_),
    serverTime: Date.now(),
  };
}

function publicContribution_(c) {
  return { id: c.id, createdAt: c.createdAt, name: c.name, team: c.team, type: c.type, refId: c.refId, refName: c.refName, qty: c.qty, amount: c.amount, note: c.note, paid: c.paid, paidAt: c.paidAt, payCode: c.payCode };
}
function publicBid_(b) {
  return { id: b.id, createdAt: b.createdAt, lotId: b.lotId, name: b.name, team: b.team, amount: b.amount };
}

function newContribution_(who, type, refId, refName, qty, amount, note) {
  const id = uid_();
  const prefix = (readCampaign_().payPrefix || 'TETTRIEUDO');
  return {
    id: id, createdAt: new Date().toISOString(), name: who.name, team: who.team, contact: who.contact,
    type: type, refId: refId, refName: refName, qty: qty, amount: amount, note: note,
    paid: false, paidAt: '', paidBy: '',
    payCode: (plain_(prefix) + ' ' + plain_(who.name).slice(0, 20) + ' ' + id.slice(-6).toUpperCase()).trim(),
    undoToken: Utilities.getUuid(),
  };
}

function cleanWho_(input) {
  const name = clean_(input.name, 80);
  const team = clean_(input.team, 80);
  const contact = clean_(input.contact, 100);
  if (!name) throw new Error('Vui lòng nhập họ và tên.');
  if (!contact) throw new Error('Vui lòng nhập email hoặc số điện thoại để Ban tổ chức liên hệ.');
  const teams = readTable_('teams').map(function (t) { return t.team; }).filter(Boolean);
  if (teams.length && teams.indexOf(team) < 0) throw new Error('Phòng ban không hợp lệ. Vui lòng tải lại trang.');
  return { name: name, team: team, contact: contact };
}

function cleanOffer_(o) {
  const kind = KINDS.indexOf(o.kind) >= 0 ? o.kind : 'product';
  const name = clean_(o.name, 150);
  if (!name) throw new Error('Vui lòng nhập tên.');
  const stock = o.stock === '' || o.stock === null || o.stock === undefined ? '' : Math.max(0, Math.floor(Number(o.stock) || 0));
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(o.date || '')) ? String(o.date) : '';
  return {
    id: clean_(o.id, 40) || (kind[0] + uid_()), kind: kind, name: name, desc: clean_(o.desc, 600),
    price: Math.max(0, Math.floor(Number(o.price) || 0)), unit: clean_(o.unit, 30) || 'suất',
    owner: clean_(o.owner, 100), stock: kind === 'activity' && stock === '' ? '' : stock, date: date,
  };
}

function cleanLot_(l) {
  const name = clean_(l.name, 150);
  if (!name) throw new Error('Vui lòng nhập tên vật phẩm.');
  const endsAt = String(l.endsAt || '');
  if (isNaN(new Date(endsAt).getTime())) throw new Error('Thời điểm kết thúc không hợp lệ.');
  return {
    id: clean_(l.id, 40) || ('l' + uid_()), name: name, desc: clean_(l.desc, 600), donor: clean_(l.donor, 100),
    startPrice: Math.max(0, Math.floor(Number(l.startPrice) || 0)), step: Math.max(1000, Math.floor(Number(l.step) || 1000)), endsAt: endsAt,
  };
}

function soldMap_() {
  const m = {};
  readTable_('contribs').forEach(function (c) { if (c.refId && c.type !== 'auction') m[c.refId] = (m[c.refId] || 0) + (c.qty || 0); });
  return m;
}

function topBid_(lotId) {
  return readTable_('bids').filter(function (b) { return b.lotId === lotId; }).sort(function (a, b) { return b.amount - a.amount; })[0] || null;
}

function readCampaign_() {
  const s = sheet_('campaign');
  const out = {};
  if (s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0]) out[String(r[0])] = r[1] instanceof Date ? r[1].toISOString() : String(r[1]);
    });
  }
  out.goal = Number(out.goal) || 0;
  return out;
}

function sheet_(k) {
  const s = SpreadsheetApp.getActive().getSheetByName(T[k].name);
  if (!s) throw new Error('Chưa cài đặt dữ liệu. Ban tổ chức cần chạy hàm setup.');
  return s;
}

function readTable_(k) {
  const t = T[k], s = sheet_(k);
  const last = s.getLastRow();
  if (last < 2) return [];
  const vals = s.getRange(2, 1, last - 1, t.head.length).getValues();
  return vals.filter(function (r) { return r[0] !== '' && r[0] !== null; }).map(function (r) {
    const o = {};
    t.head.forEach(function (h, i) { o[h] = norm_(k, h, r[i]); });
    return o;
  });
}

function norm_(k, key, v) {
  const t = T[k];
  if (v instanceof Date) {
    return key === 'date' ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  if ((t.num || []).indexOf(key) >= 0) return v === '' || v === null ? null : Number(v);
  if ((t.bool || []).indexOf(key) >= 0) return v === true || String(v).toUpperCase() === 'TRUE';
  return v === null || v === undefined ? '' : String(v);
}

function rowOf_(k, o) {
  return T[k].head.map(function (h) {
    const v = o[h];
    if (v === undefined || v === null) return '';
    // Chặn công thức (=, +, -, @) bị Sheets hiểu nhầm khi ghi chuỗi do người dùng nhập
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
    return v;
  });
}

function append_(k, o) { sheet_(k).appendRow(rowOf_(k, o)); }

function findRow_(k, id) {
  const s = sheet_(k);
  const last = s.getLastRow();
  if (last < 2) return -1;
  const ids = s.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function update_(k, id, o) {
  const r = findRow_(k, id);
  if (r < 0) throw new Error('Không tìm thấy dữ liệu cần cập nhật.');
  sheet_(k).getRange(r, 1, 1, T[k].head.length).setValues([rowOf_(k, o)]);
}

function delete_(k, id) {
  const r = findRow_(k, id);
  if (r > 0) sheet_(k).deleteRow(r);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Hệ thống đang bận, vui lòng thử lại sau vài giây.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function checkAdmin_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real || String(pw || '') !== real) throw new Error('Sai mật khẩu quản trị.');
}

/** Cắt độ dài, bỏ ký tự điều khiển. */
function clean_(s, max) {
  return String(s === undefined || s === null ? '' : s).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max || 200);
}

function plain_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .replace(/[^A-Za-z0-9 ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function uid_() {
  return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
}

function vnd_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ' đ';
}
