/**
 * Komei Hotel — Direct Booking Backend (Google Apps Script)
 * ----------------------------------------------------------
 * 配置: Google Apps Script プロジェクト (スプレッドシート紐付け推奨)
 *
 * 提供する機能:
 *   1. doPost  - 仮予約受付 / 本登録 / 決済初期化 (フロントの fetch から呼ぶ)
 *   2. doGet   - 予約情報取得 / 承認リンク処理 (管理者がメールから踏むリンク)
 *   3. メール通知 - 管理者宛の承認依頼、ゲスト宛の承認/確定通知
 *   4. Stripe Checkout Session 作成 (REST API)
 *
 * 必要な Script Properties (ファイル > プロジェクトのプロパティ > スクリプトのプロパティ):
 *   SHEET_ID            … 予約管理スプレッドシートID
 *   ADMIN_EMAIL         … 管理者の通知先 (例: komei.hotel@gmail.com)
 *   FROM_NAME           … 送信者表示名 (例: Komei Hotel)
 *   SITE_BASE_URL       … 公開サイトのベースURL (例: https://yoshinarcorp.github.io/komei)
 *   STRIPE_SECRET_KEY   … Stripe の sk_live_... または sk_test_...
 *   STRIPE_SUCCESS_PATH … /thanks.html (任意)
 *   DRIVE_FOLDER_ID     … パスポート画像保存用フォルダID
 *
 * シート構成 (1シート = 1テーブル):
 *   reservations: id, status, created_at, updated_at, checkin, checkout, nights,
 *                 adults, children, rep_name, rep_email, rep_phone, rep_country,
 *                 estimated_total, final_total, payment_method, payment_status,
 *                 stripe_session_id, token, notes, source, user_agent
 *   guests:       reservation_id, idx, name, nationality, address, occupation,
 *                 passport_no, passport_file_url
 *   logs:         ts, reservation_id, action, detail
 */

// ============ Constants ============
const HEADERS_RESERVATIONS = [
  'id','status','created_at','updated_at','checkin','checkout','nights',
  'adults','children','rep_name','rep_email','rep_phone','rep_country',
  'estimated_total','final_total','payment_method','payment_status',
  'stripe_session_id','token','notes','source','user_agent'
];
const HEADERS_GUESTS = [
  'reservation_id','idx','name','nationality','address','occupation',
  'passport_no','passport_file_url'
];
const HEADERS_LOGS = ['ts','reservation_id','action','detail'];
const HEADERS_MESSAGES = ['ts','reservation_id','sender','message'];

const STATUS = {
  REQUESTED:  'requested',     // フロントから仮予約 POST 直後
  APPROVED:   'approved',      // 管理者が承認 → ゲストに本登録URL送付
  REGISTERED: 'registered',    // 本登録完了 → 決済待ち
  PAID:       'paid',          // 決済完了 → 確定
  CANCELLED:  'cancelled',
  REJECTED:   'rejected'
};

// ============ Entry Points ============

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.type) {
      case 'reservation_request': return jsonResponse(handleReservationRequest(body));
      case 'guest_registration':  return jsonResponse(handleGuestRegistration(body));
      case 'payment_init':        return jsonResponse(handlePaymentInit(body));
      case 'mypage_message':      return jsonResponse(handleMyPageMessage(body));
      case 'mypage_change_request': return jsonResponse(handleMyPageChangeRequest(body));
      case 'admin_auth':            return jsonResponse(handleAdminAuth(body));
      case 'admin_list':            return jsonResponse(handleAdminListReservations(body));
      case 'admin_detail':          return jsonResponse(handleAdminGetDetail(body));
      case 'admin_reply':           return jsonResponse(handleAdminReply(body));
      case 'admin_update_status':   return jsonResponse(handleAdminUpdateStatus(body));
      default: return jsonResponse({ ok:false, error:'unknown type' });
    }
  } catch (err) {
    log_(null, 'doPost_error', err.toString());
    return jsonResponse({ ok:false, error: String(err) });
  }
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  try {
    if (action === 'get_reservation') {
      return jsonResponse(handleGetReservation(e.parameter));
    }
    if (action === 'approve') {
      return htmlResponse(handleApprove(e.parameter));
    }
    if (action === 'reject') {
      return htmlResponse(handleReject(e.parameter));
    }
    if (action === 'mypage_auth') {
      return jsonResponse(handleMyPageAuth(e.parameter));
    }
    if (action === 'get_messages') {
      return jsonResponse(handleGetMessages(e.parameter));
    }
    if (action === 'stripe_webhook_test') {
      return jsonResponse({ ok:true, msg:'use webhook endpoint via separate function' });
    }
    // debug_mail endpoint removed (was temporary debug tool)
    return htmlResponse('<h1>Komei Hotel API</h1><p>OK</p>');
  } catch (err) {
    log_(null, 'doGet_error', err.toString());
    return htmlResponse('<h1>Error</h1><pre>'+err+'</pre>');
  }
}

// ============ Handlers ============

function handleReservationRequest(body) {
  const sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);

  const id = generateReservationId_();
  const token = generateToken_();
  const now = new Date().toISOString();
  const nights = nightsBetween_(body.checkin, body.checkout);
  if (nights < 3) return { ok:false, error:'minimum 3 nights' };

  const row = HEADERS_RESERVATIONS.map(h => {
    switch(h) {
      case 'id': return id;
      case 'status': return STATUS.REQUESTED;
      case 'created_at': return now;
      case 'updated_at': return now;
      case 'checkin': return body.checkin;
      case 'checkout': return body.checkout;
      case 'nights': return nights;
      case 'adults': return body.adults || 0;
      case 'children': return body.children || 0;
      case 'rep_name': return body.representative.name;
      case 'rep_email': return body.representative.email;
      case 'rep_phone': return body.representative.phone;
      case 'rep_country': return body.representative.country;
      case 'estimated_total': return body.estimated_total || computeEstimatedTotal_(body.checkin, body.checkout);
      case 'token': return token;
      case 'notes': return body.notes || '';
      case 'source': return body.source || 'lp_direct';
      case 'user_agent': return body.user_agent || '';
      default: return '';
    }
  });
  sh.appendRow(row);
  log_(id, 'requested', JSON.stringify({nights:nights, total:body.estimated_total}));
  notifyAdminPendingApproval_(id, body, nights);
  notifyGuestRequestReceived_(id, body);
  return { ok:true, reservation_id: id };
}

function handleApprove(p) {
  const id = p.id; const adminToken = p.t;
  let stored = getProp_('ADMIN_TOKEN');
  if (!stored) stored = generateAndStoreAdminToken_();
  if (adminToken !== stored) {
    return '<h1>Unauthorized</h1>';
  }
  const r = findReservationRow_(id);
  if (!r) return '<h1>Not found</h1>';
  if (r.row.status !== STATUS.REQUESTED) return '<h1>Already processed</h1><p>status='+r.row.status+'</p>';

  // optional final_total override via query; fallback to nightly-rate calculation if still 0
  let finalTotal = parseInt(p.final_total || r.row.estimated_total || 0);
  if (finalTotal <= 0) {
    finalTotal = computeEstimatedTotal_(r.row.checkin, r.row.checkout);
  }
  updateReservation_(r.rowIndex, { status: STATUS.APPROVED, final_total: finalTotal, updated_at: new Date().toISOString() });
  log_(id, 'approved', 'final_total='+finalTotal);

  notifyGuestApproved_(id, r.row, finalTotal);
  return '<h1>✅ Approved</h1><p>Reservation '+id+' has been approved. Guest notified.</p>';
}

function handleReject(p) {
  const id = p.id; const adminToken = p.t;
  if (adminToken !== getProp_('ADMIN_TOKEN', '')) return '<h1>Unauthorized</h1>';
  const r = findReservationRow_(id);
  if (!r) return '<h1>Not found</h1>';
  updateReservation_(r.rowIndex, { status: STATUS.REJECTED, updated_at: new Date().toISOString() });
  log_(id, 'rejected', '');
  notifyGuestRejected_(id, r.row);
  return '<h1>❌ Rejected</h1><p>Reservation '+id+' rejected.</p>';
}

function handleGetReservation(p) {
  const id = p.id; const token = p.token;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  if (r.row.token !== token) return { ok:false, error:'invalid token' };
  return {
    ok: true,
    reservation: {
      reservation_id: r.row.id,
      status: r.row.status,
      checkin: toYMDSafe_(r.row.checkin),
      checkout: toYMDSafe_(r.row.checkout),
      adults: r.row.adults,
      children: r.row.children,
      representative_name: getRepName_(r.row),
      representative_email: r.row.rep_email,
      representative_phone: r.row.rep_phone,
      estimated_total: r.row.estimated_total,
      final_total: r.row.final_total,
      payment_status: r.row.payment_status
    }
  };
}

function handleGuestRegistration(body) {
  const id = body.reservation_id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  if (r.row.token !== body.token) return { ok:false, error:'invalid token' };
  if (r.row.status !== STATUS.APPROVED) return { ok:false, error:'invalid status: '+r.row.status };

  // Save guests
  const sh = sheet_('guests');
  ensureHeaders_(sh, HEADERS_GUESTS);
  const guests = body.guests || [];
  guests.forEach((g, i) => {
    let passportUrl = '';
    if (g.passport_image_base64) {
      passportUrl = savePassportImage_(id, i, g.name, g.passport_image_base64, g.passport_image_mime || 'image/jpeg');
    }
    sh.appendRow([
      id, i+1, g.name, g.nationality, g.address, g.occupation,
      g.passport_no || '', passportUrl
    ]);
  });

  // Update representative phone/etc if provided
  const updates = { status: STATUS.REGISTERED, updated_at: new Date().toISOString() };
  if (body.rep_phone) updates.rep_phone = body.rep_phone;
  updateReservation_(r.rowIndex, updates);
  log_(id, 'registered', 'guests='+guests.length);

  return { ok:true, reservation_id: id, token: r.row.token };
}

function handlePaymentInit(body) {
  const id = body.reservation_id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  if (r.row.token !== body.token) return { ok:false, error:'invalid token' };
  if (r.row.status !== STATUS.REGISTERED) return { ok:false, error:'invalid status: '+r.row.status };

  const total = parseInt(r.row.final_total || r.row.estimated_total || 0);

  if (body.method === 'stripe') {
    const session = createStripeCheckoutSession_(id, total, r.row.rep_email);
    updateReservation_(r.rowIndex, {
      payment_method: 'stripe',
      payment_status: 'pending',
      stripe_session_id: session.id,
      updated_at: new Date().toISOString()
    });
    log_(id, 'stripe_session_created', session.id);
    return { ok:true, checkout_url: session.url };
  }

  if (body.method === 'bank') {
    updateReservation_(r.rowIndex, {
      payment_method: 'bank',
      payment_status: 'awaiting_transfer',
      updated_at: new Date().toISOString()
    });
    log_(id, 'bank_selected', '');
    notifyAdminBankPending_(id, r.row, total);
    notifyGuestBankInstructions_(id, r.row, total);
    return { ok:true, method:'bank' };
  }

  return { ok:false, error:'unknown method' };
}

// ============ Stripe ============

function createStripeCheckoutSession_(id, amount, email) {
  const key = getProp_('STRIPE_SECRET_KEY');
  if (!key) throw new Error('STRIPE_SECRET_KEY not set');
  const base = getProp_('SITE_BASE_URL');
  const successPath = getProp_('STRIPE_SUCCESS_PATH', '/thanks.html');
  const payload = {
    'mode': 'payment',
    'payment_method_types[0]': 'card',
    'line_items[0][price_data][currency]': 'jpy',
    'line_items[0][price_data][unit_amount]': String(amount),
    'line_items[0][price_data][product_data][name]': 'Komei Hotel - Reservation ' + id,
    'line_items[0][quantity]': '1',
    'customer_email': email,
    'client_reference_id': id,
    'success_url': base + successPath + '?id=' + id,
    'cancel_url': base + '/payment.html?id=' + id
  };
  const res = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + key },
    payload: payload,
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Stripe: ' + json.error.message);
  return json;
}

/**
 * Webhook endpoint - deploy as separate web app or check via Stripe Dashboard.
 * Wire this to a separate doPost-only deployment.
 */
function stripeWebhookHandler(e) {
  // Recommended: use Stripe webhook signature verification.
  const event = JSON.parse(e.postData.contents);
  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    const id = session.client_reference_id;
    const r = findReservationRow_(id);
    if (r) {
      updateReservation_(r.rowIndex, {
        status: STATUS.PAID,
        payment_status: 'paid',
        updated_at: new Date().toISOString()
      });
      log_(id, 'paid', session.id);
      notifyGuestConfirmed_(id, r.row);
      notifyAdminConfirmed_(id, r.row);
    }
  }
  return jsonResponse({ received:true });
}

// ============ Email Notifications ============

function notifyAdminPendingApproval_(id, body, nights) {
  let adminToken = getProp_('ADMIN_TOKEN');
  if (!adminToken) adminToken = generateAndStoreAdminToken_();
  const baseUrl = ScriptApp.getService().getUrl();
  const approveUrl = baseUrl + '?action=approve&id=' + id + '&t=' + adminToken;
  const rejectUrl  = baseUrl + '?action=reject&id='  + id + '&t=' + adminToken;
  const subject = '[Komei Hotel] 新規仮予約 ' + id + ' ' + body.representative.name + ' (' + body.checkin + ' 〜 ' + body.checkout + ')';
  const html = ''
    + '<h2>新規予約申込</h2>'
    + '<table cellpadding="6">'
    + '<tr><td>予約ID</td><td><b>' + id + '</b></td></tr>'
    + '<tr><td>期間</td><td>' + body.checkin + ' 〜 ' + body.checkout + ' (' + nights + '泊)</td></tr>'
    + '<tr><td>人数</td><td>大人' + body.adults + ' / 子' + body.children + '</td></tr>'
    + '<tr><td>代表者</td><td>' + body.representative.name + ' (' + body.representative.country + ')</td></tr>'
    + '<tr><td>連絡先</td><td>' + body.representative.email + ' / ' + body.representative.phone + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '<tr><td>備考</td><td>' + (body.notes || '-') + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:24px">'
    + '<a href="' + approveUrl + '" style="background:#10b981;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;margin-right:12px">&#10004; 承認する</a>'
    + '<a href="' + rejectUrl + '" style="background:#ef4444;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px">&#10060; 却下</a>'
    + '</p>'
    + '<p style="color:#888;font-size:12px">承認時に金額を変更したい場合は承認URLに <code>&final_total=XXXXX</code> を追加してください。</p>';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'), subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestRequestReceived_(id, body) {
  const subject = '[Komei Hotel] お申込みを受付けました / Reservation request received (' + id + ')';
  const html =
    '<p>' + body.representative.name + ' 様</p>'
    + '<p>この度は Komei Hotel 光明荘へのお申込みをいただき、誠にありがとうございます。<br>'
    + '以下の内容でお申込みを承りました。担当者の確認後、24時間以内に承認のご連絡をお送りします。</p>'
    + '<table cellpadding="6"><tr><td>予約ID</td><td>' + id + '</td></tr>'
    + '<tr><td>チェックイン</td><td>' + body.checkin + '</td></tr>'
    + '<tr><td>チェックアウト</td><td>' + body.checkout + '</td></tr>'
    + '<tr><td>人数</td><td>大人' + body.adults + ' / 子' + body.children + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '</table>'
    + '<hr>'
    + '<p>Dear ' + body.representative.name + ',</p>'
    + '<p>Thank you for your reservation request at Komei Hotel. We have received your request and will reply with approval within 24 hours.</p>';
  GmailApp.sendEmail(body.representative.email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestApproved_(id, row, finalTotal) {
  const base = getProp_('SITE_BASE_URL');
  const url = base + '/register.html?id=' + id + '&token=' + row.token;
  const subject = '[Komei Hotel] ご予約が承認されました / Approved (' + id + ')';
  const html =
    '<p>' + row.rep_name + ' 様</p>'
    + '<p>お申込みいただいたご予約 <b>' + id + '</b> が承認されました。<br>'
    + '以下のリンクから宿泊者情報のご登録とお支払いにお進みください（リンクは7日間有効）。</p>'
    + '<p>確定金額: <b>¥' + Number(finalTotal).toLocaleString() + '</b></p>'
    + '<p><a href="' + url + '" style="background:#f59e0b;color:#fff;padding:14px 28px;text-decoration:none;border-radius:8px;display:inline-block">本登録に進む / Continue Registration</a></p>'
    + '<hr>'
    + '<p>Dear ' + row.rep_name + ',</p>'
    + '<p>Your reservation <b>' + id + '</b> has been approved. Total: <b>¥' + Number(finalTotal).toLocaleString() + '</b>. Please complete guest registration and payment via the link above (valid for 7 days).</p>';
  GmailApp.sendEmail(row.rep_email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestRejected_(id, row) {
  GmailApp.sendEmail(row.rep_email,
    '[Komei Hotel] お申込みについて / Regarding your request (' + id + ')',
    '',
    { htmlBody: '<p>誠に恐れ入りますが、ご希望の日程ではご案内が難しい状況です。<br>別日程でのご検討をお願いいたします。</p><hr><p>We are unable to accommodate your requested dates. Please consider alternative dates.</p>',
      name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestBankInstructions_(id, row, total) {
  const html =
    '<p>' + row.rep_name + ' 様</p>'
    + '<p>下記口座へ <b>3営業日以内</b> にお振込ください。<br>振込人名義の前に予約ID「' + id + '」をご記入ください。</p>'
    + '<p>金額: <b>¥' + Number(total).toLocaleString() + '</b></p>'
    + '<p>銀行名: 三井住友銀行 / 支店: 赤坂支店 / 普通 9527788 / 名義: カ）コウケンショウジ</p>'
    + '<p>入金確認後、確定メールをお送りします。</p>';
  GmailApp.sendEmail(row.rep_email, '[Komei Hotel] お振込のご案内 / Bank transfer instructions (' + id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyAdminBankPending_(id, row, total) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 銀行振込待ち ' + id + ' ' + getRepName_(row),
    '',
    { htmlBody: '<p>予約 ' + id + '（' + getRepName_(row) + '）が銀行振込を選択しました。入金確認後、シートで status を paid に更新してください。</p><p>金額: ¥' + Number(total).toLocaleString() + '</p>' });
}

function notifyGuestConfirmed_(id, row) {
  const html =
    '<p>' + row.rep_name + ' 様</p>'
    + '<p>お支払いが完了し、ご予約 <b>' + id + '</b> が確定いたしました。<br>'
    + 'チェックイン日が近づきましたら、入室方法等の詳細をご案内いたします。</p>'
    + '<p>チェックイン: ' + toYMDSafe_(row.checkin) + ' 16:00〜<br>チェックアウト: ' + toYMDSafe_(row.checkout) + ' 〜10:00</p>'
    + '<p style="margin-top:16px"><a href="' + getProp_('SITE_BASE_URL') + '/mypage.html?id=' + id + '&email=' + encodeURIComponent(row.rep_email) + '" style="background:#f59e0b;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;display:inline-block">マイページを確認 / View My Page</a></p>'
    + '<hr><p>Dear ' + row.rep_name + ',<br>Your reservation <b>' + id + '</b> is now confirmed. We will send check-in details closer to your arrival date.</p>';
  GmailApp.sendEmail(row.rep_email, '[Komei Hotel] ご予約確定のお知らせ / Reservation Confirmed (' + id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyAdminConfirmed_(id, row) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 決済完了 ' + id + ' ' + getRepName_(row),
    '',
    { htmlBody: '<p>予約 ' + id + '（' + getRepName_(row) + '）が決済完了し確定しました。</p>' });
}

// ============ Drive (Passport) ============

function savePassportImage_(reservationId, idx, name, base64, mime) {
  const folderId = getProp_('DRIVE_FOLDER_ID');
  const folder = DriveApp.getFolderById(folderId);
  const data = Utilities.base64Decode(base64.replace(/^data:[^;]+;base64,/, ''));
  const blob = Utilities.newBlob(data, mime, reservationId + '_' + (idx+1) + '_' + name + '.' + (mime.split('/')[1] || 'jpg'));
  const file = folder.createFile(blob);
  // limit to viewer-only access by default; do NOT set public
  return file.getUrl();
}

// ============ Data Helpers ============

/** Get representative name with fallback for old schema (rep_first_name/rep_last_name) */
function getRepName_(row) {
  if (row.rep_name) return String(row.rep_name);
  const first = row.rep_first_name || '';
  const last = row.rep_last_name || '';
  const combined = (String(first) + ' ' + String(last)).trim();
  return combined || '(unknown)';
}

// ============ Sheet Helpers ============

function sheet_(name) {
  const ss = SpreadsheetApp.openById(getProp_('SHEET_ID'));
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  }
}

function findReservationRow_(id) {
  const sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const obj = {};
      headers.forEach((h, j) => obj[h] = data[i][j]);
      return { rowIndex: i+1, row: obj };
    }
  }
  return null;
}

function updateReservation_(rowIndex, updates) {
  const sh = sheet_('reservations');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(updates).forEach(k => {
    const col = headers.indexOf(k);
    if (col >= 0) sh.getRange(rowIndex, col+1).setValue(updates[k]);
  });
}

function log_(reservationId, action, detail) {
  const sh = sheet_('logs');
  ensureHeaders_(sh, HEADERS_LOGS);
  sh.appendRow([new Date().toISOString(), reservationId || '', action, detail || '']);
}

// ============ Util ============

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function htmlResponse(html) {
  return HtmlService.createHtmlOutput(html);
}
function getProp_(key, def) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return (v == null || v === '') ? (def != null ? def : '') : v;
}
function generateReservationId_() {
  const d = new Date();
  const ymd = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyyMMdd');
  const rand = Math.floor(Math.random() * 9000 + 1000);
  return 'R' + ymd + rand;
}
function generateToken_() {
  return Utilities.getUuid().replace(/-/g, '');
}
function generateAndStoreAdminToken_() {
  const t = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', t);
  return t;
}
function nightsBetween_(ci, co) {
  const a = new Date(toYMDSafe_(ci) + 'T00:00:00Z');
  const b = new Date(toYMDSafe_(co) + 'T00:00:00Z');
  return Math.round((b - a) / 86400000);
}

/**
 * Safely convert a date value (Date object or string) to 'YYYY-MM-DD'.
 * Sheets may auto-convert stored date strings into Date objects;
 * calling toISOString() on those yields UTC which shifts the day in JST.
 * This function uses Utilities.formatDate to respect Asia/Tokyo timezone.
 */
function toYMDSafe_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = String(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  return s;
}

/**
 * Server-side price computation (fallback when estimated_total is missing).
 * Uses same logic as front-end: nightly rate * nights - 5% discount + cleaning fee.
 */
function computeEstimatedTotal_(checkin, checkout) {
  // Normalise inputs: Sheets may pass Date objects instead of strings
  function toYMD(v) {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    return String(v).slice(0, 10);
  }
  checkin  = toYMD(checkin);
  checkout = toYMD(checkout);
  const DEFAULT_RATE = 40000;
  const CLEANING_FEE = 27000;
  const CLEANING_FEE_YEAREND = 35000;
  const DIRECT_DISCOUNT = 0.05;
  // Same RATES snapshot as front-end
  const RATES = {"2026-04-06":37000,"2026-04-07":37000,"2026-04-09":50000,"2026-04-12":40000,"2026-04-17":40000,"2026-04-18":40000,"2026-04-19":40000,"2026-04-20":40000,"2026-04-21":40000,"2026-04-28":38000,"2026-04-29":38000,"2026-04-30":38000,"2026-05-01":38000,"2026-05-02":38000,"2026-05-03":38000,"2026-05-04":38000,"2026-05-05":38000,"2026-05-06":38000,"2026-05-07":38000,"2026-05-08":38000,"2026-05-09":38000,"2026-05-10":38000,"2026-05-11":38000,"2026-05-12":38000,"2026-05-13":38000,"2026-05-14":48000,"2026-05-15":51000,"2026-05-16":60000,"2026-05-17":47000,"2026-05-18":47000,"2026-05-19":46000,"2026-05-20":45000,"2026-05-21":47000,"2026-05-22":50000,"2026-05-23":59000,"2026-05-24":49000,"2026-05-25":38000,"2026-05-26":38000,"2026-05-27":38000,"2026-06-04":41000,"2026-06-05":44000,"2026-06-06":47000,"2026-06-14":39000,"2026-06-15":38000,"2026-06-16":37000,"2026-06-22":40000,"2026-06-23":40000,"2026-06-24":40000,"2026-06-25":40000,"2026-06-26":40000,"2026-07-04":42000,"2026-07-05":42000,"2026-07-06":42000,"2026-07-07":49000,"2026-07-08":47000,"2026-07-09":49000,"2026-07-10":51000,"2026-07-11":55000,"2026-07-12":50000,"2026-07-13":49000,"2026-07-14":48000,"2026-07-15":46000,"2026-07-16":50000,"2026-07-17":56000,"2026-07-18":61000,"2026-07-19":55000,"2026-07-20":47000,"2026-07-21":48000,"2026-07-22":51000,"2026-07-23":49000,"2026-07-24":53000,"2026-07-25":64000,"2026-07-26":47000,"2026-07-27":45000,"2026-07-28":44000,"2026-07-29":44000,"2026-07-30":46000,"2026-07-31":49000,"2026-08-01":59000,"2026-08-10":54000,"2026-08-11":46000,"2026-08-12":44000,"2026-08-13":44000,"2026-08-14":48000,"2026-08-15":49000,"2026-08-16":43000,"2026-08-17":41000,"2026-08-18":41000,"2026-08-19":40000,"2026-08-20":42000,"2026-09-06":43000,"2026-09-07":39000,"2026-09-08":38000,"2026-09-09":37000,"2026-09-10":39000,"2026-09-11":43000,"2026-09-12":48000,"2026-09-13":43000,"2026-09-14":38000,"2026-09-15":36000,"2026-09-16":47000,"2026-09-17":47000,"2026-09-18":54000,"2026-09-30":49000};

  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function isYearEnd(key) { const md = key.slice(5); return md >= '12-29' || md <= '01-03'; }

  let room = 0, anyYearEnd = false;
  const d = new Date(checkin + 'T00:00:00Z');
  const end = new Date(checkout + 'T00:00:00Z');
  while (d < end) {
    const key = d.getUTCFullYear() + '-' + pad(d.getUTCMonth()+1) + '-' + pad(d.getUTCDate());
    if (isYearEnd(key)) anyYearEnd = true;
    room += (RATES[key] || DEFAULT_RATE);
    d.setUTCDate(d.getUTCDate() + 1);
  }
  const cleaning = anyYearEnd ? CLEANING_FEE_YEAREND : CLEANING_FEE;
  const discount = Math.round(room * DIRECT_DISCOUNT);
  return room - discount + cleaning;
}

// ============ My Page ============

/**
 * Authenticate guest by reservation_id + email.
 * Returns reservation data if matched.
 */
function handleMyPageAuth(p) {
  const id = (p.id || '').trim();
  const email = (p.email || '').trim().toLowerCase();
  if (!id || !email) return { ok:false, error:'not_found' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };
  if (String(r.row.rep_email).toLowerCase() !== email) return { ok:false, error:'not_found' };

  return {
    ok: true,
    reservation: {
      reservation_id: r.row.id,
      status: r.row.status,
      checkin: toYMDSafe_(r.row.checkin),
      checkout: toYMDSafe_(r.row.checkout),
      adults: r.row.adults,
      children: r.row.children,
      representative_name: getRepName_(r.row),
      representative_email: r.row.rep_email,
      representative_phone: r.row.rep_phone,
      estimated_total: r.row.estimated_total,
      final_total: r.row.final_total,
      payment_status: r.row.payment_status
    }
  };
}

/**
 * Get messages for a reservation (authenticated by email).
 */
function handleGetMessages(p) {
  const id = (p.id || '').trim();
  const email = (p.email || '').trim().toLowerCase();
  if (!id || !email) return { ok:false, error:'auth_failed' };

  const r = findReservationRow_(id);
  if (!r || String(r.row.rep_email).toLowerCase() !== email) return { ok:false, error:'auth_failed' };

  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const messages = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => row[h] = data[i][j]);
    if (row.reservation_id == id) {
      messages.push({ timestamp: row.ts, sender: row.sender, message: row.message });
    }
  }
  return { ok:true, messages: messages };
}

/**
 * Guest sends a chat message.
 */
function handleMyPageMessage(body) {
  const id = (body.reservation_id || '').trim();
  const email = (body.email || '').trim().toLowerCase();
  const message = (body.message || '').trim();
  if (!id || !email || !message) return { ok:false, error:'missing_fields' };

  const r = findReservationRow_(id);
  if (!r || String(r.row.rep_email).toLowerCase() !== email) return { ok:false, error:'auth_failed' };

  // Save message
  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  sh.appendRow([new Date().toISOString(), id, 'guest', message]);
  log_(id, 'mypage_message', 'guest: ' + message.substring(0, 100));

  // Notify admin by email
  const subject = '[Komei Hotel] ゲストからメッセージ ' + getRepName_(r.row) + ' (' + id + ')';
  const replyUrl = getProp_('SITE_BASE_URL') + '/mypage.html?id=' + id;
  const html = '<h3>&#128172; ゲストからのメッセージ</h3>'
    + '<table cellpadding="6">'
    + '<tr><td>予約ID</td><td><b>' + id + '</b></td></tr>'
    + '<tr><td>ゲスト名</td><td>' + getRepName_(r.row) + '</td></tr>'
    + '<tr><td>メール</td><td>' + r.row.rep_email + '</td></tr>'
    + '</table>'
    + '<div style="background:#fef3c7;padding:16px;border-radius:8px;margin:16px 0">'
    + '<p style="white-space:pre-wrap">' + message.replace(/</g, '&lt;') + '</p>'
    + '</div>'
    + '<p><b>返信方法:</b> 管理画面のメッセージシートに直接記入するか、GASの <code>sendAdminReply</code> 関数を使用してください。</p>';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'), subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });

  return { ok:true };
}

/**
 * Guest sends a date/guest change request.
 */
function handleMyPageChangeRequest(body) {
  const id = (body.reservation_id || '').trim();
  const email = (body.email || '').trim().toLowerCase();
  const changeType = body.change_type || 'other';
  const detail = (body.detail || '').trim();
  if (!id || !email || !detail) return { ok:false, error:'missing_fields' };

  const r = findReservationRow_(id);
  if (!r || String(r.row.rep_email).toLowerCase() !== email) return { ok:false, error:'auth_failed' };

  // Save as a message too
  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  const msgText = '[変更リクエスト / Change Request: ' + changeType + ']\n' + detail;
  sh.appendRow([new Date().toISOString(), id, 'guest', msgText]);
  log_(id, 'change_request', changeType + ': ' + detail.substring(0, 200));

  // Notify admin
  const typeLabels = { date:'日程変更', guests:'人数変更', other:'その他' };
  const subject = '[Komei Hotel] 変更リクエスト ' + getRepName_(r.row) + ' (' + id + ') - ' + (typeLabels[changeType] || changeType);
  const html = '<h3>&#9999;&#65039; 変更リクエスト</h3>'
    + '<table cellpadding="6">'
    + '<tr><td>予約ID</td><td><b>' + id + '</b></td></tr>'
    + '<tr><td>ゲスト名</td><td>' + getRepName_(r.row) + '</td></tr>'
    + '<tr><td>現在の日程</td><td>' + toYMDSafe_(r.row.checkin) + ' 〜 ' + toYMDSafe_(r.row.checkout) + '</td></tr>'
    + '<tr><td>ステータス</td><td>' + r.row.status + '</td></tr>'
    + '<tr><td>カテゴリ</td><td><b>' + (typeLabels[changeType] || changeType) + '</b></td></tr>'
    + '</table>'
    + '<div style="background:#fef3c7;padding:16px;border-radius:8px;margin:16px 0">'
    + '<p style="white-space:pre-wrap">' + detail.replace(/</g, '&lt;') + '</p>'
    + '</div>'
    + '<p>マイページのメッセージ機能でゲストに直接返信できます。</p>';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'), subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });

  return { ok:true };
}

/**
 * Admin sends a reply to guest via messages sheet.
 * Run from Apps Script editor: sendAdminReply('R20260409XXXX', 'Your message here')
 */
function sendAdminReply(reservationId, message) {
  const r = findReservationRow_(reservationId);
  if (!r) { Logger.log('Reservation not found'); return; }

  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  sh.appendRow([new Date().toISOString(), reservationId, 'host', message]);
  log_(reservationId, 'admin_reply', message.substring(0, 100));

  // Notify guest by email
  const base = getProp_('SITE_BASE_URL');
  const mypageUrl = base + '/mypage.html?id=' + reservationId + '&email=' + encodeURIComponent(r.row.rep_email);
  const subject = '[Komei Hotel] メッセージが届きました / New message (' + reservationId + ')';
  const html =
    '<p>' + getRepName_(r.row) + ' 様</p>'
    + '<p>Komei Hotelからメッセージが届きました。</p>'
    + '<div style="background:#f1f5f9;padding:16px;border-radius:8px;margin:16px 0">'
    + '<p style="white-space:pre-wrap">' + message.replace(/</g, '&lt;') + '</p>'
    + '</div>'
    + '<p><a href="' + mypageUrl + '" style="background:#f59e0b;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;display:inline-block">マイページで返信する / Reply on My Page</a></p>'
    + '<hr>'
    + '<p>Dear ' + getRepName_(r.row) + ',<br>You have a new message from Komei Hotel. Click the button above to view and reply.</p>';
  GmailApp.sendEmail(r.row.rep_email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
  Logger.log('Reply sent to ' + r.row.rep_email);
}

// ============ Admin Dashboard API ============

/**
 * Verify admin token from request.
 */
function verifyAdmin_(token) {
  let stored = getProp_('ADMIN_TOKEN');
  if (!stored) stored = generateAndStoreAdminToken_();
  return token === stored;
}

/**
 * List all reservations (admin only).
 * Supports filtering by status and date range.
 */
function handleAdminListReservations(body) {
  if (!verifyAdmin_(body.admin_token)) return { ok:false, error:'unauthorized' };

  const sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok:true, reservations:[], stats:{} };

  const headers = data[0];
  const reservations = [];
  let totalRevenue = 0, upcoming = 0, pending = 0;
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => obj[h] = data[i][j]);

    // Apply filters
    if (body.status_filter && body.status_filter !== 'all' && obj.status !== body.status_filter) continue;
    if (body.date_from) {
      const ci = toYMDSafe_(obj.checkin);
      if (ci < body.date_from) continue;
    }
    if (body.date_to) {
      const ci = toYMDSafe_(obj.checkin);
      if (ci > body.date_to) continue;
    }

    reservations.push({
      id: obj.id,
      status: obj.status,
      checkin: toYMDSafe_(obj.checkin),
      checkout: toYMDSafe_(obj.checkout),
      nights: obj.nights,
      adults: obj.adults,
      children: obj.children,
      rep_name: getRepName_(obj),
      rep_email: obj.rep_email,
      rep_phone: obj.rep_phone,
      rep_country: obj.rep_country,
      estimated_total: obj.estimated_total,
      final_total: obj.final_total,
      payment_method: obj.payment_method,
      payment_status: obj.payment_status,
      created_at: obj.created_at,
      updated_at: obj.updated_at,
      source: obj.source
    });

    // Stats
    if (obj.status === STATUS.PAID) totalRevenue += parseInt(obj.final_total || obj.estimated_total || 0);
    if (obj.status === STATUS.REQUESTED) pending++;
    const ciDate = new Date(toYMDSafe_(obj.checkin) + 'T00:00:00+09:00');
    if (ciDate >= now && (obj.status === STATUS.PAID || obj.status === STATUS.APPROVED || obj.status === STATUS.REGISTERED)) upcoming++;
  }

  // Sort by created_at descending (newest first)
  reservations.sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));

  return {
    ok: true,
    reservations: reservations,
    stats: {
      total: reservations.length,
      pending: pending,
      upcoming: upcoming,
      total_revenue: totalRevenue
    }
  };
}

/**
 * Get reservation detail with guests and messages (admin only).
 */
function handleAdminGetDetail(body) {
  if (!verifyAdmin_(body.admin_token)) return { ok:false, error:'unauthorized' };

  const r = findReservationRow_(body.reservation_id);
  if (!r) return { ok:false, error:'not_found' };

  // Get guests
  const gsh = sheet_('guests');
  ensureHeaders_(gsh, HEADERS_GUESTS);
  const gdata = gsh.getDataRange().getValues();
  const gheaders = gdata[0];
  const guests = [];
  for (let i = 1; i < gdata.length; i++) {
    const obj = {};
    gheaders.forEach((h, j) => obj[h] = gdata[i][j]);
    if (obj.reservation_id == body.reservation_id) {
      guests.push(obj);
    }
  }

  // Get messages
  const msh = sheet_('messages');
  ensureHeaders_(msh, HEADERS_MESSAGES);
  const mdata = msh.getDataRange().getValues();
  const mheaders = mdata[0];
  const messages = [];
  for (let i = 1; i < mdata.length; i++) {
    const obj = {};
    mheaders.forEach((h, j) => obj[h] = mdata[i][j]);
    if (obj.reservation_id == body.reservation_id) {
      messages.push({ timestamp: obj.ts, sender: obj.sender, message: obj.message });
    }
  }

  // Get logs
  const lsh = sheet_('logs');
  ensureHeaders_(lsh, HEADERS_LOGS);
  const ldata = lsh.getDataRange().getValues();
  const lheaders = ldata[0];
  const logs = [];
  for (let i = 1; i < ldata.length; i++) {
    const obj = {};
    lheaders.forEach((h, j) => obj[h] = ldata[i][j]);
    if (obj.reservation_id == body.reservation_id) {
      logs.push({ timestamp: obj.ts, action: obj.action, detail: obj.detail });
    }
  }

  return {
    ok: true,
    reservation: {
      id: r.row.id,
      status: r.row.status,
      checkin: toYMDSafe_(r.row.checkin),
      checkout: toYMDSafe_(r.row.checkout),
      nights: r.row.nights,
      adults: r.row.adults,
      children: r.row.children,
      rep_name: getRepName_(r.row),
      rep_email: r.row.rep_email,
      rep_phone: r.row.rep_phone,
      rep_country: r.row.rep_country,
      estimated_total: r.row.estimated_total,
      final_total: r.row.final_total,
      payment_method: r.row.payment_method,
      payment_status: r.row.payment_status,
      stripe_session_id: r.row.stripe_session_id,
      notes: r.row.notes,
      source: r.row.source,
      created_at: r.row.created_at,
      updated_at: r.row.updated_at
    },
    guests: guests,
    messages: messages,
    logs: logs
  };
}

/**
 * Admin sends a reply message to guest (from dashboard).
 */
function handleAdminReply(body) {
  if (!verifyAdmin_(body.admin_token)) return { ok:false, error:'unauthorized' };

  const id = (body.reservation_id || '').trim();
  const message = (body.message || '').trim();
  if (!id || !message) return { ok:false, error:'missing_fields' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };

  // Save message
  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  sh.appendRow([new Date().toISOString(), id, 'host', message]);
  log_(id, 'admin_reply', message.substring(0, 100));

  // Notify guest by email
  const base = getProp_('SITE_BASE_URL');
  const mypageUrl = base + '/mypage.html?id=' + id + '&email=' + encodeURIComponent(r.row.rep_email);
  const subject = '[Komei Hotel] メッセージが届きました / New message (' + id + ')';
  const html =
    '<p>' + getRepName_(r.row) + ' 様</p>'
    + '<p>Komei Hotelからメッセージが届きました。</p>'
    + '<div style="background:#f1f5f9;padding:16px;border-radius:8px;margin:16px 0">'
    + '<p style="white-space:pre-wrap">' + message.replace(/</g, '&lt;') + '</p>'
    + '</div>'
    + '<p><a href="' + mypageUrl + '" style="background:#f59e0b;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;display:inline-block">マイページで返信する / Reply on My Page</a></p>'
    + '<hr>'
    + '<p>Dear ' + getRepName_(r.row) + ',<br>You have a new message from Komei Hotel. Click the button above to view and reply.</p>';
  GmailApp.sendEmail(r.row.rep_email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });

  return { ok:true };
}

/**
 * Admin updates reservation status (approve, reject, cancel, mark paid).
 */
function handleAdminUpdateStatus(body) {
  if (!verifyAdmin_(body.admin_token)) return { ok:false, error:'unauthorized' };

  const id = (body.reservation_id || '').trim();
  const newStatus = (body.new_status || '').trim();
  if (!id || !newStatus) return { ok:false, error:'missing_fields' };

  const validStatuses = [STATUS.APPROVED, STATUS.REJECTED, STATUS.CANCELLED, STATUS.PAID];
  if (validStatuses.indexOf(newStatus) === -1) return { ok:false, error:'invalid_status' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };

  const updates = { status: newStatus, updated_at: new Date().toISOString() };

  // If approving, handle final_total
  if (newStatus === STATUS.APPROVED) {
    let finalTotal = parseInt(body.final_total || r.row.estimated_total || 0);
    if (finalTotal <= 0) finalTotal = computeEstimatedTotal_(r.row.checkin, r.row.checkout);
    updates.final_total = finalTotal;
    notifyGuestApproved_(id, r.row, finalTotal);
  }

  // If marking paid via bank transfer
  if (newStatus === STATUS.PAID) {
    updates.payment_status = 'paid';
    if (!r.row.payment_method) updates.payment_method = 'bank';
    notifyGuestConfirmed_(id, r.row);
  }

  // If rejecting
  if (newStatus === STATUS.REJECTED) {
    notifyGuestRejected_(id, r.row);
  }

  updateReservation_(r.rowIndex, updates);
  log_(id, 'admin_' + newStatus, body.note || '');

  return { ok:true, new_status: newStatus };
}

/**
 * Admin auth check (validates token).
 */
function handleAdminAuth(body) {
  if (!verifyAdmin_(body.admin_token)) return { ok:false, error:'unauthorized' };
  return { ok:true };
}

// ============ One-time Setup ============
/**
 * Run this once from the Apps Script editor to initialize sheets and admin token.
 */
function initialize() {
  ensureHeaders_(sheet_('reservations'), HEADERS_RESERVATIONS);
  ensureHeaders_(sheet_('guests'), HEADERS_GUESTS);
  ensureHeaders_(sheet_('logs'), HEADERS_LOGS);
  ensureHeaders_(sheet_('messages'), HEADERS_MESSAGES);
  generateAndStoreAdminToken_();
  Logger.log('Initialized. ADMIN_TOKEN: ' + getProp_('ADMIN_TOKEN'));
}