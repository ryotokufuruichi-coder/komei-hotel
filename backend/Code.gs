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
 *                 adults, children, rep_first_name, rep_last_name, rep_email, rep_phone, rep_country,
 *                 estimated_total, final_total, payment_method, payment_status,
 *                 stripe_session_id, token, notes, source, user_agent
 *   guests:       reservation_id, idx, name, nationality, address, occupation,
 *                 passport_no, passport_file_url
 *   logs:         ts, reservation_id, action, detail
 */

// ============ Constants ============
const HEADERS_RESERVATIONS = [
  'id','status','created_at','updated_at','checkin','checkout','nights',
  'adults','children','rep_first_name','rep_last_name','rep_email','rep_phone','rep_country',
  'estimated_total','final_total','payment_method','payment_status',
  'stripe_session_id','token','notes','source','user_agent'
];
const HEADERS_GUESTS = [
  'reservation_id','idx','name','nationality','address','occupation',
  'passport_no','passport_file_url'
];
const HEADERS_LOGS = ['ts','reservation_id','action','detail'];
const HEADERS_MESSAGES = ['id','reservation_id','sender','message','timestamp','read_by_host'];
const HEADERS_REVIEWS = [
  'id','reservation_id','rep_name','rep_country','overall','cleanliness','accuracy',
  'checkin','communication','location','value','rooms','comment','private_feedback',
  'created_at','published'
];

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

    // Stripe webhook events have dot-separated type like 'checkout.session.completed'
    if (body.type && body.type.indexOf('.') !== -1 && body.data && body.data.object) {
      return stripeWebhookHandler(e);
    }

    switch (body.type) {
      case 'reservation_request':   return jsonResponse(handleReservationRequest(body));
      case 'guest_registration':    return jsonResponse(handleGuestRegistration(body));
      case 'payment_init':          return jsonResponse(handlePaymentInit(body));
      // Admin API
      case 'admin_auth':            return jsonResponse(handleAdminAuth(body));
      case 'admin_list':            return jsonResponse(handleAdminList(body));
      case 'admin_detail':          return jsonResponse(handleAdminDetail(body));
      case 'admin_update_status':   return jsonResponse(handleAdminUpdateStatus(body));
      case 'admin_reply':           return jsonResponse(handleAdminReply(body));
      // Mypage API
      case 'mypage_message':        return jsonResponse(handleMypageMessage(body));
      case 'mypage_change_request': return jsonResponse(handleMypageChangeRequest(body));
      // Review API
      case 'submit_review':         return jsonResponse(handleSubmitReview(body));
      case 'admin_list_reviews':    return jsonResponse(handleAdminListReviews(body));
      case 'admin_toggle_review':   return jsonResponse(handleAdminToggleReview(body));
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
    if (action === 'approve_form') {
      return htmlResponse(handleApproveForm(e.parameter));
    }
    if (action === 'approve') {
      return htmlResponse(handleApprove(e.parameter));
    }
    if (action === 'reject') {
      return htmlResponse(handleReject(e.parameter));
    }
    if (action === 'mypage_auth') {
      return jsonResponse(handleMypageAuth(e.parameter));
    }
    if (action === 'get_messages') {
      return jsonResponse(handleGetMessages(e.parameter));
    }
    if (action === 'public_reviews') {
      return jsonResponse(handlePublicReviews());
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
      case 'rep_first_name': return body.representative.first_name || '';
      case 'rep_last_name': return body.representative.last_name || '';
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

function handleApproveForm(p) {
  const id = p.id; const adminToken = p.t;
  let stored = getProp_('ADMIN_TOKEN');
  if (!stored) stored = generateAndStoreAdminToken_();
  if (adminToken !== stored) return '<h1>Unauthorized</h1>';
  const r = findReservationRow_(id);
  if (!r) return '<h1>Not found</h1>';
  if (r.row.status !== STATUS.REQUESTED) return '<h1>Already processed</h1><p>status='+r.row.status+'</p>';

  const estTotal = parseInt(r.row.estimated_total || 0);
  const guestName = fullName_(r.row);
  const baseUrl = ScriptApp.getService().getUrl();
  const approveAction = baseUrl + '?action=approve&id=' + id + '&t=' + adminToken;
  const rejectAction  = baseUrl + '?action=reject&id='  + id + '&t=' + adminToken;

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>承認 — ' + id + '</title>'
    + '<style>'
    + 'body{font-family:-apple-system,sans-serif;max-width:560px;margin:40px auto;padding:0 16px;color:#1e293b;background:#f8fafc}'
    + '.card{background:#fff;border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:28px;margin-bottom:20px}'
    + 'h1{font-size:22px;margin:0 0 20px}table{width:100%;border-collapse:collapse}td{padding:8px 4px;border-bottom:1px solid #e2e8f0}'
    + 'td:first-child{color:#64748b;width:100px}'
    + '.amount-box{background:#fffbeb;border:2px solid #f59e0b;border-radius:8px;padding:20px;margin:20px 0}'
    + '.amount-box label{font-weight:600;display:block;margin-bottom:8px}'
    + '.amount-box input{width:100%;font-size:24px;font-weight:700;padding:10px 14px;border:1px solid #d1d5db;border-radius:8px;box-sizing:border-box}'
    + '.btn{display:inline-block;padding:14px 32px;border-radius:8px;font-size:16px;font-weight:600;text-decoration:none;border:none;cursor:pointer;margin-right:12px}'
    + '.btn-approve{background:#10b981;color:#fff}.btn-reject{background:#ef4444;color:#fff}'
    + '.btn:hover{opacity:.9}'
    + '</style></head><body>'
    + '<div class="card"><h1>予約承認</h1>'
    + '<table>'
    + '<tr><td>予約ID</td><td><b>' + id + '</b></td></tr>'
    + '<tr><td>代表者</td><td>' + guestName + '</td></tr>'
    + '<tr><td>期間</td><td>' + toYMDSafe_(r.row.checkin) + ' 〜 ' + toYMDSafe_(r.row.checkout) + ' (' + r.row.nights + '泊)</td></tr>'
    + '<tr><td>人数</td><td>大人' + r.row.adults + ' / 子' + r.row.children + '</td></tr>'
    + '<tr><td>メール</td><td>' + r.row.rep_email + '</td></tr>'
    + '<tr><td>備考</td><td>' + (r.row.notes || '-') + '</td></tr>'
    + '</table>'
    + '<div class="amount-box">'
    + '<label>確定金額（税込）</label>'
    + '<input type="number" id="finalTotal" value="' + estTotal + '" min="0" step="1000">'
    + '<p style="color:#92400e;font-size:13px;margin:8px 0 0">概算金額: ¥' + estTotal.toLocaleString() + '　※変更がなければそのまま承認してください</p>'
    + '</div>'
    + '<div style="text-align:center;margin-top:24px">'
    + '<button class="btn btn-approve" onclick="doApprove()">✅ 承認する</button>'
    + '<a href="' + rejectAction + '" class="btn btn-reject">❌ 却下</a>'
    + '</div></div>'
    + '<script>'
    + 'function doApprove(){'
    + '  var t=document.getElementById("finalTotal").value;'
    + '  window.location="' + approveAction + '&final_total="+encodeURIComponent(t);'
    + '}'
    + '</script></body></html>';
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
      representative_first_name: maskName_(r.row.rep_first_name),
      representative_last_name: maskName_(r.row.rep_last_name),
      representative_name: maskName_(r.row.rep_first_name) + ' ' + maskName_(r.row.rep_last_name),
      representative_email: maskEmail_(r.row.rep_email),
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
  const approveFormUrl = baseUrl + '?action=approve_form&id=' + id + '&t=' + adminToken;
  const rejectUrl  = baseUrl + '?action=reject&id='  + id + '&t=' + adminToken;
  const guestName = (body.representative.first_name || '') + ' ' + (body.representative.last_name || '');
  const subject = '[Komei Hotel] 新規仮予約 ' + id + ' (' + body.checkin + ' 〜 ' + body.checkout + ')';
  const html = ''
    + '<h2>新規予約申込</h2>'
    + '<table cellpadding="6">'
    + '<tr><td>予約ID</td><td><b>' + id + '</b></td></tr>'
    + '<tr><td>期間</td><td>' + body.checkin + ' 〜 ' + body.checkout + ' (' + nights + '泊)</td></tr>'
    + '<tr><td>人数</td><td>大人' + body.adults + ' / 子' + body.children + '</td></tr>'
    + '<tr><td>代表者</td><td>' + guestName.trim() + ' (' + body.representative.country + ')</td></tr>'
    + '<tr><td>連絡先</td><td>' + body.representative.email + ' / ' + body.representative.phone + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '<tr><td>備考</td><td>' + (body.notes || '-') + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:24px">'
    + '<a href="' + approveFormUrl + '" style="background:#10b981;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;margin-right:12px">✅ 承認する（金額確認）</a>'
    + '<a href="' + rejectUrl + '" style="background:#ef4444;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px">❌ 却下</a>'
    + '</p>';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'), subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestRequestReceived_(id, body) {
  const guestNameFull = ((body.representative.first_name || '') + ' ' + (body.representative.last_name || '')).trim();
  const subject = '[Komei Hotel] お申込みを受付けました / Reservation request received (' + id + ')';
  const html =
    '<p>' + guestNameFull + ' 様</p>'
    + '<p>この度は Komei Hotel 光明荘へのお申込みをいただき、誠にありがとうございます。<br>'
    + '以下の内容でお申込みを承りました。担当者の確認後、24時間以内に承認のご連絡をお送りします。</p>'
    + '<table cellpadding="6"><tr><td>予約ID</td><td>' + id + '</td></tr>'
    + '<tr><td>チェックイン</td><td>' + body.checkin + '</td></tr>'
    + '<tr><td>チェックアウト</td><td>' + body.checkout + '</td></tr>'
    + '<tr><td>人数</td><td>大人' + body.adults + ' / 子' + body.children + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '</table>'
    + '<hr>'
    + '<p>Dear ' + guestNameFull + ',</p>'
    + '<p>Thank you for your reservation request at Komei Hotel. We have received your request and will reply with approval within 24 hours.</p>';
  GmailApp.sendEmail(body.representative.email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestApproved_(id, row, finalTotal) {
  const base = getProp_('SITE_BASE_URL');
  const url = base + '/register.html?id=' + id + '&token=' + row.token;
  const name = fullName_(row);
  const subject = '[Komei Hotel] ご予約が承認されました / Approved (' + id + ')';
  const html =
    '<p>' + name + ' 様</p>'
    + '<p>お申込みいただいたご予約 <b>' + id + '</b> が承認されました。<br>'
    + '以下のリンクから宿泊者情報のご登録とお支払いにお進みください（リンクは7日間有効）。</p>'
    + '<p>確定金額: <b>¥' + Number(finalTotal).toLocaleString() + '</b></p>'
    + '<p><a href="' + url + '" style="background:#f59e0b;color:#fff;padding:14px 28px;text-decoration:none;border-radius:8px;display:inline-block">本登録に進む / Continue Registration</a></p>'
    + '<hr>'
    + '<p>Dear ' + name + ',</p>'
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
    '<p>' + fullName_(row) + ' 様</p>'
    + '<p>下記口座へ <b>3営業日以内</b> にお振込ください。<br>振込人名義の前に予約ID「' + id + '」をご記入ください。</p>'
    + '<p>金額: <b>¥' + Number(total).toLocaleString() + '</b></p>'
    + '<p>銀行名: 三井住友銀行 / 支店: 赤坂支店 / 普通 9527788 / 名義: カ）コウケンショウジ</p>'
    + '<p>入金確認後、確定メールをお送りします。</p>';
  GmailApp.sendEmail(row.rep_email, '[Komei Hotel] お振込のご案内 / Bank transfer instructions (' + id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyAdminBankPending_(id, row, total) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 銀行振込待ち ' + id,
    '',
    { htmlBody: '<p>予約 ' + id + ' が銀行振込を選択しました。入金確認後、シートで status を paid に更新してください。</p><p>金額: ¥' + Number(total).toLocaleString() + '</p>' });
}

function notifyGuestConfirmed_(id, row) {
  const name = fullName_(row);
  const html =
    '<p>' + name + ' 様</p>'
    + '<p>お支払いが完了し、ご予約 <b>' + id + '</b> が確定いたしました。<br>'
    + 'チェックイン日が近づきましたら、入室方法等の詳細をご案内いたします。</p>'
    + '<p>チェックイン: ' + toYMDSafe_(row.checkin) + ' 16:00〜<br>チェックアウト: ' + toYMDSafe_(row.checkout) + ' 〜10:00</p>'
    + '<hr><p>Dear ' + name + ',<br>Your reservation <b>' + id + '</b> is now confirmed. We will send check-in details closer to your arrival date.</p>';
  GmailApp.sendEmail(row.rep_email, '[Komei Hotel] ご予約確定のお知らせ / Reservation Confirmed (' + id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyAdminConfirmed_(id, row) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 決済完了 ' + id,
    '',
    { htmlBody: '<p>予約 ' + id + ' が決済完了し確定しました。</p>' });
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

/** Build full display name from first + last */
function fullName_(row) {
  const f = String(row.rep_first_name || '').trim();
  const l = String(row.rep_last_name || '').trim();
  if (f && l) return f + ' ' + l;
  return f || l || '';
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
 * Mask a name for privacy: '山田太郎' → '山***', 'John Smith' → 'J*** S***'
 */
function maskName_(name) {
  if (!name) return '';
  const s = String(name).trim();
  // Detect spaces (Western-style name)
  if (s.indexOf(' ') !== -1) {
    return s.split(/\s+/).map(function(w) { return w.charAt(0) + '***'; }).join(' ');
  }
  // CJK or single-token name
  return s.charAt(0) + '***';
}

/**
 * Mask an email for privacy: 'user@example.com' → 'u***@e***.com'
 */
function maskEmail_(email) {
  if (!email) return '';
  const s = String(email).trim();
  const at = s.indexOf('@');
  if (at <= 0) return '***';
  const local = s.substring(0, at);
  const domain = s.substring(at + 1);
  const dot = domain.lastIndexOf('.');
  if (dot <= 0) return local.charAt(0) + '***@***';
  const domainName = domain.substring(0, dot);
  const tld = domain.substring(dot);
  return local.charAt(0) + '***@' + domainName.charAt(0) + '***' + tld;
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
 * Dynamic pricing: base ¥30,000 + ¥5,000 per month ahead, year-end overrides.
 * Max 10% direct-booking discount vs Airbnb.
 */
function computeEstimatedTotal_(checkin, checkout) {
  // Normalise inputs: Sheets may pass Date objects instead of strings
  function toYMD(v) {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    return String(v).slice(0, 10);
  }
  checkin  = toYMD(checkin);
  checkout = toYMD(checkout);

  const CLEANING_FEE = 27000;
  const CLEANING_FEE_YEAREND = 35000;
  const DIRECT_DISCOUNT = 0.10; // max 10% off vs Airbnb

  const YEAREND_RATES = {
    '12-27': 100000, '12-28': 100000, '12-29': 100000,
    '12-30': 110000, '12-31': 110000, '01-01': 108000, '01-02': 105000
  };

  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function isYearEnd(key) { const md = key.slice(5); return YEAREND_RATES[md] !== undefined; }
  function getRate(dateStr) {
    const md = dateStr.slice(5);
    if (YEAREND_RATES[md] !== undefined) return YEAREND_RATES[md];
    const today = new Date();
    const todayYM = today.getFullYear() * 12 + today.getMonth();
    const dateYM = parseInt(dateStr.slice(0, 4)) * 12 + (parseInt(dateStr.slice(5, 7)) - 1);
    const monthDiff = Math.max(0, dateYM - todayYM);
    return 30000 + monthDiff * 5000;
  }

  let room = 0, anyYearEnd = false;
  const d = new Date(checkin + 'T00:00:00Z');
  const end = new Date(checkout + 'T00:00:00Z');
  while (d < end) {
    const key = d.getUTCFullYear() + '-' + pad(d.getUTCMonth()+1) + '-' + pad(d.getUTCDate());
    if (isYearEnd(key)) anyYearEnd = true;
    room += getRate(key);
    d.setUTCDate(d.getUTCDate() + 1);
  }
  const cleaning = anyYearEnd ? CLEANING_FEE_YEAREND : CLEANING_FEE;
  const discount = Math.round(room * DIRECT_DISCOUNT);
  return room - discount + cleaning;
}

// ============ Admin API ============

function verifyAdminToken_(token) {
  if (!token) return false;
  return token === getProp_('ADMIN_TOKEN');
}

function handleAdminAuth(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  return { ok:true };
}

function handleAdminList(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };

  const sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const allRows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach(function(h, j) { obj[h] = data[i][j]; });
    allRows.push(obj);
  }

  const unrepliedMap = buildUnrepliedMap_();

  const statusFilter = body.status_filter || 'all';
  const dateFrom = body.date_from || '';
  const dateTo   = body.date_to   || '';

  const filtered = allRows.filter(function(r) {
    if (statusFilter !== 'all' && r.status !== statusFilter) return false;
    if (dateFrom && toYMDSafe_(r.checkin) < dateFrom) return false;
    if (dateTo   && toYMDSafe_(r.checkin) > dateTo)   return false;
    return true;
  });

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  const reservations = filtered.map(function(r) {
    return {
      id:             r.id,
      status:         r.status,
      checkin:        toYMDSafe_(r.checkin),
      checkout:       toYMDSafe_(r.checkout),
      nights:         r.nights,
      adults:         r.adults,
      children:       r.children,
      rep_name:       fullName_(r),
      rep_email:      r.rep_email,
      estimated_total:r.estimated_total,
      final_total:    r.final_total,
      payment_method: r.payment_method,
      payment_status: r.payment_status,
      source:         r.source,
      unreplied:      unrepliedMap[r.id] || 0
    };
  }).sort(function(a, b) { return b.id.localeCompare(a.id); });

  return { ok:true, reservations:reservations, stats:buildAdminStats_(allRows, unrepliedMap, today) };
}

function buildUnrepliedMap_() {
  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  if (sh.getLastRow() <= 1) return {};
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const si = headers.indexOf('sender'), ri = headers.indexOf('read_by_host'), ii = headers.indexOf('reservation_id');
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][si] === 'guest' && !data[i][ri]) {
      const rid = String(data[i][ii]);
      map[rid] = (map[rid] || 0) + 1;
    }
  }
  return map;
}

function buildAdminStats_(allRows, unrepliedMap, today) {
  let pending = 0, upcoming = 0, revenue = 0, needsReply = 0;
  allRows.forEach(function(r) {
    if (r.status === 'requested') pending++;
    if (['approved','registered','paid'].indexOf(r.status) >= 0 && toYMDSafe_(r.checkin) >= today) upcoming++;
    if (r.status === 'paid') revenue += Number(r.final_total || 0);
    if (unrepliedMap[r.id]) needsReply++;
  });
  return { total:allRows.length, pending:pending, upcoming:upcoming, total_revenue:revenue, needs_reply:needsReply };
}

function handleAdminDetail(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const id = body.reservation_id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };

  // Guests
  const gsh = sheet_('guests');
  ensureHeaders_(gsh, HEADERS_GUESTS);
  const gdata = gsh.getLastRow() > 1 ? gsh.getDataRange().getValues() : [HEADERS_GUESTS];
  const gheaders = gdata[0];
  const guests = [];
  for (let i = 1; i < gdata.length; i++) {
    if (String(gdata[i][0]) === id) {
      const g = {};
      gheaders.forEach(function(h, j) { g[h] = gdata[i][j]; });
      guests.push(g);
    }
  }

  // Messages — mark guest messages as read by host
  const messages = getMessages_(id);
  markMessagesReadByHost_(id);

  // Logs (newest first)
  const lsh = sheet_('logs');
  ensureHeaders_(lsh, HEADERS_LOGS);
  const ldata = lsh.getLastRow() > 1 ? lsh.getDataRange().getValues() : [HEADERS_LOGS];
  const lheaders = ldata[0];
  const logs = [];
  for (let i = 1; i < ldata.length; i++) {
    if (String(ldata[i][1]) === id) {
      const l = {};
      lheaders.forEach(function(h, j) { l[h] = String(ldata[i][j]); });
      l.timestamp = l.ts;
      logs.push(l);
    }
  }
  logs.reverse();

  const reservation = {
    id:             r.row.id,
    status:         r.row.status,
    checkin:        toYMDSafe_(r.row.checkin),
    checkout:       toYMDSafe_(r.row.checkout),
    nights:         r.row.nights,
    adults:         r.row.adults,
    children:       r.row.children,
    rep_name:       fullName_(r.row),
    rep_email:      r.row.rep_email,
    rep_phone:      r.row.rep_phone,
    rep_country:    r.row.rep_country,
    estimated_total:r.row.estimated_total,
    final_total:    r.row.final_total,
    payment_method: r.row.payment_method,
    payment_status: r.row.payment_status,
    source:         r.row.source,
    notes:          r.row.notes,
    created_at:     r.row.created_at instanceof Date ? r.row.created_at.toISOString() : String(r.row.created_at)
  };

  return { ok:true, reservation:reservation, guests:guests, messages:messages, logs:logs };
}

function handleAdminUpdateStatus(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const id = body.reservation_id;
  const newStatus = body.new_status;
  if (['approved','rejected','cancelled','paid'].indexOf(newStatus) < 0) return { ok:false, error:'invalid status' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };

  updateReservation_(r.rowIndex, { status:newStatus, updated_at:new Date().toISOString() });
  log_(id, 'admin_status_change', 'to='+newStatus);

  if (newStatus === 'approved') {
    let finalTotal = parseInt(r.row.final_total || r.row.estimated_total || 0);
    if (finalTotal <= 0) finalTotal = computeEstimatedTotal_(r.row.checkin, r.row.checkout);
    updateReservation_(r.rowIndex, { final_total:finalTotal });
    notifyGuestApproved_(id, r.row, finalTotal);
  } else if (newStatus === 'rejected') {
    notifyGuestRejected_(id, r.row);
  } else if (newStatus === 'paid') {
    const latest = findReservationRow_(id);
    notifyGuestConfirmed_(id, latest ? latest.row : r.row);
    notifyAdminConfirmed_(id, r.row);
  }

  return { ok:true };
}

function handleAdminReply(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const id = body.reservation_id;
  const message = (body.message || '').trim();
  if (!message) return { ok:false, error:'no message' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };

  addMessage_(id, 'host', message);
  log_(id, 'admin_reply', message.substring(0, 100));

  const base       = getProp_('SITE_BASE_URL');
  const mypageUrl  = base + '/mypage.html?id=' + id + '&email=' + encodeURIComponent(r.row.rep_email);
  const name       = fullName_(r.row);
  const html =
    '<p>' + name + ' 様</p>'
    + '<p>Komei Hotel からメッセージが届いています：</p>'
    + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb;margin:12px 0">'
    + message.replace(/\n/g,'<br>') + '</blockquote>'
    + '<p><a href="' + mypageUrl + '" style="background:#f59e0b;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px">マイページで確認 →</a></p>'
    + '<hr><p>Dear ' + name + ', you have a new message from Komei Hotel. '
    + '<a href="' + mypageUrl + '">View on My Page →</a></p>';
  GmailApp.sendEmail(r.row.rep_email,
    '[Komei Hotel] ホストからメッセージ / Message from Host (' + id + ')',
    '', { htmlBody:html, name:getProp_('FROM_NAME','Komei Hotel') });

  return { ok:true };
}

// ============ Mypage API ============

function handleMypageAuth(params) {
  const id    = (params.id    || '').trim();
  const email = (params.email || '').trim().toLowerCase();
  if (!id || !email) return { ok:false, error:'not_found' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };
  if (String(r.row.rep_email).trim().toLowerCase() !== email) return { ok:false, error:'not_found' };

  return {
    ok: true,
    reservation: {
      reservation_id:       r.row.id,
      status:               r.row.status,
      checkin:              toYMDSafe_(r.row.checkin),
      checkout:             toYMDSafe_(r.row.checkout),
      adults:               r.row.adults,
      children:             r.row.children,
      representative_name:  fullName_(r.row),
      representative_email: r.row.rep_email,
      estimated_total:      r.row.estimated_total,
      final_total:          r.row.final_total,
      payment_status:       r.row.payment_status
    }
  };
}

function handleGetMessages(params) {
  const id    = (params.id    || '').trim();
  const email = (params.email || '').trim().toLowerCase();
  if (!id || !email) return { ok:false, error:'not_found' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };
  if (String(r.row.rep_email).trim().toLowerCase() !== email) return { ok:false, error:'not_found' };

  return { ok:true, messages:getMessages_(id) };
}

function handleMypageMessage(body) {
  const id      = (body.reservation_id || '').trim();
  const email   = (body.email          || '').trim().toLowerCase();
  const message = (body.message        || '').trim();
  if (!message) return { ok:false, error:'no message' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };
  if (String(r.row.rep_email).trim().toLowerCase() !== email) return { ok:false, error:'not_found' };

  addMessage_(id, 'guest', message);
  log_(id, 'guest_message', message.substring(0, 100));

  const adminUrl = getProp_('SITE_BASE_URL') + '/admin.html';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] ゲストからメッセージ / Guest Message (' + id + ')',
    '',
    { htmlBody: '<p>予約 <b>' + id + '</b>（' + fullName_(r.row) + '）からメッセージ：</p>'
        + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb">'
        + message.replace(/\n/g,'<br>') + '</blockquote>'
        + '<p><a href="' + adminUrl + '">管理画面で確認 →</a></p>' });

  return { ok:true };
}

function handleMypageChangeRequest(body) {
  const id         = (body.reservation_id || '').trim();
  const email      = (body.email          || '').trim().toLowerCase();
  const changeType = body.change_type || 'other';
  const detail     = (body.detail         || '').trim();
  if (!detail) return { ok:false, error:'no detail' };

  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not_found' };
  if (String(r.row.rep_email).trim().toLowerCase() !== email) return { ok:false, error:'not_found' };

  const typeLabel = ({ date:'日程変更', guests:'人数変更', other:'その他' })[changeType] || changeType;
  const fullMsg   = '[変更リクエスト: ' + typeLabel + ']\n' + detail;
  addMessage_(id, 'guest', fullMsg);
  log_(id, 'change_request', changeType + ': ' + detail.substring(0, 100));

  const adminUrl = getProp_('SITE_BASE_URL') + '/admin.html';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 変更リクエスト / Change Request (' + id + ')',
    '',
    { htmlBody: '<p>予約 <b>' + id + '</b>（' + fullName_(r.row) + '）からの変更リクエスト：</p>'
        + '<p>種類: <b>' + typeLabel + '</b></p>'
        + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb">'
        + detail.replace(/\n/g,'<br>') + '</blockquote>'
        + '<p><a href="' + adminUrl + '">管理画面で確認 →</a></p>' });

  return { ok:true };
}

// ============ Messages Sheet Helpers ============

function addMessage_(reservationId, sender, message) {
  const sh    = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  const msgId = 'M' + new Date().getTime() + ('000' + Math.floor(Math.random() * 1000)).slice(-3);
  sh.appendRow([msgId, reservationId, sender, message, new Date().toISOString(), false]);
}

function getMessages_(reservationId) {
  const sh = sheet_('messages');
  ensureHeaders_(sh, HEADERS_MESSAGES);
  if (sh.getLastRow() <= 1) return [];
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const messages = [];
  for (let i = 1; i < data.length; i++) {
    const m = {};
    headers.forEach(function(h, j) { m[h] = data[i][j]; });
    if (String(m.reservation_id) === String(reservationId)) {
      messages.push({
        id:           m.id,
        sender:       m.sender,
        message:      m.message,
        timestamp:    m.timestamp instanceof Date ? m.timestamp.toISOString() : String(m.timestamp),
        read_by_host: m.read_by_host
      });
    }
  }
  messages.sort(function(a, b) { return a.timestamp.localeCompare(b.timestamp); });
  return messages;
}

function markMessagesReadByHost_(reservationId) {
  const sh = sheet_('messages');
  if (sh.getLastRow() <= 1) return;
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const si  = headers.indexOf('sender');
  const ri  = headers.indexOf('read_by_host');
  const idi = headers.indexOf('reservation_id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idi]) === String(reservationId)
        && data[i][si] === 'guest'
        && !data[i][ri]) {
      sh.getRange(i + 1, ri + 1).setValue(true);
    }
  }
}

// ============ Review Handlers ============

function handleSubmitReview(body) {
  // Verify reservation exists and guest is authorized
  const rsh = sheet_('reservations');
  const rData = rsh.getDataRange().getValues();
  const rHeaders = rData[0];
  const idIdx = rHeaders.indexOf('id');
  const emailIdx = rHeaders.indexOf('rep_email');
  const statusIdx = rHeaders.indexOf('status');
  const coIdx = rHeaders.indexOf('checkout');
  const fnIdx = rHeaders.indexOf('rep_first_name');
  const lnIdx = rHeaders.indexOf('rep_last_name');
  const countryIdx = rHeaders.indexOf('rep_country');

  let found = null;
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][idIdx]) === String(body.reservation_id)
        && String(rData[i][emailIdx]).toLowerCase() === String(body.email).toLowerCase()) {
      found = rData[i];
      break;
    }
  }
  if (!found) return { ok: false, error: 'Reservation not found' };
  if (found[statusIdx] !== 'paid') return { ok: false, error: 'Only completed stays can be reviewed' };

  // Check checkout date has passed
  const coDate = new Date(found[coIdx]);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (coDate > today) return { ok: false, error: 'Review available after checkout' };

  // Check if already reviewed
  const revSh = sheet_('reviews');
  ensureHeaders_(revSh, HEADERS_REVIEWS);
  if (revSh.getLastRow() > 1) {
    const revData = revSh.getDataRange().getValues();
    const revRidIdx = revData[0].indexOf('reservation_id');
    for (let i = 1; i < revData.length; i++) {
      if (String(revData[i][revRidIdx]) === String(body.reservation_id)) {
        return { ok: false, error: 'Already reviewed' };
      }
    }
  }

  // Validate ratings (1-5)
  const categories = ['overall', 'cleanliness', 'accuracy', 'checkin', 'communication', 'location', 'value', 'rooms'];
  for (const cat of categories) {
    const val = Number(body[cat]);
    if (!val || val < 1 || val > 5) return { ok: false, error: 'Invalid rating for ' + cat };
  }

  const repName = (found[lnIdx] + ' ' + found[fnIdx]).trim();
  const reviewId = 'REV-' + Utilities.getUuid().substring(0, 8);
  const now = new Date().toISOString();

  const row = HEADERS_REVIEWS.map(h => {
    switch (h) {
      case 'id': return reviewId;
      case 'reservation_id': return body.reservation_id;
      case 'rep_name': return repName;
      case 'rep_country': return found[countryIdx] || '';
      case 'overall': return Number(body.overall);
      case 'cleanliness': return Number(body.cleanliness);
      case 'accuracy': return Number(body.accuracy);
      case 'checkin': return Number(body.checkin);
      case 'communication': return Number(body.communication);
      case 'location': return Number(body.location);
      case 'value': return Number(body.value);
      case 'rooms': return Number(body.rooms);
      case 'comment': return (body.comment || '').substring(0, 2000);
      case 'private_feedback': return (body.private_feedback || '').substring(0, 2000);
      case 'created_at': return now;
      case 'published': return false;
      default: return '';
    }
  });
  revSh.appendRow(row);
  log_(body.reservation_id, 'review_submitted', 'Overall: ' + body.overall + '/5');

  // Notify admin
  try {
    const adminEmail = getProp_('ADMIN_EMAIL');
    const subject = '【Komei Hotel】新しいレビュー (' + repName + ' ★' + body.overall + ')';
    const html = '<h3>新しいレビューが投稿されました</h3>'
      + '<p><b>予約ID:</b> ' + body.reservation_id + '<br>'
      + '<b>ゲスト:</b> ' + repName + '<br>'
      + '<b>総合評価:</b> ' + '★'.repeat(body.overall) + ' (' + body.overall + '/5)<br>'
      + '<b>コメント:</b><br>' + (body.comment || '(なし)').replace(/\n/g, '<br>') + '</p>'
      + (body.private_feedback ? '<p><b>プライベートフィードバック:</b><br>' + body.private_feedback.replace(/\n/g, '<br>') + '</p>' : '')
      + '<p><a href="' + getProp_('SITE_BASE_URL') + '/admin.html">管理画面で確認</a></p>';
    MailApp.sendEmail({ to: adminEmail, subject: subject, htmlBody: html, name: getProp_('FROM_NAME') || 'Komei Hotel' });
  } catch (e) {
    log_(body.reservation_id, 'review_notify_error', e.toString());
  }

  return { ok: true };
}

function handleAdminListReviews(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok: false, error: 'unauthorized' };

  const sh = sheet_('reviews');
  ensureHeaders_(sh, HEADERS_REVIEWS);
  if (sh.getLastRow() <= 1) return { ok: true, reviews: [], stats: { total: 0, avg: 0, published: 0 } };

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const reviews = [];
  let totalOverall = 0;
  let pubCount = 0;

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = data[i][j]; });
    reviews.push(obj);
    totalOverall += Number(obj.overall) || 0;
    if (obj.published) pubCount++;
  }

  reviews.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

  return {
    ok: true,
    reviews: reviews,
    stats: {
      total: reviews.length,
      avg: reviews.length > 0 ? Math.round((totalOverall / reviews.length) * 10) / 10 : 0,
      published: pubCount
    }
  };
}

function handleAdminToggleReview(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok: false, error: 'unauthorized' };
  if (!body.review_id) return { ok: false, error: 'review_id required' };

  const sh = sheet_('reviews');
  ensureHeaders_(sh, HEADERS_REVIEWS);
  if (sh.getLastRow() <= 1) return { ok: false, error: 'No reviews' };

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const pubIdx = headers.indexOf('published');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(body.review_id)) {
      const newVal = !data[i][pubIdx];
      sh.getRange(i + 1, pubIdx + 1).setValue(newVal);
      log_(data[i][headers.indexOf('reservation_id')], 'review_publish_toggle', newVal ? 'published' : 'unpublished');
      return { ok: true, published: newVal };
    }
  }
  return { ok: false, error: 'Review not found' };
}

function handlePublicReviews() {
  const sh = sheet_('reviews');
  ensureHeaders_(sh, HEADERS_REVIEWS);
  if (sh.getLastRow() <= 1) return { ok: true, reviews: [], avg: {}, count: 0 };

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const reviews = [];
  const sums = { overall: 0, cleanliness: 0, accuracy: 0, checkin: 0, communication: 0, location: 0, value: 0, rooms: 0 };

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = data[i][j]; });
    if (!obj.published) continue;
    // Public: exclude private_feedback
    reviews.push({
      rep_name: obj.rep_name,
      rep_country: obj.rep_country,
      overall: obj.overall,
      cleanliness: obj.cleanliness,
      accuracy: obj.accuracy,
      checkin: obj.checkin,
      communication: obj.communication,
      location: obj.location,
      value: obj.value,
      rooms: obj.rooms,
      comment: obj.comment,
      created_at: obj.created_at
    });
    for (const k in sums) sums[k] += Number(obj[k]) || 0;
  }

  const count = reviews.length;
  const avg = {};
  if (count > 0) {
    for (const k in sums) avg[k] = Math.round((sums[k] / count) * 10) / 10;
  }

  reviews.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

  return { ok: true, reviews: reviews, avg: avg, count: count };
}

// ============ Review Request Email (Daily Trigger) ============

function sendReviewRequestEmails() {
  const sh = sheet_('reservations');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = (h) => headers.indexOf(h);

  const revSh = sheet_('reviews');
  ensureHeaders_(revSh, HEADERS_REVIEWS);
  const reviewedIds = new Set();
  if (revSh.getLastRow() > 1) {
    const revData = revSh.getDataRange().getValues();
    const ridIdx = revData[0].indexOf('reservation_id');
    for (let i = 1; i < revData.length; i++) {
      reviewedIds.add(String(revData[i][ridIdx]));
    }
  }

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const baseUrl = getProp_('SITE_BASE_URL') || '';
  const fromName = getProp_('FROM_NAME') || 'Komei Hotel';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idx('status')] !== 'paid') continue;

    const coStr = String(row[idx('checkout')]).substring(0, 10);
    if (coStr !== yStr) continue;

    const rid = String(row[idx('id')]);
    if (reviewedIds.has(rid)) continue;

    const email = row[idx('rep_email')];
    const name = (row[idx('rep_last_name')] + ' ' + row[idx('rep_first_name')]).trim();
    const token = row[idx('token')];

    const mypageUrl = baseUrl + '/mypage.html?id=' + rid + '&email=' + encodeURIComponent(email);
    const googlePlaceId = getProp_('GOOGLE_PLACE_ID') || '';
    const googleReviewUrl = googlePlaceId
      ? 'https://search.google.com/local/writereview?placeid=' + googlePlaceId
      : 'https://www.google.com/maps/search/Komei+Hotel+光明荘+東駒形';

    const subject = '【Komei Hotel】ご宿泊ありがとうございました — レビューのお願い';
    const html = '<div style="max-width:600px;margin:0 auto;font-family:sans-serif;">'
      + '<h2 style="color:#d97706;">Komei Hotel 光明荘</h2>'
      + '<p>' + name + ' 様</p>'
      + '<p>この度はKomei Hotelにご宿泊いただき、誠にありがとうございました。</p>'
      + '<p>ご滞在はいかがでしたか？今後のサービス向上のため、ぜひレビューをお聞かせください。</p>'
      + '<p style="text-align:center;margin:30px 0;">'
      + '<a href="' + mypageUrl + '" style="display:inline-block;background:#f59e0b;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:16px;">レビューを書く</a>'
      + '</p>'
      + '<p style="text-align:center;margin:20px 0;">'
      + '<a href="' + googleReviewUrl + '" style="display:inline-block;background:#4285f4;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;">📍 Googleにもレビューを書く</a>'
      + '</p>'
      + '<p style="color:#94a3b8;font-size:13px;">マイページにログイン後、「レビュー」タブからご記入いただけます。<br>Googleレビューもいただけると大変嬉しいです。</p>'
      + '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">'
      + '<p style="color:#94a3b8;font-size:12px;">Komei Hotel 光明荘<br>〒130-0005 東京都墨田区東駒形20-5<br>komei.hotel@gmail.com</p>'
      + '</div>';

    try {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: html, name: fromName });
      log_(rid, 'review_request_sent', email);
    } catch (e) {
      log_(rid, 'review_request_error', e.toString());
    }
  }
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
  ensureHeaders_(sheet_('reviews'), HEADERS_REVIEWS);
  generateAndStoreAdminToken_();
  Logger.log('Initialized. ADMIN_TOKEN=' + getProp_('ADMIN_TOKEN'));
}

/**
 * Run once to set up daily review request trigger.
 * GASエディタで手動実行してください。
 */
function setupReviewTrigger() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendReviewRequestEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create daily trigger at 10:00 AM JST
  ScriptApp.newTrigger('sendReviewRequestEmails')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('Review request email trigger set for 10:00 AM daily.');
}

// manualResend wrapper removed (was one-time debug tool)
