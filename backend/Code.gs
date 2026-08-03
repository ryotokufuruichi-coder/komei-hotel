/**
 * Komei Hotel  EDirect Booking Backend (Google Apps Script)
 * ----------------------------------------------------------
 * 配置: Google Apps Script プロジェクト(スプレッドシート紐付け推奨)
 *
 * 提供する機能:
 *   1. doPost  - 仮予約受付 / 本登録 / 決済初期化 (フロントの fetch から呼ぶ)
 *   2. doGet   - 予約情報取得 / 承認リンク処理(管理者メールから踏むリンク)
 *   3. メール通知 - 管理者への承認依頼、ゲスト宛の承認・確定通知
 *   4. Stripe Checkout Session 作成 (REST API)
 *
 * 必要な Script Properties (ファイル > プロジェクトのプロパティ > スクリプトのプロパティ):
 *   SHEET_ID            … 予約管理スプレッドシートID
 *   ADMIN_EMAIL         … 管理者の通知先(例: komei.hotel@gmail.com)
 *   FROM_NAME           … 送信者表示名(例: Komei Hotel)
 *   SITE_BASE_URL       … 公開サイトのベースURL (例: https://yoshinarcorp.github.io/komei)
 *   STRIPE_SECRET_KEY   … Stripe の sk_live_... または sk_test_...
 *   STRIPE_SUCCESS_PATH … /thanks.html (任意)
 *   DRIVE_FOLDER_ID     … パスポート画像保存用フォルダID
 *
 * シート構成 (1シート= 1テーブル):
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
  'estimated_total','final_total','ota_price','payment_method','payment_status',
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
const HEADERS_AUTO_REPLIES = [
  'intent','label','priority','enabled','keywords_ja','keywords_en','reply_ja','reply_en','updated_at'
];

const STATUS = {
  REQUESTED:  'requested',     // フロントから仮予約POST直後
  APPROVED:   'approved',      // 管理者承認→ゲストに本登録URL送付
  REGISTERED: 'registered',    // 本登録完了→決済待ち
  PAID:       'paid',          // 決済完了→確定
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
      case 'passport_upload':       return jsonResponse(handlePassportUpload(body));
      case 'payment_init':          return jsonResponse(handlePaymentInit(body));
      // Admin API
      case 'admin_auth':            return jsonResponse(handleAdminAuth(body));
      case 'admin_list':            return jsonResponse(handleAdminList(body));
      case 'admin_detail':          return jsonResponse(handleAdminDetail(body));
      case 'admin_update_status':   return jsonResponse(handleAdminUpdateStatus(body));
      case 'admin_approve':         return jsonResponse(handleApprovePost(body));
      case 'admin_reject':          return jsonResponse(handleRejectPost(body));
      case 'admin_reply':           return jsonResponse(handleAdminReply(body));
      // Mypage API
      case 'mypage_message':        return jsonResponse(handleMypageMessage(body));
      case 'mypage_change_request': return jsonResponse(handleMypageChangeRequest(body));
      // Review API
      case 'submit_review':         return jsonResponse(handleSubmitReview(body));
      case 'admin_list_reviews':    return jsonResponse(handleAdminListReviews(body));
      case 'admin_toggle_review':   return jsonResponse(handleAdminToggleReview(body));
      // Auto-reply API
      case 'admin_list_auto_replies':   return jsonResponse(handleAdminListAutoReplies(body));
      case 'admin_update_auto_reply':   return jsonResponse(handleAdminUpdateAutoReply(body));
      case 'admin_reset_auto_replies':  return jsonResponse(handleAdminResetAutoReplies(body));
      case 'admin_test_auto_reply':     return jsonResponse(handleAdminTestAutoReply(body));
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

  // 値をフィールド名で用意し、シートの「実際のヘッダー順」に合わせて書き込む（列順ズレ対策）
  const vals = {
    id: id, status: STATUS.REQUESTED, created_at: now, updated_at: now,
    checkin: body.checkin, checkout: body.checkout, nights: nights,
    adults: body.adults || 0, children: body.children || 0,
    rep_first_name: body.representative.first_name || '',
    rep_last_name: body.representative.last_name || '',
    rep_email: body.representative.email, rep_phone: body.representative.phone,
    rep_country: body.representative.country,
    estimated_total: body.estimated_total || computeEstimatedTotal_(body.checkin, body.checkout),
    ota_price: body.ota_price || '', token: token,
    notes: body.notes || '', source: body.source || 'lp_direct',
    user_agent: body.user_agent || ''
  };
  const _headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(_headers.map(function(h){ return vals[h] !== undefined ? vals[h] : ''; }));
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
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>承認 ' + esc_(id) + '</title>'
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
    + '<tr><td>予約ID</td><td><b>' + esc_(id) + '</b></td></tr>'
    + '<tr><td>代表者</td><td>' + esc_(guestName) + '</td></tr>'
    + '<tr><td>期間</td><td>' + toYMDSafe_(r.row.checkin) + ' ～ ' + toYMDSafe_(r.row.checkout) + ' (' + r.row.nights + '泊)</td></tr>'
    + '<tr><td>人数</td><td>大人' + r.row.adults + ' / 子供' + r.row.children + '</td></tr>'
    + '<tr><td>メール</td><td>' + esc_(r.row.rep_email) + '</td></tr>'
    + '<tr><td>備考</td><td>' + esc_(r.row.notes || '-') + '</td></tr>'
    + '</table>'
    + '<div class="amount-box">'
    + '<label>確定金額（税込）</label>'
    + '<input type="number" id="finalTotal" value="' + estTotal + '" min="0" step="1000">'
    + '<p style="color:#92400e;font-size:13px;margin:8px 0 0">概算金額: ¥' + estTotal.toLocaleString() + '　※変更がなければそのまま承認してください</p>'
    + '</div>'
    + '<div style="text-align:center;margin-top:24px">'
    + '<button class="btn btn-approve" onclick="doAction(\'approve\')">✅ 承認する</button>'
    + '<button class="btn btn-reject" onclick="doAction(\'reject\')">❌ 却下</button>'
    + '</div></div>'
    + '<script>'
    + 'function doAction(act){'
    + '  var body={type:"admin_"+act,admin_token:"' + adminToken + '",id:"' + id + '"};'
    + '  if(act==="approve") body.final_total=document.getElementById("finalTotal").value;'
    + '  fetch("' + baseUrl + '",{method:"POST",body:JSON.stringify(body),headers:{"Content-Type":"text/plain;charset=utf-8"},redirect:"follow"})'
    + '  .then(function(r){return r.json()}).then(function(d){'
    + '    if(d.ok) document.body.innerHTML="<h1>"+(act==="approve"?"✅ Approved":"❌ Rejected")+"</h1><p>Reservation ' + id + ' processed.</p>";'
    + '    else document.body.innerHTML="<h1>Error</h1><p>"+(d.error||"Unknown")+"</p>";'
    + '  }).catch(function(e){alert("Error: "+e)});'
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

/** POST-based approve handler (called from approve form via fetch) */
function handleApprovePost(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const id = body.id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  if (r.row.status !== STATUS.REQUESTED) return { ok:false, error:'already processed (status=' + r.row.status + ')' };
  let finalTotal = parseInt(body.final_total || r.row.estimated_total || 0);
  if (finalTotal <= 0) finalTotal = computeEstimatedTotal_(r.row.checkin, r.row.checkout);
  updateReservation_(r.rowIndex, { status: STATUS.APPROVED, final_total: finalTotal, updated_at: new Date().toISOString() });
  log_(id, 'approved', 'final_total=' + finalTotal);
  notifyGuestApproved_(id, r.row, finalTotal);
  return { ok:true };
}

/** POST-based reject handler (called from approve form via fetch) */
function handleRejectPost(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const id = body.id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  updateReservation_(r.rowIndex, { status: STATUS.REJECTED, updated_at: new Date().toISOString() });
  log_(id, 'rejected', '');
  notifyGuestRejected_(id, r.row);
  return { ok:true };
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
      payment_status: r.row.payment_status,
      guests: getGuestSummary_(r.row.id)
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
    // Frontend sends passport_image / passport_image_type; older payloads used
    // passport_image_base64 / passport_image_mime. Read both (tolerant).
    const b64  = g.passport_image_base64 || g.passport_image || '';
    const mime = g.passport_image_mime   || g.passport_image_type || 'image/jpeg';
    let passportUrl = '';
    if (b64) {
      passportUrl = savePassportImage_(id, i, g.name, b64, mime);
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

/**
 * Guest summary for a reservation, used by the passport re-upload page to show
 * which guests still owe a passport. A guest "still owes" when they are foreign
 * (nationality != JP) and passport_file_url is empty. Read-only.
 */
function getGuestSummary_(reservationId) {
  const sh = sheet_('guests');
  ensureHeaders_(sh, HEADERS_GUESTS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const H = data[0];
  const cIdx = H.indexOf('idx'), cName = H.indexOf('name'),
        cNat = H.indexOf('nationality'), cUrl = H.indexOf('passport_file_url');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(reservationId)) continue;
    const url = cUrl >= 0 ? data[i][cUrl] : '';
    out.push({
      idx: data[i][cIdx],
      name: maskName_(data[i][cName]),
      nationality: data[i][cNat],
      has_passport: !!url
    });
  }
  return out;
}

/**
 * Passport re-upload (後日提出): accept a passport image for one guest of an
 * existing reservation, even after registration/payment is complete.
 * Token-gated; rejects only cancelled/rejected reservations.
 * body: { reservation_id, token, guest_idx,
 *         passport_image (base64) [+ passport_image_type], passport_no? }
 */
function handlePassportUpload(body) {
  const id = body.reservation_id;
  const r = findReservationRow_(id);
  if (!r) return { ok:false, error:'not found' };
  if (r.row.token !== body.token) return { ok:false, error:'invalid token' };
  if (r.row.status === STATUS.CANCELLED || r.row.status === STATUS.REJECTED)
    return { ok:false, error:'invalid status: ' + r.row.status };

  const guestIdx = Number(body.guest_idx);
  if (!guestIdx) return { ok:false, error:'guest_idx required' };
  const b64  = body.passport_image_base64 || body.passport_image || '';
  const mime = body.passport_image_mime   || body.passport_image_type || 'image/jpeg';
  if (!b64) return { ok:false, error:'passport image required' };

  const sh = sheet_('guests');
  ensureHeaders_(sh, HEADERS_GUESTS);
  const data = sh.getDataRange().getValues();
  const H = data[0];
  const cIdx = H.indexOf('idx'), cName = H.indexOf('name'),
        cNo = H.indexOf('passport_no'), cUrl = H.indexOf('passport_file_url');
  let rowNum = -1, gname = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id) && Number(data[i][cIdx]) === guestIdx) {
      rowNum = i + 1; gname = data[i][cName]; break;
    }
  }
  if (rowNum < 0) return { ok:false, error:'guest not found' };

  const url = savePassportImage_(id, guestIdx - 1, gname || ('guest' + guestIdx), b64, mime);
  sh.getRange(rowNum, cUrl + 1).setValue(url);
  if (body.passport_no && cNo >= 0) sh.getRange(rowNum, cNo + 1).setValue(body.passport_no);

  log_(id, 'passport_uploaded', 'guest_idx=' + guestIdx);
  try { notifyAdminPassport_(id, guestIdx, gname); } catch (e) {}
  return { ok:true, reservation_id:id, guest_idx:guestIdx };
}

function notifyAdminPassport_(id, guestIdx, gname) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] パスポート後日提出 ' + id,
    '',
    { htmlBody: '<p>予約 ' + esc_(id) + ' の宿泊者 #' + esc_(String(guestIdx)) +
      '（' + esc_(gname || '') + '）がパスポートを後日提出しました。Drive をご確認ください。</p>' });
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
  // Stripe webhook signature verification
  const payload = e.postData.contents;
  const sigHeader = e.parameter['Stripe-Signature'] || (e.headers && e.headers['Stripe-Signature']) || '';
  const whSecret = getProp_('STRIPE_WEBHOOK_SECRET');
  if (whSecret) {
    if (!verifyStripeSignature_(payload, sigHeader, whSecret)) {
      return jsonResponse({ error: 'Invalid signature' }, 403);
    }
  }
  const event = JSON.parse(payload);
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
      // 後日提出パスポートがあれば、決済完了時に提出案内メールを自動送信
      try {
        var _pc = pendingPassportCount_(id);
        if (_pc > 0) {
          sendPassportRequestEmail_(r.row, _pc, false);
          log_(id, 'passport_req_sent', 'on_paid pending=' + _pc);
          // 決済時点で既に過ぎているマイルストーンは抑止（直後の重複リマインド防止）
          var _d = daysUntilCheckin_(r.row.checkin);
          [30, 21, 7, 3].forEach(function (m) { if (_d <= m) log_(id, 'passport_reminder', 'm=' + m); });
        }
      } catch (_e) { log_(id, 'passport_req_error', String(_e)); }
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
  const guestName = esc_((body.representative.first_name || '') + ' ' + (body.representative.last_name || ''));
  const subject = '[Komei Hotel] 新規仮予約 ' + id + ' (' + body.checkin + ' ～ ' + body.checkout + ')';
  const html = ''
    + '<h2>新規予約申込</h2>'
    + '<table cellpadding="6">'
    + '<tr><td>予約ID</td><td><b>' + esc_(id) + '</b></td></tr>'
    + '<tr><td>期間</td><td>' + esc_(body.checkin) + ' ～ ' + esc_(body.checkout) + ' (' + nights + '泊)</td></tr>'
    + '<tr><td>人数</td><td>大人' + parseInt(body.adults) + ' / 子供' + parseInt(body.children) + '</td></tr>'
    + '<tr><td>代表者</td><td>' + guestName.trim() + ' (' + esc_(body.representative.country) + ')</td></tr>'
    + '<tr><td>連絡先</td><td>' + esc_(body.representative.email) + ' / ' + esc_(body.representative.phone) + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '<tr><td>OTA参考価格</td><td>' + (body.ota_price ? '¥' + Number(body.ota_price).toLocaleString() : '未入力') + '</td></tr>'
    + '<tr><td>備考</td><td>' + esc_(body.notes || '-') + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:24px">'
    + '<a href="' + approveFormUrl + '" style="background:#10b981;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;margin-right:12px">✅ 承認する（金額確認）</a>'
    + '<a href="' + rejectUrl + '" style="background:#ef4444;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px">❌ 却下</a>'
    + '</p>';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'), subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestRequestReceived_(id, body) {
  const guestNameFull = esc_(((body.representative.first_name || '') + ' ' + (body.representative.last_name || '')).trim());
  const subject = '[Komei Hotel] お申込みを受付けました / Reservation request received (' + id + ')';
  const html =
    '<p>' + guestNameFull + ' 様</p>'
    + '<p>この度は Komei Hotel 光明荘へのお申込みをいただき、誠にありがとうございます。<br>'
    + '以下の内容でお申込みを承りました。担当者の確認後、24時間以内に承認のご連絡をお送りします。</p>'
    + '<table cellpadding="6"><tr><td>予約ID</td><td>' + esc_(id) + '</td></tr>'
    + '<tr><td>チェックイン</td><td>' + esc_(body.checkin) + '</td></tr>'
    + '<tr><td>チェックアウト</td><td>' + esc_(body.checkout) + '</td></tr>'
    + '<tr><td>人数</td><td>大人' + parseInt(body.adults) + ' / 子供' + parseInt(body.children) + '</td></tr>'
    + '<tr><td>概算金額</td><td>¥' + Number(body.estimated_total).toLocaleString() + '</td></tr>'
    + '</table>'
    + '<hr>'
    + '<p>Dear ' + guestNameFull + ',</p>'
    + '<p>Thank you for your reservation request at Komei Hotel. We have received your request and will reply with approval within 24 hours.</p>';
  GmailApp.sendEmail(body.representative.email, subject, '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyGuestApproved_(id, row, finalTotal) {
  const base = getProp_('SITE_BASE_URL');
  let token = row.token;
  if (!token) {                                     // ★防御: tokenが空なら生成して保存
    token = Utilities.getUuid().replace(/-/g, '');
    const r = findReservationRow_(id);
    if (r) updateReservation_(r.rowIndex, { token: token });
  }
  const url = base + '/register.html?id=' + id + '&token=' + token;
  const name = fullName_(row);
  const subject = '[Komei Hotel] ご予約が承認されました / Approved (' + id + ')';
  const html =
    '<p>' + esc_(name) + ' 様</p>'
    + '<p>お申込みいただいたご予約 <b>' + esc_(id) + '</b> が承認されました。<br>'
    + '以下のリンクから宿泊者情報のご登録とお支払いにお進みください（リンクは7日間有効です）。</p>'
    + '<p>確定金額: <b>¥' + Number(finalTotal).toLocaleString() + '</b></p>'
    + '<p><a href="' + url + '" style="background:#f59e0b;color:#fff;padding:14px 28px;text-decoration:none;border-radius:8px;display:inline-block">本登録に進む / Continue Registration</a></p>'
    + '<hr>'
    + '<p>Dear ' + esc_(name) + ',</p>'
    + '<p>Your reservation <b>' + esc_(id) + '</b> has been approved. Total: <b>¥' + Number(finalTotal).toLocaleString() + '</b>. Please complete guest registration and payment via the link above (valid for 7 days).</p>';
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
    '<p>' + esc_(fullName_(row)) + ' 様</p>'
    + '<p>下記口座へ <b>3営業日以内</b> にお振込ください。<br>振込人名義の前に予約ID「' + esc_(id) + '」をご記入ください。</p>'
    + '<p>金額: <b>¥' + Number(total).toLocaleString() + '</b></p>'
    + '<p>銀行名: 三井住友銀行 / 支店: 赤坂支店 / 普通 9527788 / 名義: カ）コウケンショウジ</p>'
    + '<p>入金確認後、確定メールをお送りします。</p>';
  GmailApp.sendEmail(row.rep_email, '[Komei Hotel] お振込のご案内 / Bank transfer instructions (' + id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
}

function notifyAdminBankPending_(id, row, total) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 銀行振込待ち ' + id,
    '',
    { htmlBody: '<p>予約 ' + esc_(id) + ' が銀行振込を選択しました。入金確認後、シートで status を paid に更新してください。</p><p>金額: ¥' + Number(total).toLocaleString() + '</p>' });
}

function notifyGuestConfirmed_(id, row) {
  // A: 決済確定時のウェルカム（多言語・チェックイン概要＋ハウスルール＋前日ガイド予告）
  try { sendStay_(row, 'welcome'); }
  catch (e) { log_(id, 'welcome_email_error', String(e)); }
}

function notifyAdminConfirmed_(id, row) {
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 決済完了 ' + id,
    '',
    { htmlBody: '<p>予約 ' + esc_(id) + ' が決済完了し確定しました。</p>' });
}

// ============ Passport follow-up (後日提出) ============
// On payment: if any guest deferred their passport, email the guest a passport.html
// link. Then a daily time-trigger (sendPassportReminders) re-sends reminders at
// 30 / 21 / 7 / 3 days before check-in until every pending passport is submitted.

/** Count guests who still owe a passport (foreign nationality + no file saved). */
function pendingPassportCount_(reservationId) {
  var g = getGuestSummary_(reservationId);
  var n = 0;
  g.forEach(function (x) {
    if (x.nationality && String(x.nationality).toUpperCase() !== 'JP' && !x.has_passport) n++;
  });
  return n;
}

/** Pick an email language from the reservation's country field. */
function pickLang_(country) {
  var c = String(country || '').toUpperCase();
  if (c === 'JP' || c === 'JPN' || c.indexOf('JAPAN') >= 0 || c.indexOf('日本') >= 0) return 'ja';
  if (['TW', 'CN', 'HK', 'MO'].indexOf(c) >= 0 || /TAIWAN|CHINA|HONG|MACAU|CHINESE/.test(c)) return 'zh';
  if (c === 'KR' || /KOREA/.test(c)) return 'ko';
  return 'en';
}

/** Localized subject/body/button for the passport request/reminder email. */
function passportI18n_(lang, n, ci, isReminder) {
  var L = {
    ja: { s: (isReminder ? '【リマインド】' : '') + 'パスポートのご提出をお願いします',
      p: 'ご予約（チェックイン ' + ci + '）につきまして、あと <b>' + n + '名</b> のパスポート画像が未提出です。日本の法令によりチェックイン前までのご提出が必須です。下記からアップロードをお願いします（未提出の方のみ表示されます）。',
      b: 'パスポートを提出する' },
    en: { s: (isReminder ? 'Reminder: ' : '') + 'Passport needed before check-in',
      p: 'For your stay (check-in ' + ci + '), <b>' + n + ' guest(s)</b> still need to submit a passport photo, required by Japanese law before check-in. Please upload it below (only the pending guest is shown).',
      b: 'Submit passport' },
    zh: { s: (isReminder ? '【提醒】' : '') + '請於入住前提交護照',
      p: '關於您的預訂（入住 ' + ci + '），尚有 <b>' + n + ' 位</b> 需要提交護照照片。依日本法規，入住前為必須。請透過下方連結上傳（僅顯示尚未提交的成員）。',
      b: '提交護照' },
    ko: { s: (isReminder ? '【알림】' : '') + '체크인 전 여권 제출이 필요합니다',
      p: '예약(체크인 ' + ci + ') 관련하여 <b>' + n + '분</b>의 여권 이미지가 미제출 상태입니다. 일본 법령에 따라 체크인 전까지 제출이 필수입니다. 아래에서 업로드해 주세요(미제출자만 표시됩니다).',
      b: '여권 제출' }
  };
  return L[lang] || L.en;
}

/** Send the passport request/reminder email to the reservation's guest. */
function sendPassportRequestEmail_(row, n, isReminder) {
  var to = row.rep_email;
  if (!to) return false;
  var base = getProp_('SITE_BASE_URL', 'https://komei.yoshinarcorp.com');
  var url = base + '/passport.html?id=' + row.id + '&token=' + row.token;
  var ci = toYMDSafe_(row.checkin);
  var lang = pickLang_(row.rep_country);
  var t = passportI18n_(lang, n, ci, isReminder);
  var en = passportI18n_('en', n, ci, isReminder);
  var btn = '<p><a href="' + url + '" style="background:#f59e0b;color:#fff;padding:14px 28px;text-decoration:none;border-radius:8px;display:inline-block">' + t.b + (lang !== 'en' ? ' / ' + en.b : '') + '</a></p>'
    + '<p style="font-size:12px;color:#666">' + url + '</p>';
  var html = '<p>' + esc_(fullName_(row)) + ' 様 / Dear guest,</p><p>' + t.p + '</p>' + btn;
  if (lang !== 'en') html += '<hr><p>' + en.p + '</p>';
  GmailApp.sendEmail(to, '[Komei Hotel] ' + t.s + ' (' + row.id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
  return true;
}

/** Days from today (JST) until a YYYY-MM-DD date; negative if past. */
function daysUntilDate_(ymd) {
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var a = new Date(todayStr + 'T00:00:00+09:00');
  var b = new Date(ymd + 'T00:00:00+09:00');
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}
/** Days from today (JST) until a check-in date; negative if past. */
function daysUntilCheckin_(ci) { return daysUntilDate_(toYMDSafe_(ci)); }

/** Map reservation_id -> {milestoneDays: true} already sent, from the logs sheet. */
function passportReminderSentMap_() {
  var sh = sheet_('logs');
  ensureHeaders_(sh, HEADERS_LOGS);
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== 'passport_reminder') continue;
    var id = data[i][1];
    var m = String(data[i][3] || '').match(/m=(\d+)/);
    if (!m) continue;
    if (!map[id]) map[id] = {};
    map[id][Number(m[1])] = true;
  }
  return map;
}

/**
 * Daily time-trigger target. For every PAID reservation with a pending passport,
 * send one reminder at each of 30 / 21 / 7 / 3 days before check-in (until submitted).
 */
function sendPassportReminders() {
  var MILESTONES = [30, 21, 7, 3];
  var sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  var data = sh.getDataRange().getValues();
  var H = data[0];
  var iId = H.indexOf('id'), iStatus = H.indexOf('status'), iCheckin = H.indexOf('checkin');
  var sentMap = passportReminderSentMap_();
  var checked = 0, emailed = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iStatus]) !== STATUS.PAID) continue;
    var id = data[i][iId];
    if (!id) continue;
    var ci = data[i][iCheckin];
    if (!ci) continue;
    var d = daysUntilCheckin_(ci);
    if (d < 0) continue;
    var pending = pendingPassportCount_(id);
    if (pending <= 0) continue;
    checked++;
    var sent = sentMap[id] || {};
    var due = MILESTONES.filter(function (m) { return d <= m && !sent[m]; });
    if (!due.length) continue;
    var row = {};
    for (var j = 0; j < H.length; j++) row[H[j]] = data[i][j];
    try {
      if (sendPassportRequestEmail_(row, pending, true)) emailed++;
      due.forEach(function (m) { log_(id, 'passport_reminder', 'm=' + m); });
    } catch (e) { log_(id, 'passport_reminder_error', String(e)); }
  }
  log_(null, 'passport_reminders_run', 'checked=' + checked + ' emailed=' + emailed);
  return { checked: checked, emailed: emailed };
}

/** Run ONCE from the editor: install/refresh the single daily guest-email trigger. */
function installDailyEmails() {
  var names = ['sendPassportReminders', 'runDailyGuestEmails'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyGuestEmails').timeBased().everyDays(1).atHour(10).create();
  return 'installed: runDailyGuestEmails daily ~10:00 JST (passport reminders + stay emails A-F)';
}

/** Daily trigger target: passport reminders (B set) + stay/check-in emails. */
function runDailyGuestEmails() {
  var out = {};
  try { out.passport = sendPassportReminders(); } catch (e) { log_(null, 'passport_reminders_error', String(e)); }
  try { out.stay = sendStayEmails(); } catch (e) { log_(null, 'stay_emails_error', String(e)); }
  return out;
}

// ============ Stay / check-in guide emails (A-F) ============
// A (welcome) is sent on payment by notifyGuestConfirmed_. B-F are date-based,
// sent by sendStayEmails() from the daily trigger. Multilingual (primary + EN).
// Sensitive entry info (smart-lock code, Wi-Fi password) lives ONLY in the
// day-before "checkin_guide" email, behind the HOUSE_MANUAL_URL Drive guide.

var STAY_FACTS_ = {
  addr: '4-20-5 Higashikomagata, Sumida-ku, Tokyo 130-0005',
  map: 'https://maps.app.goo.gl/jVzhawyQTeTFfkLcA',
  tel: '03-6899-5681',
  wifi: 'Komei-Guest',
  transfer: 'https://tokyo-door-to-door.netlify.app',
  tours: 'https://tokyo-experience.web.app'
};

function stayEmail_(kind, lang, ctx) {
  var F = STAY_FACTS_;
  var guide = ctx.guide || '';
  function gbtn(label) {
    return guide ? '<p><a href="' + guide + '" style="background:#f59e0b;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;display:inline-block">' + label + '</a></p><p style="font-size:12px;color:#666">' + guide + '</p>' : '';
  }
  var info = '&#8505;&#65039; ' + F.addr + '<br>Map: ' + F.map + ' &#65295; TEL: ' + F.tel;
  var ci = ctx.ci, co = ctx.co, id = esc_(ctx.id);
  var D = {
    ja: {
      welcome: { s: 'ご予約確定のお知らせ', b:
        '<p>お支払いが完了し、ご予約 <b>' + id + '</b> が確定しました。ありがとうございます！</p>'
        + '<p>チェックイン: ' + ci + ' 16:00<br>チェックアウト: ' + co + ' 10:00（レイトアウト不可）</p>'
        + '<p>入室方法・WiFi等の詳しいチェックインガイドは、<b>チェックイン前日</b>に改めてお送りします。</p>'
        + '<p>[ハウスルール] 全館禁煙（屋上・バルコニー含む）／住宅街のため21:00以降はお静かに・大人数のパーティー不可／ゴミは45L指定袋で分別</p>'
        + '<p>送迎: ' + F.transfer + '<br>体験・ツアー: ' + F.tours + '</p><p>' + info + '</p>' },
      arrival: { s: 'ご到着予定時刻について', b:
        '<p>まもなくチェックインです。準備のため、<b>ご到着予定時刻</b>をお知らせいただけますか？（セルフチェックイン・16:00以降いつでも可）</p>'
        + gbtn('ハウスマニュアルを見る')
        + '<p>チェックイン: ' + ci + ' 16:00 / チェックアウト: ' + co + ' 10:00</p>' },
      checkin_guide: { s: 'チェックインガイド（明日ご入室）', b:
        '<p>明日 <b>' + ci + ' 16:00〜</b> チェックインです。セルフチェックインのため、<b>必ず事前に</b>チェックインガイドをご確認ください（スマートロック暗証番号・WiFi・写真付きアクセス）。</p>'
        + gbtn('チェックインガイドを開く')
        + '<p>住所: ' + F.addr + '<br>Map: ' + F.map + '<br>WiFi SSID: <b>' + F.wifi + '</b>（パスワードはガイド内）</p>'
        + '<p>チェックアウト: ' + co + ' 10:00 / 全館禁煙 / 緊急連絡: ' + F.tel + '</p>' },
      garbage: { s: 'ゴミの分別について', b:
        '<p>おはようございます。ゴミは <b>45L指定袋</b> で分別（燃える／瓶／缶／ペットボトル）をお願いします。</p>'
        + '<p>※未分別の場合、仕分け料 &yen;4,400 を頂戴することがあります。満杯時は屋外ゴミ庫へ（使用後は施錠）。</p>' },
      checkout: { s: 'チェックアウトのご案内', b:
        '<p>明日 <b>' + co + ' 10:00</b> チェックアウトです（レイトアウト不可）。</p>'
        + '<p>ゴミは分別のうえ屋外ゴミ庫（施錠）または室内へ。<br>忘れ物は3日間保管・着払いで郵送可。<br>チェックアウト時刻が分かればお知らせください。</p>' },
      review: { s: 'ありがとうございました', b:
        '<p>この度はご滞在いただき誠にありがとうございました。素敵なゲストにお会いできて嬉しかったです。</p>'
        + '<p>もしよろしければレビューをいただけますと今後の励みになります。またのお越しを心よりお待ちしております。</p>' }
    },
    en: {
      welcome: { s: 'Reservation confirmed', b:
        '<p>Your payment is complete and reservation <b>' + id + '</b> is confirmed. Thank you!</p>'
        + '<p>Check-in: ' + ci + ' 16:00<br>Check-out: ' + co + ' 10:00 (no late check-out)</p>'
        + '<p>We will send the full check-in guide (entry method, Wi-Fi) the day before check-in.</p>'
        + '<p>[House rules] No smoking anywhere (incl. rooftop/balcony). Residential area, please keep quiet after 21:00, no large parties. Separate garbage into 45L bags.</p>'
        + '<p>Transfers: ' + F.transfer + '<br>Experiences: ' + F.tours + '</p><p>' + info + '</p>' },
      arrival: { s: 'Your estimated arrival time', b:
        '<p>Your check-in is coming up. To prepare, could you let us know your <b>estimated arrival time</b>? (Self check-in, anytime from 16:00.)</p>'
        + gbtn('View house manual')
        + '<p>Check-in: ' + ci + ' 16:00 / Check-out: ' + co + ' 10:00</p>' },
      checkin_guide: { s: 'Check-in guide (arriving tomorrow)', b:
        '<p>You check in tomorrow from <b>' + ci + ' 16:00</b>. This is a self check-in, so please read the check-in guide in advance (smart-lock code, Wi-Fi, access map with photos).</p>'
        + gbtn('Open check-in guide')
        + '<p>Address: ' + F.addr + '<br>Map: ' + F.map + '<br>Wi-Fi SSID: <b>' + F.wifi + '</b> (password inside the guide)</p>'
        + '<p>Check-out: ' + co + ' 10:00 / No smoking / Emergency: ' + F.tel + '</p>' },
      garbage: { s: 'Garbage separation', b:
        '<p>Good morning. Please separate garbage using the <b>45L bags</b> (burnable / glass / cans / PET bottles).</p>'
        + '<p>A sorting fee of &yen;4,400 may apply if garbage is not separated. When bins are full, use the outdoor storage and lock it.</p>' },
      checkout: { s: 'Check-out information', b:
        '<p>Check-out is tomorrow by <b>' + co + ' 10:00</b> (no late check-out).</p>'
        + '<p>Separate garbage and place it in the outdoor storage (lock it) or leave it inside.<br>Lost items are kept for 3 days and can be mailed (shipping collect).<br>Please let us know your check-out time if possible.</p>' },
      review: { s: 'Thank you', b:
        '<p>Thank you very much for staying with us. It was a pleasure to host you.</p>'
        + '<p>If you have a moment, we would be grateful for a review. We hope to welcome you again!</p>' }
    },
    zh: {
      welcome: { s: '預訂已確定', b:
        '<p>您的付款已完成，預訂 <b>' + id + '</b> 已正式確定，謝謝您！</p>'
        + '<p>入住：' + ci + ' 16:00<br>退房：' + co + ' 10:00（不可延遲退房）</p>'
        + '<p>詳細的入住指南（進門方式、WiFi）將於入住前一天寄給您。</p>'
        + '<p>[住宿規則] 全館禁菸（含屋頂/陽台）。位於住宅區，21:00 後請保持安靜、禁止大型派對。垃圾請以 45L 專用袋分類。</p>'
        + '<p>接送：' + F.transfer + '<br>體驗行程：' + F.tours + '</p><p>' + info + '</p>' },
      arrival: { s: '您的預計抵達時間', b:
        '<p>即將入住。為方便準備，可否告知您的<b>預計抵達時間</b>？（自助入住，16:00 後皆可）</p>'
        + gbtn('查看住宿手冊')
        + '<p>入住：' + ci + ' 16:00 / 退房：' + co + ' 10:00</p>' },
      checkin_guide: { s: '入住指南（明日入住）', b:
        '<p>您將於明日 <b>' + ci + ' 16:00</b> 起入住。本館為自助入住，請務必<b>事先</b>閱讀入住指南（智慧鎖密碼、WiFi、含照片的路線圖）。</p>'
        + gbtn('開啟入住指南')
        + '<p>地址：' + F.addr + '<br>地圖：' + F.map + '<br>WiFi SSID：<b>' + F.wifi + '</b>（密碼在指南內）</p>'
        + '<p>退房：' + co + ' 10:00 / 全館禁菸 / 緊急聯絡：' + F.tel + '</p>' },
      garbage: { s: '垃圾分類', b:
        '<p>早安。垃圾請使用 <b>45L 專用袋</b>分類（可燃/玻璃瓶/罐/寶特瓶）。</p>'
        + '<p>※ 未分類可能收取 &yen;4,400 分類費。垃圾滿時請放入室外垃圾庫並上鎖。</p>' },
      checkout: { s: '退房須知', b:
        '<p>明日 <b>' + co + ' 10:00</b> 退房（不可延遲退房）。</p>'
        + '<p>垃圾請分類後放入室外垃圾庫（上鎖）或留在室內。<br>遺失物保管 3 天，可貨到付款寄送。<br>若已知退房時間，請告知我們。</p>' },
      review: { s: '感謝您的入住', b:
        '<p>非常感謝您這次的入住，很高興能招待您。</p>'
        + '<p>若您方便，懇請給予評價，將是我們最大的鼓勵。期待再次為您服務！</p>' }
    },
    ko: {
      welcome: { s: '예약이 확정되었습니다', b:
        '<p>결제가 완료되어 예약 <b>' + id + '</b> 이 확정되었습니다. 감사합니다!</p>'
        + '<p>체크인: ' + ci + ' 16:00<br>체크아웃: ' + co + ' 10:00 (레이트 체크아웃 불가)</p>'
        + '<p>입실 방법·Wi-Fi 등 상세 체크인 가이드는 체크인 전날 보내드립니다.</p>'
        + '<p>[하우스 룰] 전관 금연(옥상/발코니 포함). 주택가이므로 21:00 이후 정숙, 대규모 파티 금지. 쓰레기는 45L 지정 봉투로 분리.</p>'
        + '<p>픽업: ' + F.transfer + '<br>체험/투어: ' + F.tours + '</p><p>' + info + '</p>' },
      arrival: { s: '예상 도착 시간 안내', b:
        '<p>곧 체크인입니다. 준비를 위해 <b>예상 도착 시간</b>을 알려주시겠어요? (셀프 체크인, 16:00 이후 언제든 가능)</p>'
        + gbtn('하우스 매뉴얼 보기')
        + '<p>체크인: ' + ci + ' 16:00 / 체크아웃: ' + co + ' 10:00</p>' },
      checkin_guide: { s: '체크인 가이드(내일 입실)', b:
        '<p>내일 <b>' + ci + ' 16:00</b>부터 체크인입니다. 셀프 체크인이므로 <b>사전에</b> 체크인 가이드를 확인해 주세요(스마트록 번호, Wi-Fi, 사진 포함 약도).</p>'
        + gbtn('체크인 가이드 열기')
        + '<p>주소: ' + F.addr + '<br>지도: ' + F.map + '<br>Wi-Fi SSID: <b>' + F.wifi + '</b>(비밀번호는 가이드 안)</p>'
        + '<p>체크아웃: ' + co + ' 10:00 / 전관 금연 / 긴급 연락: ' + F.tel + '</p>' },
      garbage: { s: '쓰레기 분리배출', b:
        '<p>좋은 아침입니다. 쓰레기는 <b>45L 지정 봉투</b>로 분리해 주세요(가연/유리병/캔/페트병).</p>'
        + '<p>※ 미분리 시 분리 수수료 &yen;4,400가 부과될 수 있습니다. 가득 차면 실외 보관고에 넣고 잠가 주세요.</p>' },
      checkout: { s: '체크아웃 안내', b:
        '<p>내일 <b>' + co + ' 10:00</b> 체크아웃입니다(레이트 체크아웃 불가).</p>'
        + '<p>쓰레기는 분리 후 실외 보관고(잠금) 또는 실내에 두세요.<br>분실물은 3일간 보관, 착불 발송 가능.<br>체크아웃 시간을 알면 알려주세요.</p>' },
      review: { s: '감사합니다', b:
        '<p>이번에 숙박해 주셔서 진심으로 감사합니다. 모실 수 있어 기뻤습니다.</p>'
        + '<p>괜찮으시면 리뷰를 남겨 주시면 큰 힘이 됩니다. 다시 뵙기를 기대하겠습니다!</p>' }
    }
  };
  var t = (D[lang] || D.en)[kind] || D.en[kind];
  return { subject: t.s, body: t.b };
}

/** Send a stay email (kind) to the reservation guest; primary language + EN. */
function sendStay_(row, kind) {
  var to = row.rep_email;
  if (!to) return false;
  var lang = pickLang_(row.rep_country);
  var ctx = {
    id: row.id, ci: toYMDSafe_(row.checkin), co: toYMDSafe_(row.checkout),
    name: fullName_(row),
    guide: getProp_('HOUSE_MANUAL_URL', 'https://drive.google.com/file/d/1hyeG_lICB7TpTTGsE-CJ-QaK6F8OHyaJ/view?usp=sharing')
  };
  var m = stayEmail_(kind, lang, ctx);
  var en = (lang !== 'en') ? stayEmail_(kind, 'en', ctx) : null;
  var html = '<p>' + esc_(ctx.name) + ' 様 / Dear guest,</p>' + m.body + (en ? '<hr>' + en.body : '');
  GmailApp.sendEmail(to, '[Komei Hotel] ' + m.subject + ' (' + row.id + ')', '', { htmlBody: html, name: getProp_('FROM_NAME', 'Komei Hotel') });
  return true;
}

/** Map reservation_id -> {kind: true} for stay emails already sent (from logs). */
function stayEmailSentMap_() {
  var sh = sheet_('logs');
  ensureHeaders_(sh, HEADERS_LOGS);
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== 'stay_email') continue;
    var id = data[i][1];
    var m = String(data[i][3] || '').match(/kind=(\w+)/);
    if (!m) continue;
    if (!map[id]) map[id] = {};
    map[id][m[1]] = true;
  }
  return map;
}

/**
 * Date-based stay emails B-F for every PAID reservation (one per run, dedup via logs):
 *   arrival: 3 days before check-in | checkin_guide: 1 day before
 *   garbage: morning after check-in | checkout: 1 day before check-out
 *   review: after check-out
 */
function sendStayEmails() {
  var sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  var data = sh.getDataRange().getValues();
  var H = data[0];
  var iId = H.indexOf('id'), iStatus = H.indexOf('status'), iCi = H.indexOf('checkin'), iCo = H.indexOf('checkout');
  var sent = stayEmailSentMap_();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iStatus]) !== STATUS.PAID) continue;
    var id = data[i][iId];
    if (!id) continue;
    var ci = data[i][iCi];
    if (!ci) continue;
    var co = data[i][iCo];
    var dCi = daysUntilCheckin_(ci);
    var dCo = co ? daysUntilDate_(toYMDSafe_(co)) : null;
    var s = sent[id] || {};
    var kind = null;
    if (dCi <= 3 && dCi >= 1 && !s.arrival) kind = 'arrival';
    else if (dCi <= 1 && dCi >= 0 && !s.checkin_guide) kind = 'checkin_guide';
    else if (dCi <= -1 && dCo !== null && dCo >= 2 && !s.garbage) kind = 'garbage';
    else if (dCo !== null && dCo <= 1 && dCo >= 0 && !s.checkout) kind = 'checkout';
    else if (dCo !== null && dCo <= -1 && dCo >= -3 && !s.review) kind = 'review';
    if (!kind) continue;
    var row = {};
    for (var j = 0; j < H.length; j++) row[H[j]] = data[i][j];
    try {
      if (sendStay_(row, kind)) { log_(id, 'stay_email', 'kind=' + kind); count++; }
    } catch (e) { log_(id, 'stay_email_error', kind + ':' + String(e)); }
  }
  log_(null, 'stay_emails_run', 'sent=' + count);
  return { sent: count };
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

/** Stripe webhook signature verification (HMAC-SHA256) */
function verifyStripeSignature_(payload, sigHeader, secret) {
  try {
    var pairs = {};
    sigHeader.split(',').forEach(function(part) {
      var kv = part.trim().split('=');
      if (kv.length === 2) pairs[kv[0]] = kv[1];
    });
    var timestamp = pairs['t'];
    var expectedSig = pairs['v1'];
    if (!timestamp || !expectedSig) return false;
    // Reject events older than 5 minutes (tolerance for clock skew)
    var age = Math.floor(Date.now() / 1000) - parseInt(timestamp);
    if (Math.abs(age) > 300) return false;
    var signedPayload = timestamp + '.' + payload;
    var mac = Utilities.computeHmacSha256Signature(signedPayload, secret);
    var computed = mac.map(function(b) {
      var hex = (b < 0 ? b + 256 : b).toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    }).join('');
    return computed === expectedSig;
  } catch (err) {
    log_(null, 'stripe_sig_error', err.toString());
    return false;
  }
}

/** HTML-escape to prevent XSS in admin email HTML */
function esc_(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
 * Mask a name for privacy: '山田太郁E →'山***', 'John Smith' →'J*** S***'
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
 * Mask an email for privacy: 'user@example.com' →'u***@e***.com'
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

  // Messages  Emark guest messages as read by host
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
    '<p>' + name + ' 様/p>'
    + '<p>Komei Hotel からメッセージが届いています！</p>'
    + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb;margin:12px 0">'
    + message.replace(/\n/g,'<br>') + '</blockquote>'
    + '<p><a href="' + mypageUrl + '" style="background:#f59e0b;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px">マイページで確認する</a></p>'
    + '<hr><p>Dear ' + name + ', you have a new message from Komei Hotel. '
    + '<a href="' + mypageUrl + '">View on My Page →/a></p>';
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
    { htmlBody: '<p>予約 <b>' + id + '</b>（' + esc_(fullName_(r.row)) + '）からメッセージ。</p>'
        + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb">'
        + message.replace(/\n/g,'<br>') + '</blockquote>'
        + '<p><a href="' + adminUrl + '">管理画面で確認する</a></p>' });

  // ----- Auto-reply: deterministic, template-based -----
  let autoReply = null;
  try {
    const lang  = detectLang_(message);
    const match = matchAutoReply_(message, lang);
    if (match && match.reply) {
      addMessage_(id, 'auto', match.reply);
      log_(id, 'auto_reply', match.intent + ' (' + lang + ')');
      autoReply = { intent: match.intent, label: match.label, reply: match.reply };
    }
  } catch (autoErr) {
    log_(id, 'auto_reply_error', String(autoErr).substring(0, 200));
  }

  return { ok:true, auto_reply: autoReply };
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
  const fullMsg   = '[変更リクエスト ' + typeLabel + ']\n' + detail;
  addMessage_(id, 'guest', fullMsg);
  log_(id, 'change_request', changeType + ': ' + detail.substring(0, 100));

  const adminUrl = getProp_('SITE_BASE_URL') + '/admin.html';
  GmailApp.sendEmail(getProp_('ADMIN_EMAIL'),
    '[Komei Hotel] 変更リクエスト/ Change Request (' + id + ')',
    '',
    { htmlBody: '<p>予約 <b>' + id + '</b>（' + esc_(fullName_(r.row)) + '）からの変更リクエスト！</p>'
        + '<p>種額 <b>' + typeLabel + '</b></p>'
        + '<blockquote style="border-left:4px solid #f59e0b;padding:12px;background:#fffbeb">'
        + detail.replace(/\n/g,'<br>') + '</blockquote>'
        + '<p><a href="' + adminUrl + '">管理画面で確認する</a></p>' });

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
    const subject = '【Komei Hotel】新しいレビュー (' + esc_(repName) + ' ★' + body.overall + ')';
    const html = '<h3>新しいレビューが投稿されました</h3>'
      + '<p><b>予約ID:</b> ' + esc_(body.reservation_id) + '<br>'
      + '<b>ゲスト:</b> ' + esc_(repName) + '<br>'
      + '<b>総合評価:</b> ' + '★'.repeat(body.overall) + ' (' + body.overall + '/5)<br>'
      + '<b>コメント</b><br>' + esc_(body.comment || '(なし)').replace(/\n/g, '<br>') + '</p>'
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
    const googleReviewUrl = getProp_('GOOGLE_REVIEW_URL') || 'https://www.google.com/maps/search/Komei+Hotel+光明荘+東駒形';

    const subject = '【Komei Hotel】ご宿泊ありがとうございました  Eレビューのお願い';
    const html = '<div style="max-width:600px;margin:0 auto;font-family:sans-serif;">'
      + '<h2 style="color:#d97706;">Komei Hotel 光明荘/h2>'
      + '<p>' + name + ' 様/p>'
      + '<p>この度はKomei Hotelにご宿泊いただき、誠にありがとうございました</p>'
      + '<p>ご滞在はいかがでしたか？今後のサービス向上のため、ぜひレビューをお聞かせください</p>'
      + '<p style="text-align:center;margin:30px 0;">'
      + '<a href="' + mypageUrl + '" style="display:inline-block;background:#f59e0b;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:16px;">レビューを書い/a>'
      + '</p>'
      + '<p style="text-align:center;margin:20px 0;">'
      + '<a href="' + googleReviewUrl + '" style="display:inline-block;background:#4285f4;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;">📍 Googleにもレビューを書い/a>'
      + '</p>'
      + '<p style="color:#94a3b8;font-size:13px;">マイページにログイン後、「レビュー」タブからご記入いただけます<br>Googleレビューもいただけると大変嬉しいです</p>'
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

// ============ Review Trigger Setup ============

function setupReviewTrigger() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendReviewRequestEmails') ScriptApp.deleteTrigger(t);
  });
  // Daily at 10:00 JST
  ScriptApp.newTrigger('sendReviewRequestEmails')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('Review trigger set: daily 10:00 JST');
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
  ensureHeaders_(sheet_('auto_replies'), HEADERS_AUTO_REPLIES);
  seedAutoReplies_(false); // false = do not overwrite existing
  generateAndStoreAdminToken_();
  Logger.log('Initialized. ADMIN_TOKEN=' + getProp_('ADMIN_TOKEN'));
}

// manualResend wrapper removed (was one-time debug tool)

// ============================================================
// Auto-reply system
// Templates extracted from past Airbnb host conversations.
// Stored in `auto_replies` sheet so admins can edit without code changes.
// ============================================================

/**
 * Default auto-reply templates. The 18 intents below cover the recurring
 * questions we saw across past Airbnb conversations (passport submission,
 * check-in/-out times, smart-lock code, WiFi, address, garbage rules,
 * smoking, noise, late check-out, transfers, tours, scam warning, etc.).
 *
 * Matching is keyword-based with priority order; higher priority wins
 * when multiple intents match. `fallback` matches anything as a safety
 * net so every guest message gets an instant acknowledgement.
 */
function defaultAutoReplies_() {
  return [
    {
      intent:'passport_register', label:'パスポート登録案内', priority:90,
      keywords_ja:'パスポート,名簿,ゲスト情報,本人確認,onestay',
      keywords_en:'passport,guest info,guest information,register,roster,onestay,check-in form',
      reply_ja:'ご連絡ありがとうございます。チェックインの手続きとして、ご予約確認メールに記載のオンラインチェックインリンクから、全ゲストの情報（国籍・パスポート番号・氏名・住所・連絡先）をご登録ください。リンクが見つからない場合はお知らせください、再送いたします。\n\n光明荘',
      reply_en:'Thank you for your message. To complete your check-in, please submit guest information (nationality, passport number, name, address, contact) for all guests via the online check-in link sent in your booking confirmation email. If you cannot find the link, please let us know and we will resend it.\n\nKomei Hotel'
    },
    {
      intent:'checkin_time', label:'チェックイン時間', priority:80,
      keywords_ja:'チェックイン時間,何時から,入室時間,いつ入れる',
      keywords_en:'check-in time,what time can,arrival time,checkin time,when can we check in',
      reply_ja:'チェックインは16:00（午後4:00）以降ご利用いただけます。ご到着予定時刻をお知らせいただけますと、スムーズにお迎えできます。\n\n光明荘',
      reply_en:'Check-in is available from 4:00 PM (16:00) onwards. Please let us know your estimated arrival time so we can prepare for your stay.\n\nKomei Hotel'
    },
    {
      intent:'checkout_time', label:'チェックアウト時間', priority:80,
      keywords_ja:'チェックアウト時間,何時まで,退室時間,出る時間',
      keywords_en:'check-out time,checkout time,what time leave,departure time',
      reply_ja:'チェックアウトは午前10:00までとなっております。次のゲストのご案内のため、レイトチェックアウトは原則お受けできません。事情がある場合は事前にご相談ください。\n\n光明荘',
      reply_en:'Check-out is by 10:00 AM. Late check-out is generally not available because we need to prepare the room for the next guest. If you have special circumstances, please contact us in advance.\n\nKomei Hotel'
    },
    {
      intent:'late_checkout', label:'レイトチェックアウト依頼', priority:85,
      keywords_ja:'レイトチェックアウト,延長,延泊,遅く出',
      keywords_en:'late check-out,late checkout,extend,stay later,extend stay',
      reply_ja:'レイトチェックアウトは原則お受けできません。次のゲストの清掃に支障が出る場合、追加料金を頂戴することがございます。延泊をご希望の場合は、チェックアウト前日18:00までにご連絡をお願いいたします。\n\n光明荘',
      reply_en:'Late check-out is generally not available. If your delay impacts cleaning for the next guest, additional fees may apply. To extend your stay, please contact us by 6:00 PM the day before check-out.\n\nKomei Hotel'
    },
    {
      intent:'smartlock_code', label:'入室暗証番号', priority:85,
      keywords_ja:'暗証番号,鍵,スマートロック,入室方法,ドア,開かない',
      keywords_en:'access code,door code,key,smart lock,smartlock,how to enter,how to get in,door won\'t open,can\'t open',
      reply_ja:'スマートロックの暗証番号は、チェックイン日の前日にチェックインガイドと一緒にメールでお送りします。すでにチェックイン済みで暗証番号にお困りの場合は、本メッセージにご返信のうえ、緊急時は03-6899-5681までお電話ください。\n\n光明荘',
      reply_en:'The smart-lock access code is emailed to you the day before check-in along with the check-in guide. If you have already checked in and are having trouble with the code, please reply here, and call 03-6899-5681 for emergencies.\n\nKomei Hotel'
    },
    {
      intent:'wifi', label:'Wi-Fi情報', priority:75,
      keywords_ja:'wifi,wi-fi,ワイファイ,インターネット,パスワード',
      keywords_en:'wifi,wi-fi,internet,password,network',
      reply_ja:'館内全体でWiFiをご利用いただけます。SSIDは「Komei-Guest」、パスワードはチェックイン前日にチェックインガイドと一緒にメールでお知らせします。\n\n光明荘',
      reply_en:'WiFi is available throughout the property. The SSID is "Komei-Guest" and the password will be sent to you the day before check-in together with the check-in guide.\n\nKomei Hotel'
    },
    {
      intent:'address', label:'住所・アクセス', priority:75,
      keywords_ja:'住所,場所,行き方,地図,アクセス,どこ',
      keywords_en:'address,location,where is,how to get there,map,directions,google map',
      reply_ja:'住所：〒130-0005 東京都墨田区東駒形4-20-5\nGoogleマップ：https://maps.app.goo.gl/jVzhawyQTeTFfkLcA\n\nGoogleマップに加え、チェックインガイドの写真付きアクセスマップも併せてご確認ください。\n\n光明荘',
      reply_en:'Property address: 4-20-5 Higashikomagata, Sumida-ku, Tokyo 130-0005, Japan\nGoogle Maps: https://maps.app.goo.gl/jVzhawyQTeTFfkLcA\n\nWhen visiting, please use the access map with photos in the check-in guide together with Google Maps.\n\nKomei Hotel'
    },
    {
      intent:'garbage_rules', label:'ゴミ出しルール', priority:70,
      keywords_ja:'ゴミ,ごみ,分別,捨て方,袋',
      keywords_en:'garbage,trash,rubbish,waste,recycle,separate,bin',
      reply_ja:'ゴミは、燃えるゴミ・瓶・缶・ペットボトルに分別をお願いいたします。\n施設で用意した45L指定袋のみをご利用ください（紙袋など他の袋は使用不可）。\n長期滞在の方には、収集日を別途メッセージでご案内いたします。\n\n光明荘',
      reply_en:'Please separate burnable garbage, glass bottles, cans, and PET bottles.\nUse only the 45L garbage bags provided by our facility (paper bags or other bags are not allowed).\nFor long-term stays, collection days will be communicated separately.\n\nKomei Hotel'
    },
    {
      intent:'smoking', label:'禁煙ルール', priority:65,
      keywords_ja:'タバコ,たばこ,煙草,喫煙,吸って',
      keywords_en:'smoke,smoking,cigarette,vape,vaping',
      reply_ja:'当宿は屋内・バルコニーを含む屋外も含め、敷地内すべて禁煙となっております。何卒ご理解とご協力をお願いいたします。\n\n光明荘',
      reply_en:'Smoking is strictly prohibited anywhere on the property, including indoors and outdoors (terrace included). Thank you for your understanding and cooperation.\n\nKomei Hotel'
    },
    {
      intent:'noise', label:'騒音・近隣配慮', priority:65,
      keywords_ja:'騒音,うるさい,音,パーティー,深夜',
      keywords_en:'noise,loud,party,quiet,neighbor,complain',
      reply_ja:'当宿は住宅街にございます。通常の会話は問題ございませんが、大声でのパーティーや電話通話はお控えください。21:00以降は特にお静かにお過ごしいただきますようお願いいたします。繰り返し苦情が寄せられた場合、警察への通報や追加料金が発生することがございます。\n\n光明荘',
      reply_en:'We are located in a residential area. Normal conversation is fine, but please avoid loud parties and phone calls. Please keep noise to a minimum, especially after 9:00 PM. Repeated complaints may result in police involvement and additional charges.\n\nKomei Hotel'
    },
    {
      intent:'arrival_time', label:'到着予定時刻のお伺い', priority:60,
      keywords_ja:'到着,到着予定,何時に着,つきます,着きます',
      keywords_en:'arrive,arrival,arriving,we will arrive,coming at',
      reply_ja:'ご到着予定時刻をお知らせいただきありがとうございます。事前にご準備を整えてお迎えいたします。お気をつけてお越しください。\n\n光明荘',
      reply_en:'Thank you for letting us know your estimated arrival time. We will have everything ready for you. Have a safe trip!\n\nKomei Hotel'
    },
    {
      intent:'transfer', label:'送迎・空港シャトル', priority:55,
      keywords_ja:'送迎,シャトル,空港,タクシー,ピックアップ',
      keywords_en:'transfer,shuttle,airport pickup,airport pick-up,taxi,pick up',
      reply_ja:'有料の送迎サービス（空港送迎を含む）をご用意しております。手荷物の心配なく快適にご移動いただけます。\n詳細はこちら：https://tokyo-door-to-door.netlify.app/#tours\n\nご希望の場合はその旨ご返信ください。\n\n光明荘',
      reply_en:'We offer paid transfer services including airport pickup. Travel comfortably with no luggage hassle or transfers.\nDetails: https://tokyo-door-to-door.netlify.app/#tours\n\nIf you would like to book, please reply to this message.\n\nKomei Hotel'
    },
    {
      intent:'tour', label:'観光ツアー', priority:50,
      keywords_ja:'ツアー,観光,富士山,日光,体験',
      keywords_en:'tour,sightseeing,fuji,nikko,experience,day trip',
      reply_ja:'富士山・日光などへの1日プライベートツアーをご用意しております。信頼できる地元のガイドによる、上質な体験をお楽しみいただけます。\n詳細はこちら：https://tokyo-experience.web.app\n\n光明荘',
      reply_en:'We offer 1-day private tours to Mt. Fuji, Nikko, and other destinations. Enjoy private, high-quality experiences crafted through trusted local connections.\nDetails: https://tokyo-experience.web.app\n\nKomei Hotel'
    },
    {
      intent:'payment_security', label:'決済詐欺への警告', priority:99,
      keywords_ja:'クレジットカード,カード番号,支払い情報,銀行振込,暗号資産',
      keywords_en:'credit card,card number,payment info,bank transfer,wire transfer,bitcoin,gift card',
      reply_ja:'⚠️【安全のお知らせ】Komei Hotelは、本チャットでクレジットカードや支払い情報をお伺いすることは絶対にございません。第三者から決済情報を求められた場合は、必ず本アプリまたは公式メール（komei.hotel@gmail.com）から直接ご確認ください。\n\n光明荘',
      reply_en:'⚠️ For your safety: Komei Hotel will NEVER ask for credit card or payment information via chat. If you are asked for payment details by anyone, please verify directly with us through this app or our official email (komei.hotel@gmail.com).\n\nKomei Hotel'
    },
    {
      intent:'emergency', label:'緊急時連絡先', priority:95,
      keywords_ja:'緊急,救急,事故,怪我,火事,警察',
      keywords_en:'emergency,urgent,accident,injury,fire,police,help me',
      reply_ja:'緊急の際は、03-6899-5681までお電話、または komei.hotel@gmail.com にメールでご連絡ください。スタッフが早急に対応いたします。\n\n光明荘',
      reply_en:'For emergencies, please call 03-6899-5681 directly, or email komei.hotel@gmail.com. Our team will respond as quickly as possible.\n\nKomei Hotel'
    },
    {
      intent:'thanks', label:'お礼への返信', priority:30,
      keywords_ja:'ありがとう,感謝,助かりました',
      keywords_en:'thank you,thanks,appreciate,grateful',
      reply_ja:'温かいお言葉をありがとうございます！素敵なご滞在となりますように。何かございましたらお気軽にご連絡ください 😊\n\n光明荘',
      reply_en:'Thank you so much for your kind words! We hope you have a wonderful stay. Please do not hesitate to reach out if you need anything 😊\n\nKomei Hotel'
    },
    {
      intent:'greeting', label:'挨拶への返信', priority:20,
      keywords_ja:'こんにちは,はじめまして,お世話になります,よろしく',
      keywords_en:'hello,hi there,good morning,good evening,nice to meet',
      reply_ja:'ご連絡ありがとうございます。Komei Hotelへようこそ。スタッフより通常12時間以内にご返信いたします。お急ぎの場合は03-6899-5681までお電話ください。\n\n光明荘',
      reply_en:'Thank you for your message and welcome to Komei Hotel. Our team will get back to you within 12 hours. For urgent matters, please call 03-6899-5681.\n\nKomei Hotel'
    },
    {
      intent:'fallback', label:'デフォルト返信（該当なし）', priority:1,
      keywords_ja:'*',
      keywords_en:'*',
      reply_ja:'ご連絡ありがとうございます。スタッフに通知され、通常12時間以内にご返信いたします。お急ぎの場合は03-6899-5681、または komei.hotel@gmail.com までご連絡ください。\n\n光明荘',
      reply_en:'Thank you for your message. Our team has been notified and will respond within 12 hours. For urgent matters, please call 03-6899-5681 or email komei.hotel@gmail.com.\n\nKomei Hotel'
    }
  ];
}

/**
 * Insert defaults into the auto_replies sheet. If `overwrite` is true,
 * existing rows are wiped first; otherwise only intents not already
 * present are appended.
 */
function seedAutoReplies_(overwrite) {
  const sh = sheet_('auto_replies');
  ensureHeaders_(sh, HEADERS_AUTO_REPLIES);

  const defaults = defaultAutoReplies_();
  const now = new Date().toISOString();

  if (overwrite && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_AUTO_REPLIES.length).clearContent();
  }

  const existingIntents = {};
  if (sh.getLastRow() > 1) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) existingIntents[String(data[i][0])] = true;
  }

  const rows = [];
  defaults.forEach(function(d) {
    if (!overwrite && existingIntents[d.intent]) return;
    rows.push([
      d.intent, d.label, d.priority, true,
      d.keywords_ja, d.keywords_en,
      d.reply_ja, d.reply_en, now
    ]);
  });
  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS_AUTO_REPLIES.length).setValues(rows);
  }
}

/**
 * Read all auto-reply rules from the sheet, sorted by priority desc.
 */
function getAutoReplies_() {
  const sh = sheet_('auto_replies');
  ensureHeaders_(sh, HEADERS_AUTO_REPLIES);
  if (sh.getLastRow() <= 1) {
    seedAutoReplies_(false);
    if (sh.getLastRow() <= 1) return [];
  }
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    const o = {};
    headers.forEach(function(h, j) { o[h] = data[i][j]; });
    o.priority = Number(o.priority) || 0;
    o.enabled  = (o.enabled === true || String(o.enabled).toLowerCase() === 'true');
    rules.push(o);
  }
  rules.sort(function(a, b) { return b.priority - a.priority; });
  return rules;
}

/**
 * Find the auto-reply rule whose keyword list (for the given language)
 * has the strongest match against the message. Returns null when only
 * the wildcard fallback would match and we want to skip auto-reply.
 */
function matchAutoReply_(message, lang) {
  const rules = getAutoReplies_();
  const lc = String(message).toLowerCase();
  let best = null;

  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    if (!r.enabled) continue;
    const kwField = (lang === 'ja') ? r.keywords_ja : r.keywords_en;
    const kws = String(kwField || '').split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean);
    if (kws.length === 0) continue;

    if (kws.indexOf('*') !== -1) {
      if (!best) best = r; // fallback only if nothing else matched yet
      continue;
    }

    let hits = 0;
    for (let k = 0; k < kws.length; k++) {
      if (lc.indexOf(kws[k]) !== -1) hits++;
    }
    if (hits > 0) { best = r; break; } // priority-sorted, first hit wins
  }
  if (!best) return null;

  return {
    intent: best.intent,
    label:  best.label,
    reply:  (lang === 'ja') ? best.reply_ja : best.reply_en
  };
}

/**
 * Crude language detection. If the message contains Japanese script
 * (hiragana/katakana/CJK), return 'ja'; otherwise 'en'. Sufficient for
 * picking which template variant to send back.
 */
function detectLang_(text) {
  if (!text) return 'en';
  return /[぀-ヿ㐀-鿿]/.test(text) ? 'ja' : 'en';
}

// ============ Auto-reply admin handlers ============

function handleAdminListAutoReplies(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  return { ok:true, rules: getAutoReplies_() };
}

function handleAdminUpdateAutoReply(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const intent = (body.intent || '').trim();
  if (!intent) return { ok:false, error:'no intent' };

  const sh = sheet_('auto_replies');
  ensureHeaders_(sh, HEADERS_AUTO_REPLIES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = {};
  headers.forEach(function(h, j) { idx[h] = j; });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.intent]) === intent) {
      const r = i + 1;
      if (body.label       != null) sh.getRange(r, idx.label       + 1).setValue(body.label);
      if (body.priority    != null) sh.getRange(r, idx.priority    + 1).setValue(Number(body.priority) || 0);
      if (body.enabled     != null) sh.getRange(r, idx.enabled     + 1).setValue(body.enabled === true || body.enabled === 'true');
      if (body.keywords_ja != null) sh.getRange(r, idx.keywords_ja + 1).setValue(body.keywords_ja);
      if (body.keywords_en != null) sh.getRange(r, idx.keywords_en + 1).setValue(body.keywords_en);
      if (body.reply_ja    != null) sh.getRange(r, idx.reply_ja    + 1).setValue(body.reply_ja);
      if (body.reply_en    != null) sh.getRange(r, idx.reply_en    + 1).setValue(body.reply_en);
      sh.getRange(r, idx.updated_at + 1).setValue(new Date().toISOString());
      log_(null, 'auto_reply_update', intent);
      return { ok:true };
    }
  }
  return { ok:false, error:'intent not found' };
}

function handleAdminResetAutoReplies(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  seedAutoReplies_(true); // overwrite all
  log_(null, 'auto_reply_reset', 'all defaults restored');
  return { ok:true };
}

function handleAdminTestAutoReply(body) {
  if (!verifyAdminToken_(body.admin_token)) return { ok:false, error:'unauthorized' };
  const message = (body.message || '').trim();
  if (!message) return { ok:false, error:'no message' };
  const lang  = body.lang || detectLang_(message);
  const match = matchAutoReply_(message, lang);
  return { ok:true, lang:lang, match:match };
}

// =====================================================================
// Review Request Automation
// ---------------------------------------------------------------------
// Daily cron sends a single review-request email per reservation, the day
// after checkout. Idempotent — re-running the cron won't double-send because
// we record a 'review_request_sent' entry in the logs sheet and skip
// reservations that already have one.
//
// Setup steps (one-time):
//   1. In the Apps Script editor, run `setupReviewRequestTrigger()` once.
//   2. Authorize MailApp + ScriptApp scopes when prompted.
//   3. From then on, the trigger fires daily at ~09:00 JST and sends emails.
//
// Manual test / fire now:
//   - Run `sendReviewRequestEmails()` directly in the editor.
//   - Run `sendReviewRequestEmails({dryRun: true})` to log without sending.
// =====================================================================

/** Window for picking up post-stay reservations (days since checkout). */
const REVIEW_REQUEST_DAYS_AFTER_CHECKOUT_MIN = 1;
const REVIEW_REQUEST_DAYS_AFTER_CHECKOUT_MAX = 14; // grace window for missed cron runs

/** One-time setup: install a daily 09:00 JST trigger for sendReviewRequestEmails. */
function setupReviewRequestTrigger() {
  // Remove any existing triggers for the same function to keep it idempotent.
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendReviewRequestEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendReviewRequestEmails')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('Daily review-request trigger installed (09:00 JST)');
}

/**
 * Main cron job. Iterates paid reservations whose checkout date is within the
 * post-stay window and sends a single review-request email per reservation.
 *
 * @param {{dryRun?:boolean}} opts - dryRun=true logs what would happen but
 *     does not send emails or mark logs.
 * @return {{considered:number, sent:number, skipped:number, errors:number}}
 */
function sendReviewRequestEmails(opts) {
  opts = opts || {};
  const dryRun = !!opts.dryRun;

  const sh = sheet_('reservations');
  ensureHeaders_(sh, HEADERS_RESERVATIONS);
  if (sh.getLastRow() < 2) {
    return { considered:0, sent:0, skipped:0, errors:0 };
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_RESERVATIONS.length).getValues();
  const idx = {};
  HEADERS_RESERVATIONS.forEach(function(h, i) { idx[h] = i; });

  const alreadySent = collectReviewRequestSentIds_();
  const now = new Date();
  const todayStr = formatDateJst_(now);

  let considered = 0;
  let sent = 0;
  let skipped = 0;
  let errors = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const reservationId = row[idx.id];
    const status = row[idx.status];
    const checkoutRaw = row[idx.checkout];
    const email = row[idx.rep_email];
    if (!reservationId || !email) continue;
    if (status !== STATUS.PAID) continue;

    const checkout = parseDate_(checkoutRaw);
    if (!checkout) continue;
    const daysSince = daysBetween_(checkout, now);
    if (daysSince < REVIEW_REQUEST_DAYS_AFTER_CHECKOUT_MIN) continue;
    if (daysSince > REVIEW_REQUEST_DAYS_AFTER_CHECKOUT_MAX) continue;

    considered++;
    if (alreadySent[reservationId]) {
      skipped++;
      continue;
    }

    const rowObj = {};
    HEADERS_RESERVATIONS.forEach(function(h, j) { rowObj[h] = row[j]; });

    try {
      if (!dryRun) {
        sendReviewRequestEmail_(rowObj);
        log_(reservationId, 'review_request_sent', todayStr + ' to ' + email);
      } else {
        log_(reservationId, 'review_request_sent_dryrun', todayStr + ' to ' + email);
      }
      sent++;
    } catch (e) {
      errors++;
      log_(reservationId, 'review_request_error', String(e));
    }
  }

  const summary = { considered:considered, sent:sent, skipped:skipped, errors:errors };
  Logger.log('Review request cron: ' + JSON.stringify(summary) + (dryRun ? ' [dry-run]' : ''));
  return summary;
}

/** Gather reservation IDs that already received a request (or had a dry-run logged). */
function collectReviewRequestSentIds_() {
  const sh = sheet_('logs');
  ensureHeaders_(sh, HEADERS_LOGS);
  const seen = {};
  if (sh.getLastRow() < 2) return seen;
  const logs = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_LOGS.length).getValues();
  const actionIdx = HEADERS_LOGS.indexOf('action');
  const resIdx = HEADERS_LOGS.indexOf('reservation_id');
  for (let i = 0; i < logs.length; i++) {
    const a = logs[i][actionIdx];
    if (a === 'review_request_sent' || a === 'review_request_sent_dryrun') {
      seen[logs[i][resIdx]] = true;
    }
  }
  return seen;
}

/** Compose + send the review request email (language picked from rep_country). */
function sendReviewRequestEmail_(row) {
  const adminEmail = getProp_('ADMIN_EMAIL', '');
  const fromName = getProp_('FROM_NAME', 'Komei Hotel');
  const baseUrl = getProp_('SITE_BASE_URL', 'https://komei.yoshinarcorp.com');
  const mypageUrl = baseUrl + '/mypage.html?id=' + encodeURIComponent(row.id)
    + '&email=' + encodeURIComponent(row.rep_email)
    + '#reviewSection';
  // Use GOOGLE_REVIEW_URL Script Property if set (any valid review URL, e.g.
  // a g.page/r/.../review link, or maps.google.com/?cid=XXX). Falls back to
  // empty so the "leave a Google review" CTA is simply omitted from the email
  // when not configured. Note: writereview?placeid= only accepts the ChIJ
  // form, so we never construct it ourselves anymore — store the full URL.
  const googleReviewUrl = getProp_('GOOGLE_REVIEW_URL', '') || '';

  const isJa = isJapaneseGuest_(row);
  const subject = isJa
    ? 'ご滞在ありがとうございました — レビューのお願い ｜ Komei Hotel 光明荘'
    : 'Thank you for staying — share your experience ｜ Komei Hotel';

  const body = isJa
    ? buildReviewRequestBodyJa_(row, mypageUrl, googleReviewUrl)
    : buildReviewRequestBodyEn_(row, mypageUrl, googleReviewUrl);

  const options = { name: fromName };
  if (adminEmail) options.bcc = adminEmail;
  MailApp.sendEmail(row.rep_email, subject, body, options);
}

/** Heuristic: prefer Japanese if rep_country is JP or contains Japanese chars. */
function isJapaneseGuest_(row) {
  const country = String(row.rep_country || '').trim().toUpperCase();
  if (country === 'JP' || country === 'JAPAN' || country === '日本') return true;
  const name = String(row.rep_first_name || '') + String(row.rep_last_name || '');
  // Hiragana / Katakana / CJK Unified Ideographs ranges
  return /[぀-ゟ゠-ヿ一-鿿]/.test(name);
}

function buildReviewRequestBodyJa_(row, mypageUrl, googleReviewUrl) {
  const name = fullName_(row) || 'お客様';
  const lines = [];
  lines.push(name + ' 様');
  lines.push('');
  lines.push('先日は Komei Hotel 光明荘にご宿泊いただき、誠にありがとうございました。');
  lines.push('滞在中にお気づきの点や、印象に残ったことがあれば、ぜひ短いレビューを');
  lines.push('お寄せいただけませんでしょうか。');
  lines.push('');
  lines.push('▼ 1分でレビューを投稿（マイページ）');
  lines.push(mypageUrl);
  lines.push('');
  if (googleReviewUrl) {
    lines.push('▼ Google マップにもレビューをいただけると大変励みになります');
    lines.push(googleReviewUrl);
    lines.push('');
  }
  lines.push('お声は今後の運営とサービス改善の何よりの参考にさせていただきます。');
  lines.push('またのご来訪を心よりお待ち申し上げております。');
  lines.push('');
  lines.push('--');
  lines.push('Komei Hotel 光明荘');
  lines.push('東京都墨田区東駒形20-5');
  lines.push('https://komei.yoshinarcorp.com/');
  return lines.join('\n');
}

function buildReviewRequestBodyEn_(row, mypageUrl, googleReviewUrl) {
  const name = fullName_(row) || 'Guest';
  const lines = [];
  lines.push('Dear ' + name + ',');
  lines.push('');
  lines.push('Thank you for staying with us at Komei Hotel. We hope you enjoyed your time');
  lines.push('in Asakusa and Sumida.');
  lines.push('');
  lines.push('If you have a minute, we would deeply appreciate a short review of your stay.');
  lines.push('It helps us improve and helps other travelers decide.');
  lines.push('');
  lines.push('▸ Leave a review in 1 minute (your guest page):');
  lines.push('  ' + mypageUrl);
  lines.push('');
  if (googleReviewUrl) {
    lines.push('▸ Or share on Google Maps:');
    lines.push('  ' + googleReviewUrl);
    lines.push('');
  }
  lines.push('Thank you again — we hope to welcome you back to Tokyo someday.');
  lines.push('');
  lines.push('--');
  lines.push('Komei Hotel 光明荘');
  lines.push('20-5 Higashikomagata, Sumida, Tokyo, Japan');
  lines.push('https://komei.yoshinarcorp.com/');
  return lines.join('\n');
}

/** Parse a checkout cell which may be a Date object or a YYYY-MM-DD string. */
function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (!s) return null;
  // Sheets normally returns dates as Date objects; this is a defensive fallback.
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Whole-day delta between two dates (b - a) in days, JST-anchored midnight. */
function daysBetween_(a, b) {
  const tz = 'Asia/Tokyo';
  const aStr = Utilities.formatDate(a, tz, 'yyyy-MM-dd');
  const bStr = Utilities.formatDate(b, tz, 'yyyy-MM-dd');
  const aMid = new Date(aStr + 'T00:00:00+09:00');
  const bMid = new Date(bStr + 'T00:00:00+09:00');
  return Math.round((bMid.getTime() - aMid.getTime()) / 86400000);
}

function formatDateJst_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/** 一度だけ実行: token列が空の全予約に token を補完し、登録リンクをログ出力 */
function backfillTokens() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('reservations');
  if (!sh) { ss.getSheets().forEach(function(s){ if(!sh && s.getRange(1,1,1,s.getLastColumn()).getValues()[0].indexOf('token')>=0) sh=s; }); }
  if (!sh) { Logger.log('reservations タブが見つかりません'); return; }
  var data = sh.getDataRange().getValues(), H = data[0];
  var cT=H.indexOf('token'), cId=H.indexOf('id'), cE=H.indexOf('rep_email'), cS=H.indexOf('status');
  if (cT<0||cId<0){ Logger.log('token/id 列なし: '+H.join(',')); return; }
  var base='https://komei.yoshinarcorp.com', n=0;
  for (var i=1;i<data.length;i++){
    if(!data[i][cId]) continue;
    if(String(data[i][cT]||'').trim()===''){ var t=Utilities.getUuid().replace(/-/g,''); sh.getRange(i+1,cT+1).setValue(t); data[i][cT]=t; n++; }
    Logger.log(data[i][cId]+' | '+(cS>=0?data[i][cS]:'')+' | '+(cE>=0?data[i][cE]:'')+' | '+base+'/register.html?id='+data[i][cId]+'&token='+data[i][cT]);
  }
  Logger.log('=== '+n+'件に token を補完 ===');
}

function diagnoseHeaders() {
  var sh = SpreadsheetApp.getActive().getSheetByName('reservations');
  var actual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log('実ヘッダー(' + actual.length + '列): ' + actual.join(' | '));
  Logger.log('token列の位置: col' + (actual.indexOf('token') + 1));
  Logger.log('期待(HEADERS_RESERVATIONS): ' + HEADERS_RESERVATIONS.join(' | '));
  Logger.log('generateToken_() の出力: "' + generateToken_() + '"');
  var last = sh.getLastRow(), tc = actual.indexOf('token') + 1;
  Logger.log('最新行(テスト予約)のtoken列の値: "' + (tc > 0 ? sh.getRange(last, tc).getValue() : 'token列が無い') + '"');
}

