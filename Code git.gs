/**
 * IICSA Summer School – Mini Enrolment & Payment Portal
 * Google Apps Script web app + Google Sheet + Square Payment Links
 *
 * Flow:
 *  1. Parent fills the form (parent details + any number of children).
 *  2. Server re-calculates the total (sibling discount), creates a reference ID,
 *     saves a "Pending" row, creates a Square payment link for that exact amount,
 *     and emails the parent the reference + link.
 *  3. Parent pays on Square -> Square redirects back to this portal (?ref=...).
 *  4. Portal checks the Square order; if paid, the row becomes "Paid" and a
 *     confirmation email is sent. A 10-minute background check also catches
 *     parents who close the tab before being redirected.
 *
 * Secrets are NOT stored in this file. Put them in
 *   Project Settings -> Script Properties:
 *     SQUARE_ACCESS_TOKEN   (from developer.squareup.com -> your app -> Credentials)
 *     SQUARE_LOCATION_ID    (same page -> Locations)
 *     SQUARE_ENV            sandbox  or  production
 */

// ======================= EDIT THESE =======================
const CONFIG = {
  SCHOOL_NAME: 'IICSA Summer School',
  TERM_LABEL: 'STAR Morning Program 2026',        // shown on the form + Square checkout
  CURRENCY: 'AUD',
  FIRST_CHILD_FEE: 120,                         // STAR core program – first child
  EXTRA_CHILD_FEE: 100,                         // STAR core program – each additional child
  ADDON_FEE: 100,                               // Youth Add-on (1–2 pm) – per child, optional
  MAX_CHILDREN: 8,
  CAMPUSES: ["Iqra College – O'Halloran Hill"],   // one entry = campus question is hidden
  MAX_AGE: 18,                                  // children must be under this age on the day they apply
  ADMIN_EMAIL: '',                              // optional: gets a copy of every paid registration
  SENDER_NAME: 'Summer School (STAR)',           // name parents see in their inbox
  REPLY_TO: 'schoolsummer46@gmail.com',        // parents' replies go here
  REF_PREFIX: 'IICSA-',
  PUBLIC_URL: 'https://iicsastar.me/',                             // your GitHub page, e.g. 'https://yourname.github.io/star/' (leave '' to use the Google link)

  // ---- Welcome / confirmation email (sent once payment is confirmed) ----
  PROGRAM_NAME: 'IICSA STAR Summer School Program 2026/27',
  PROGRAM_START: '2026-12-14',                  // STAR group is worked out from each child's age on this day
  WHATSAPP_LINK: 'https://chat.whatsapp.com/LMWEJ1QfnegG1O5XffIHMT',
  BOOKLET_LINK: 'https://drive.google.com/file/d/1trz3bWNtzVumugUWYT8aIJLLbde07Z5O/view?usp=sharing',                             // Google Drive share link of the booklet PDF (shown as a button)
  CONTACT_PHONE: '0481 312 394',
  CONTACT_EMAIL: 'schoolsummer46@gmail.com',
  SIGN_OFF: 'Latessa<br>IICSA Summer School Team',
  AGE_GROUPS: [                                 // [lowest age, highest age, group name]
    [5, 6, 'STAR 1A'],
    [7, 8, 'STAR 1B'],
    [9, 11, 'STAR 2'],
    [12, 99, 'STAR 3'],
  ],
};
// ==========================================================

const SHEET_REG = 'Registrations';
const SHEET_KIDS = 'Children';
const REG_HEADERS = ['Reference', 'Submitted', 'Parent name', 'Email', 'Phone', 'Campus',
  'No. of children', 'Youth add-ons', 'Children', 'Total (AUD)', 'Status', 'Paid at',
  'Square order ID', 'Square payment ID', 'Payment link', 'Notes'];
const KID_HEADERS = ['Reference', 'Child name', 'Date of birth', 'Age', 'Youth add-on',
  'Campus', 'Parent name', 'Parent email', 'Status'];   // keep Status last
const COL = Object.fromEntries(REG_HEADERS.map((h, i) => [h, i + 1]));

// ---------------------------------------------------------------- web app
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.config = JSON.stringify(publicConfig_());
  t.ref = JSON.stringify((e && e.parameter && e.parameter.ref) || '');
  t.testpay = JSON.stringify(!!(e && e.parameter && e.parameter.testpay === '1'));
  return t.evaluate()
    .setTitle(CONFIG.SCHOOL_NAME + ' – Enrolment')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // allows the GitHub page to show the portal
}

function publicConfig_() {
  return {
    schoolName: CONFIG.SCHOOL_NAME, term: CONFIG.TERM_LABEL, currency: CONFIG.CURRENCY,
    testMode: mode_() === 'test',
    firstFee: CONFIG.FIRST_CHILD_FEE, extraFee: CONFIG.EXTRA_CHILD_FEE, addonFee: CONFIG.ADDON_FEE,
    maxChildren: CONFIG.MAX_CHILDREN, campuses: CONFIG.CAMPUSES, maxAge: CONFIG.MAX_AGE,
  };
}

// ---------------------------------------------------------------- pricing
function calcTotal_(children) {
  const n = children.length;
  if (n < 1) return 0;
  const addOns = children.filter(c => c.addOn).length;
  return CONFIG.FIRST_CHILD_FEE + (n - 1) * CONFIG.EXTRA_CHILD_FEE + addOns * CONFIG.ADDON_FEE;
}

// ---------------------------------------------------------------- submit
/** Called from the page. Returns {ref, total, url}. */
function submitRegistration(form) {
  const data = validate_(form);
  const total = calcTotal_(data.children);   // never trust the browser's total

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let ref, rowIndex;
  try {
    const sh = sheet_(SHEET_REG, REG_HEADERS);
    ref = newRef_(sh);
    sh.appendRow([ref, new Date(), data.parentName, data.email, "'" + data.phone, data.campus,
      data.children.length, data.children.filter(c => c.addOn).length,
      data.children.map(c => `${c.name} (age ${c.age}${c.addOn ? ' + Youth add-on' : ''})`).join('; '),
      total, 'Pending', '', '', '', '', data.notes]);
    rowIndex = sh.getLastRow();
    const kids = sheet_(SHEET_KIDS, KID_HEADERS);
    data.children.forEach(c => kids.appendRow([ref, c.name, c.dob, c.age, c.addOn ? 'Yes' : 'No', data.campus,
      data.parentName, data.email, 'Pending']));
  } finally {
    lock.releaseLock();
  }

  const link = createPaymentLink_(ref, total, data);
  const sh = sheet_(SHEET_REG, REG_HEADERS);
  sh.getRange(rowIndex, COL['Square order ID']).setValue(link.orderId);
  sh.getRange(rowIndex, COL['Payment link']).setValue(link.url);

  // An email failure (e.g. daily Gmail limit reached) must not stop the registration –
  // the parent still sees the reference and Pay button on screen.
  try { sendReferenceEmail_(data, ref, total, link.url); }
  catch (err) { console.error('Reference email failed for ' + ref + ': ' + err.message); }
  return { ref: ref, total: total, url: link.url };
}

function validate_(f) {
  const clean = s => String(s == null ? '' : s).trim().slice(0, 200);
  const d = {
    parentName: clean(f.parentName), email: clean(f.email).toLowerCase(),
    phone: clean(f.phone), campus: clean(f.campus), notes: clean(f.notes),
    children: (Array.isArray(f.children) ? f.children : []).map(c => ({
      name: clean(c.name), dob: clean(c.dob), addOn: c.addOn === true,
    })).filter(c => c.name),
  };
  if (!d.parentName) throw new Error('Please enter the parent/guardian name.');
  if (!EMAIL_RE.test(d.email) || d.email.indexOf('..') >= 0) throw new Error('Please enter a valid email address.');
  d.phone = auPhone_(d.phone);
  if (!d.phone) throw new Error('Please enter a valid Australian phone number (e.g. 0412 345 678 or +61 412 345 678).');
  const today = Utilities.formatDate(new Date(), 'Australia/Adelaide', 'yyyy-MM-dd');
  const oldest = (Number(today.slice(0, 4)) - CONFIG.MAX_AGE) + today.slice(4);
  d.children.forEach((c, i) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(c.dob) || c.dob > today || c.dob <= oldest)
      throw new Error(`Child ${i + 1}: date of birth must be a real date and the child must be under ${CONFIG.MAX_AGE}.`);
    c.age = ageOn_(c.dob, today);
  });
  if (CONFIG.CAMPUSES.indexOf(d.campus) < 0) throw new Error('Please choose a campus.');
  if (!d.children.length) throw new Error('Please add at least one child.');
  if (d.children.length > CONFIG.MAX_CHILDREN) throw new Error('Too many children on one form.');
  if (!f.consent) throw new Error('Please tick the agreement box.');
  return d;
}

const EMAIL_RE = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;

/** Returns a tidy Australian number, or '' if invalid. Accepts 04xx…, 0[2378]…, +61…, 61… */
function auPhone_(raw) {
  let n = String(raw).replace(/[\s\-().]/g, '');
  if (n.indexOf('+61') === 0) n = '0' + n.slice(3);
  else if (/^61\d{9}$/.test(n)) n = '0' + n.slice(2);
  if (/^0[2-478]\d{8}$/.test(n) === false) return '';
  return n[1] === '4' ? `${n.slice(0, 4)} ${n.slice(4, 7)} ${n.slice(7)}` : `(${n.slice(0, 2)}) ${n.slice(2, 6)} ${n.slice(6)}`;
}

function ageOn_(dob, today) {
  let a = Number(today.slice(0, 4)) - Number(dob.slice(0, 4));
  if (today.slice(5) < dob.slice(5)) a--;
  return a;
}

function newRef_(sh) {
  const alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';     // no 0/O/1/I/L confusion
  const existing = new Set(sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : []);
  for (let i = 0; i < 50; i++) {
    let s = CONFIG.REF_PREFIX;
    for (let j = 0; j < 5; j++) s += alphabet[Math.floor(Math.random() * alphabet.length)];
    if (!existing.has(s)) return s;
  }
  throw new Error('Could not create a reference – please try again.');
}

// ---------------------------------------------------------------- Square
function square_(method, path, body) {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty('SQUARE_ACCESS_TOKEN');
  if (!token) throw new Error('Portal not configured yet (missing Square token).');
  const base = (p.getProperty('SQUARE_ENV') || 'sandbox') === 'production'
    ? 'https://connect.squareup.com' : 'https://connect.squareupsandbox.com';
  const res = UrlFetchApp.fetch(base + path, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, 'Square-Version': '2025-01-23' },
    payload: body ? JSON.stringify(body) : undefined,
  });
  const json = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) {
    console.error(path, res.getContentText());
    throw new Error('Square error: ' + ((json.errors && json.errors[0] && json.errors[0].detail) || res.getResponseCode()));
  }
  return json;
}

/**
 * 'test'       – no Square token yet (or SQUARE_ENV = test): payments are simulated, no money moves
 * 'sandbox'    – Square test environment (test cards)
 * 'production' – real payments
 */
function mode_() {
  const p = PropertiesService.getScriptProperties();
  const env = (p.getProperty('SQUARE_ENV') || '').trim().toLowerCase();
  if (!p.getProperty('SQUARE_ACCESS_TOKEN') || env === 'test') return 'test';
  return env === 'production' ? 'production' : 'sandbox';
}

/** Called from the pretend checkout page in test mode only. */
function testPay(ref) {
  if (mode_() !== 'test') throw new Error('Test payments are switched off – real Square payments are active.');
  ref = String(ref || '').trim().toUpperCase();
  const sh = sheet_(SHEET_REG, REG_HEADERS);
  const row = findRow_(sh, ref);
  if (!row) throw new Error('Reference not found.');
  const v = sh.getRange(row, 1, 1, REG_HEADERS.length).getValues()[0];
  if (v[COL.Status - 1] !== 'Paid') {
    sh.getRange(row, COL.Status).setValue('Paid');
    sh.getRange(row, COL['Paid at']).setValue(new Date());
    sh.getRange(row, COL['Square payment ID']).setValue('TEST-PAYMENT');
    markChildren_(ref, 'Paid');
    try { sendPaidEmail_(v); } catch (err) { console.error('Paid email failed for ' + ref + ': ' + err.message); }
  }
  return checkStatus(ref);
}

function createPaymentLink_(ref, total, data) {
  if (mode_() === 'test') {
    return { url: publicUrl_() + '?ref=' + encodeURIComponent(ref) + '&testpay=1', orderId: 'TEST-' + ref };
  }
  const locationId = PropertiesService.getScriptProperties().getProperty('SQUARE_LOCATION_ID');
  const r = square_('post', '/v2/online-checkout/payment-links', {
    idempotency_key: ref + '-' + Utilities.getUuid(),
    quick_pay: {
      name: `${CONFIG.TERM_LABEL} – ${ref} (${data.children.length} child${data.children.length > 1 ? 'ren' : ''})`,
      price_money: { amount: Math.round(total * 100), currency: CONFIG.CURRENCY },
      location_id: locationId,
    },
    payment_note: ref,
    pre_populated_data: { buyer_email: data.email },
    checkout_options: { redirect_url: publicUrl_() + '?ref=' + encodeURIComponent(ref) },
  });
  return { url: r.payment_link.url, orderId: r.payment_link.order_id };
}

/** Returns {paid:boolean, paymentId:string} for a Square order. */
function orderPaid_(orderId) {
  const o = square_('get', '/v2/orders/' + encodeURIComponent(orderId)).order || {};
  const tenders = o.tenders || [];
  const due = o.net_amount_due_money ? o.net_amount_due_money.amount : null;
  const paid = o.state === 'COMPLETED' || (tenders.length > 0 && (due === null || due === 0));
  return { paid: paid, paymentId: tenders.length ? tenders[0].payment_id || tenders[0].id : '' };
}

// ---------------------------------------------------------------- status
/** Called from the page. Looks up a reference and refreshes it from Square if still pending. */
function checkStatus(ref) {
  ref = String(ref || '').trim().toUpperCase();
  if (ref && ref.indexOf(CONFIG.REF_PREFIX) !== 0) ref = CONFIG.REF_PREFIX + ref;
  const sh = sheet_(SHEET_REG, REG_HEADERS);
  const row = findRow_(sh, ref);
  if (!row) return { found: false, ref: ref };
  let v = sh.getRange(row, 1, 1, REG_HEADERS.length).getValues()[0];
  if (v[COL.Status - 1] === 'Pending' && v[COL['Square order ID'] - 1] && !String(v[COL['Square order ID'] - 1]).startsWith('TEST-')) {
    try {
      refreshRow_(sh, row);
    } catch (err) {
      // e.g. the order was made under different Square keys – show it as pending instead of an error
      console.error('Status check failed for ' + ref + ': ' + err.message);
    }
    v = sh.getRange(row, 1, 1, REG_HEADERS.length).getValues()[0];
  }
  const paid = v[COL.Status - 1] === 'Paid';
  const paidAt = v[COL['Paid at'] - 1];
  return {
    found: true, ref: ref, status: v[COL.Status - 1], total: v[COL['Total (AUD)'] - 1],
    children: v[COL['No. of children'] - 1],
    childList: String(v[COL.Children - 1]).split('; '),
    campus: v[COL.Campus - 1],
    parentFirst: String(v[COL['Parent name'] - 1]).split(' ')[0],
    email: maskEmail_(v[COL.Email - 1]),
    paidAt: paid && paidAt ? Utilities.formatDate(new Date(paidAt), 'Australia/Adelaide', 'd MMM yyyy, h:mm a') : '',
    receipt: paid && !String(v[COL['Square payment ID'] - 1]).startsWith('TEST') ? String(v[COL['Square payment ID'] - 1]).slice(-8).toUpperCase() : '',
    url: paid ? '' : v[COL['Payment link'] - 1],
  };
}

function refreshRow_(sh, row) {
  const v = sh.getRange(row, 1, 1, REG_HEADERS.length).getValues()[0];
  const r = orderPaid_(v[COL['Square order ID'] - 1]);
  if (!r.paid) return false;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (sh.getRange(row, COL.Status).getValue() === 'Paid') return true;   // already done
    sh.getRange(row, COL.Status).setValue('Paid');
    sh.getRange(row, COL['Paid at']).setValue(new Date());
    sh.getRange(row, COL['Square payment ID']).setValue(r.paymentId);
    markChildren_(v[0], 'Paid');
  } finally {
    lock.releaseLock();
  }
  try { sendPaidEmail_(v); } catch (err) { console.error('Paid email failed for ' + v[0] + ': ' + err.message); }
  return true;
}

function markChildren_(ref, status) {
  const kids = sheet_(SHEET_KIDS, KID_HEADERS);
  if (kids.getLastRow() < 2) return;
  const refs = kids.getRange(2, 1, kids.getLastRow() - 1, 1).getValues();
  refs.forEach((r, i) => { if (r[0] === ref) kids.getRange(i + 2, KID_HEADERS.length).setValue(status); });
}

/** Background job (every 10 min) – catches parents who paid but closed the tab. */
function checkPendingPayments() {
  const sh = sheet_(SHEET_REG, REG_HEADERS);
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, REG_HEADERS.length).getValues();
  const cutoff = Date.now() - 1000 * 60 * 60 * 24 * 30;              // look back 30 days
  rows.forEach((v, i) => {
    if (v[COL.Status - 1] !== 'Pending' || !v[COL['Square order ID'] - 1]) return;
    if (String(v[COL['Square order ID'] - 1]).startsWith('TEST-')) return;
    if (new Date(v[COL.Submitted - 1]).getTime() < cutoff) return;
    try { refreshRow_(sh, i + 2); } catch (err) { console.error(v[0], err); }
  });
}

// ---------------------------------------------------------------- email
function sendReferenceEmail_(d, ref, total, url) {
  const kids = d.children.map(c => `<li>${esc_(c.name)} (age ${c.age}) – STAR Morning Program${c.addOn ? ' + Youth Add-on (1–2 pm)' : ''}</li>`).join('');
  MailApp.sendEmail({ ...mailOpts_(),
    to: d.email,
    subject: `${CONFIG.SCHOOL_NAME} – your enrolment reference ${ref}`,
    htmlBody: `<p>Assalamu alaikum ${esc_(d.parentName)},</p>
      <p>Thank you for registering for the <b>${esc_(CONFIG.TERM_LABEL)}</b> at ${esc_(d.campus)}.</p>
      <p style="font-size:18px">Your reference ID: <b>${ref}</b><br>Amount due: <b>$${total.toFixed(2)} ${CONFIG.CURRENCY}</b></p>
      <ul>${kids}</ul>
      <p><a href="${url}" style="background:#0f6b4f;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none">Pay now</a></p>
      <p>Your place is confirmed once payment is received. Fees are due one week before the program begins (Monday 7 December 2026) unless other arrangements are made with admin. If you have already paid, you can ignore the button.</p>
      <p>Check your status anytime: <a href="${publicUrl_()}?ref=${ref}">${publicUrl_()}?ref=${ref}</a></p>
      <p>Jazakum Allahu khairan,<br>${esc_(CONFIG.SCHOOL_NAME)}</p>`,
  });
}

function sendPaidEmail_(v) {
  const ref = v[0], email = v[COL.Email - 1];
  const html = welcomeEmailHtml_(v);
  MailApp.sendEmail({ ...mailOpts_(), to: email, htmlBody: html,
    subject: `Welcome to STAR Summer School – enrolment confirmed (${ref})` });
  if (CONFIG.ADMIN_EMAIL) {
    MailApp.sendEmail({ ...mailOpts_(), to: CONFIG.ADMIN_EMAIL, subject: `PAID: ${ref} – ${v[COL['Parent name'] - 1]}`, htmlBody: html });
  }
}

/** Builds the combined "payment confirmed + welcome" email for one registration row. */
function welcomeEmailHtml_(v) {
  const ref = v[0];
  const kids = kidsFor_(ref);
  const anyAddon = kids.some(k => k.addOn);
  const total = Number(v[COL['Total (AUD)'] - 1]).toFixed(2);
  const bookletUrl = CONFIG.BOOKLET_LINK;
  const P = 'margin:0 0 14px;line-height:1.55';

  const kidRows = kids.map(k => {
    const g = starGroup_(k.dob);
    return `<tr>
      <td style="padding:8px 0;border-bottom:1px solid #eee"><b>${esc_(k.name)}</b>${g.age != null ? ` <span style="color:#777">(age ${g.age})</span>` : ''}
        <br><span style="color:#777;font-size:13px">STAR Program${k.addOn ? ' + Youth Add-on' : ''}</span></td>
      <td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;white-space:nowrap">
        ${g.name ? `Join <b style="color:#0f6b4f">${g.name}</b>` : `<span style="color:#b3261e">Please contact us</span>`}</td></tr>`;
  }).join('');

  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#2f3550;max-width:580px;margin:0 auto">

  <div style="background:#e9f7f1;border-radius:12px;padding:16px 18px;margin-bottom:20px">
    <div style="font-size:13px;color:#0f6b4f;font-weight:bold;letter-spacing:.5px">PAYMENT CONFIRMED ✓</div>
    <div style="font-size:15px;margin-top:6px">Reference <b>${ref}</b> · Paid <b>$${total} ${CONFIG.CURRENCY}</b></div>
  </div>

  <p style="${P}">Assalamu alaikum dear parents,</p>
  <p style="${P}">We are excited to welcome your child to the <b>${esc_(CONFIG.PROGRAM_NAME)}</b>. Alhamdulillah, we are looking forward to a wonderful three weeks of learning, creativity, exploration, and Islamic values.${anyAddon ? ` As you have also enrolled in the <b>Youth Program Add On</b>, we are excited to engage with the students in some fun and interesting activities.` : ''}</p>
  <p style="${P}">Please note for this year, we will not be having an operational canteen but will have some packaged items to purchase.${anyAddon ? ` As you are enrolled for the full day, please provide a full lunch as you would for a school day.` : ''}</p>

  <h3 style="font-size:17px;margin:22px 0 8px">1. Join our WhatsApp Community</h3>
  <p style="${P}">To ensure that all parents receive important updates, daily reminders, excursion details, and announcements, we kindly ask you to join our official Summer School WhatsApp Community:</p>
  <p style="margin:0 0 16px"><a href="${CONFIG.WHATSAPP_LINK}" style="display:inline-block;background:#25d366;color:#fff;font-weight:bold;padding:12px 22px;border-radius:24px;text-decoration:none">Join the WhatsApp Community</a></p>

  <h3 style="font-size:17px;margin:22px 0 8px">2. Join your child's STAR group</h3>
  <p style="${P}">Once you enter the community, you will see sub-groups for each STAR age category. Based on your enrolment, please join:</p>
  <table style="width:100%;border-collapse:collapse;margin:0 0 12px">${kidRows}</table>
  <p style="margin:0 0 14px;font-size:13px;color:#777">STAR 1A: ages 5–6 · STAR 1B: ages 7–8 · STAR 2: ages 9–11 · STAR 3: ages 12+</p>
  <p style="${P}">Joining the correct sub-group ensures you receive only the relevant updates for your child's class — including excursion schedules, daily reminders, messages from teachers, pick-up/drop-off alerts, and any urgent notices. Please join as soon as possible so you do not miss any information before the program begins.</p>

  ${bookletUrl ? `<h3 style="font-size:17px;margin:22px 0 8px">3. Read the information booklet</h3>
  <p style="${P}">Please read our parent information booklet. It covers dates, pick-up, what to bring, and more.</p>
  <p style="margin:0 0 16px"><a href="${bookletUrl}" style="display:inline-block;background:#0f6b4f;color:#fff;font-weight:bold;padding:12px 22px;border-radius:24px;text-decoration:none">View the information booklet</a></p>` : ''}

  <div style="border-top:1px solid #eee;margin-top:22px;padding-top:16px">
    <p style="${P}">If you have any questions, feel free to contact us:<br>
      📞 <a href="tel:${CONFIG.CONTACT_PHONE.replace(/\s/g, '')}" style="color:#2f3550">${CONFIG.CONTACT_PHONE}</a><br>
      📧 <a href="mailto:${CONFIG.CONTACT_EMAIL}" style="color:#2f3550">${CONFIG.CONTACT_EMAIL}</a></p>
    <p style="${P}">We look forward to a beautiful, beneficial, and memorable Summer School experience for your child, in shā’ Allāh.</p>
    <p style="${P}">Warm regards,<br>${CONFIG.SIGN_OFF}</p>
  </div>
</div>`;
}

/** Children of one registration, read from the Children tab: [{name, dob:'yyyy-mm-dd', addOn}] */
function kidsFor_(ref) {
  const sh = sheet_(SHEET_KIDS, KID_HEADERS);
  if (sh.getLastRow() < 2) return [];
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const col = h => KID_HEADERS.indexOf(h);
  return sh.getRange(2, 1, sh.getLastRow() - 1, KID_HEADERS.length).getValues()
    .filter(r => r[0] === ref)
    .map(r => {
      const d = r[col('Date of birth')];
      return { name: r[col('Child name')], addOn: r[col('Youth add-on')] === 'Yes',
        dob: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d) };
    });
}

/** STAR group from date of birth, using the child's age on the first day of the program. */
function starGroup_(dob) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dob)) return { age: null, name: '' };
  const age = ageOn_(dob, CONFIG.PROGRAM_START);
  const g = CONFIG.AGE_GROUPS.find(([lo, hi]) => age >= lo && age <= hi);
  return { age: age, name: g ? g[2] : '' };
}

/** Run this from the editor to preview the welcome email: sends a sample to CONTACT_EMAIL using the latest registration. */
function sendTestWelcomeEmail() {
  const sh = sheet_(SHEET_REG, REG_HEADERS);
  if (sh.getLastRow() < 2) throw new Error('Make one test registration first.');
  const v = sh.getRange(sh.getLastRow(), 1, 1, REG_HEADERS.length).getValues()[0];
  MailApp.sendEmail({ ...mailOpts_(), to: CONFIG.CONTACT_EMAIL, htmlBody: welcomeEmailHtml_(v),
    subject: `[TEST] Welcome to STAR Summer School – enrolment confirmed (${v[0]})` });
  console.log('Test email sent to ' + CONFIG.CONTACT_EMAIL + (CONFIG.BOOKLET_LINK ? '.' : ' – booklet button hidden because BOOKLET_LINK is empty.'));
}

// ---------------------------------------------------------------- helpers
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e6f2ee');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow_(sh, ref) {
  if (sh.getLastRow() < 2) return 0;
  const refs = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = refs.length - 1; i >= 0; i--) if (refs[i][0] === ref) return i + 2;
  return 0;
}

/** The address parents use: the GitHub page if set, otherwise the Google web app link. */
function publicUrl_() {
  return (CONFIG.PUBLIC_URL || ScriptApp.getService().getUrl()).replace(/\?.*$/, '');
}

function maskEmail_(e) {
  const [u, d] = String(e).split('@');
  return d ? u.slice(0, 2) + '•••@' + d : '';
}

function mailOpts_() {
  const o = { name: CONFIG.SENDER_NAME };
  if (CONFIG.REPLY_TO) o.replyTo = CONFIG.REPLY_TO;
  return o;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---------------------------------------------------------------- one-time setup
/** Run once from the editor: creates the sheets and the 10-minute background check. */
function setup() {
  sheet_(SHEET_REG, REG_HEADERS);
  sheet_(SHEET_KIDS, KID_HEADERS);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkPendingPayments')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkPendingPayments').timeBased().everyMinutes(10).create();
  const p = PropertiesService.getScriptProperties();
  console.log('Setup done. Payment mode: ' + mode_().toUpperCase() +
    (mode_() === 'test' ? ' (no Square keys yet – payments are simulated)' : ''));
}
