// ═══════════════════════════════════════════════════════════════════════════
// MSRC 2026 – Google Apps Script Backend
// ─────────────────────────────────────────────────────────────────────────
// SETUP INSTRUCTIONS:
//   1. Create a new Google Spreadsheet and copy its ID from the URL.
//   2. Open Extensions → Apps Script, paste this code.
//   3. In Apps Script: Project Settings → Script Properties, add:
//        SPREADSHEET_ID  → <your sheet ID>
//        ADMIN_EMAIL     → <your admin email>
//        FRONTEND_URL    → <your deployed frontend URL>
//        DRIVE_FOLDER_ID → <optional: Google Drive folder ID for file uploads>
//   4. Deploy → New Deployment → Web App
//        Execute as: Me
//        Who has access: Anyone
//   5. Copy the deployment URL and paste it as API_URL in index.html.
//   6. Open your Sheet URL → ?action=setup to initialise all sheets.
// ═══════════════════════════════════════════════════════════════════════════

const SS_ID             = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ADMIN_EMAIL       = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail();
const BASE_URL          = PropertiesService.getScriptProperties().getProperty('FRONTEND_URL') || '';
const DRIVE_FOLDER_ID   = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID') || '';
const CONF_NAME         = 'Medical Students Research Conference 2026';
const PASSWORD_SALT     = 'MSRC_2026_SECURE_SALT';

// ── Sheet Names ──────────────────────────────────────────────────────────────
const SH = {
  USERS:        'Users',
  SESSIONS:     'Sessions',
  ABSTRACTS:    'Abstracts',
  CO_AUTHORS:   'CoAuthors',
  EVALUATIONS:  'Evaluations',
  ASSIGNMENTS:  'Assignments',
  EVAL_APPS:    'EvalApplications',
  REGISTRATIONS:'Registrations',
  DISC_CODES:   'DiscountCodes',
  CERTIFICATES: 'Certificates',
  LOGS:         'AdminLogs'
};

// Sheet column definitions
const COLUMNS = {
  Users:            ['id','full_name','email','password_hash','phone','university','academic_level','region','nationality','gender','roles','email_verified','reset_token','reset_expiry','created_at'],
  Sessions:         ['token','user_id','expires_at'],
  Abstracts:        ['id','user_id','title','type','specialty','background','methods','results','conclusion','iban','coi','coi_text','status','track','file_url','photo_url','avg_score','submitted_at'],
  CoAuthors:        ['id','abstract_id','name','email','university'],
  Evaluations:      ['id','abstract_id','evaluator_id','originality','methodology','clarity','relevance','avg_score','recommendation','track_assigned','comments','submitted_at'],
  Assignments:      ['id','evaluator_id','abstract_id','assigned_at','eval_status'],
  EvalApplications: ['id','user_id','academic_level','specialty','years_exp','prev_exp','prev_exp_detail','cv_url','status','applied_at'],
  Registrations:    ['id','user_id','phone','national_id','workshop','discount_code','base_amount','discount_amount','total_amount','payment_status','payment_ref','mode','registered_at'],
  DiscountCodes:    ['code','type','amount','max_uses','used_count','active'],
  Certificates:     ['id','user_id','cert_id','cert_type','issued_at'],
  AdminLogs:        ['id','admin_id','action','target','details','timestamp']
};

// ════════════════════════════════════════════════════════════════════════════
// CORE HELPERS
// ════════════════════════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const cols = COLUMNS[name];
    if (cols) {
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
      sheet.getRange(1, 1, 1, cols.length)
        .setBackground('#0f2444').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function uuid() { return Utilities.getUuid(); }

function hashPassword(password) {
  const raw = password + PASSWORD_SALT;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
}

function cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

function ok(data) {
  return cors(ContentService.createTextOutput(JSON.stringify({ ok: true, ...data })));
}

function fail(msg, code) {
  return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg, code: code || 400 })));
}

// ── Sheet Read/Write ─────────────────────────────────────────────────────────

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(String);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRow(sheet, colName, value) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { row: null, rowIndex: -1 };
  const headers = data[0].map(String);
  const colIdx = headers.indexOf(colName);
  if (colIdx === -1) return { row: null, rowIndex: -1 };
  const lv = String(value).toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx]).toLowerCase() === lv) {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = data[i][j]; });
      return { row: obj, rowIndex: i + 1 };
    }
  }
  return { row: null, rowIndex: -1 };
}

function updateRow(sheet, rowIndex, updates) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const row = [...data[rowIndex - 1]];
  Object.keys(updates).forEach(key => {
    const idx = headers.indexOf(key);
    if (idx !== -1) row[idx] = updates[key];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function appendRow(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const row = headers.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : '');
  sheet.appendRow(row);
}

// ════════════════════════════════════════════════════════════════════════════
// SESSION MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════

function createSession(userId) {
  const token = generateToken();
  const expires = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000); // 7 days
  const sheet = getSheet(SH.SESSIONS);
  appendRow(sheet, { token, user_id: userId, expires_at: expires.toISOString() });
  // Prune old sessions periodically
  try { pruneOldSessions(sheet); } catch(e) {}
  return token;
}

function pruneOldSessions(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const expiryIdx = data[0].indexOf('expires_at');
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    if (new Date(data[i][expiryIdx]) < now) sheet.deleteRow(i + 1);
  }
}

function getUserFromToken(token) {
  if (!token || token.length < 10) return null;
  const { row: sess } = findRow(getSheet(SH.SESSIONS), 'token', token);
  if (!sess) return null;
  if (new Date(sess.expires_at) < new Date()) return null;
  const { row: user } = findRow(getSheet(SH.USERS), 'id', sess.user_id);
  return user || null;
}

function requireAuth(token) {
  const user = getUserFromToken(token);
  if (!user) throw new Error('UNAUTHORIZED');
  return user;
}

function hasRole(user, ...roles) {
  const userRoles = String(user.roles).split(',').map(r => r.trim());
  return roles.some(r => userRoles.includes(r));
}

function requireRole(user, ...roles) {
  if (!hasRole(user, ...roles)) throw new Error('FORBIDDEN');
}

// ════════════════════════════════════════════════════════════════════════════
// HTTP ENTRY POINTS
// ════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const p = e.parameter || {};
    const token = p.token || '';
    switch (p.action) {
      case 'setup':                return handleSetup();
      case 'get_profile':          return handleGetProfile(token);
      case 'get_my_abstracts':     return handleGetMyAbstracts(token);
      case 'get_assignments':      return handleGetAssignments(token);
      case 'get_eval_status':      return handleGetEvalStatus(token);
      case 'get_my_certs':         return handleGetMyCerts(token);
      case 'verify_cert':          return handleVerifyCert(p.cert_id);
      case 'admin_stats':          return handleAdminStats(token);
      case 'admin_abstracts':      return handleAdminAbstracts(token, p);
      case 'admin_evaluator_apps': return handleAdminEvalApps(token);
      case 'admin_registrations':  return handleAdminRegistrations(token);
      case 'admin_users':          return handleAdminUsers(token);
      case 'export_csv':           return handleExportCSV(token, p.type);
      default:                     return fail('Unknown GET action');
    }
  } catch (ex) {
    if (ex.message === 'UNAUTHORIZED') return fail('Unauthorized', 401);
    if (ex.message === 'FORBIDDEN')    return fail('Forbidden', 403);
    return fail(ex.message);
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return fail('No body');
    const body = JSON.parse(e.postData.contents);
    const token = body.token || '';
    switch (body.action) {
      case 'register':               return handleRegister(body);
      case 'login':                  return handleLogin(body);
      case 'logout':                 return handleLogout(token);
      case 'forgot_password':        return handleForgotPassword(body);
      case 'reset_password':         return handleResetPassword(body);
      case 'update_profile':         return handleUpdateProfile(token, body);
      case 'submit_abstract':        return handleSubmitAbstract(token, body);
      case 'apply_evaluator':        return handleApplyEvaluator(token, body);
      case 'submit_evaluation':      return handleSubmitEvaluation(token, body);
      case 'verify_discount':        return handleVerifyDiscount(body);
      case 'register_conference':    return handleRegisterConference(token, body);
      case 'admin_update_abstract':  return handleAdminUpdateAbstract(token, body);
      case 'admin_process_eval_app': return handleAdminProcessEvalApp(token, body);
      case 'admin_assign_abstracts': return handleAdminAssignAbstracts(token, body);
      case 'admin_add_onsite':       return handleAdminAddOnsite(token, body);
      case 'admin_update_user_role': return handleAdminUpdateUserRole(token, body);
      case 'admin_issue_cert':       return handleAdminIssueCert(token, body);
      case 'admin_bulk_certs':       return handleAdminBulkCerts(token);
      default:                       return fail('Unknown POST action');
    }
  } catch (ex) {
    if (ex.message === 'UNAUTHORIZED') return fail('Unauthorized', 401);
    if (ex.message === 'FORBIDDEN')    return fail('Forbidden', 403);
    return fail(ex.message);
  }
}

// ════════════════════════════════════════════════════════════════════════════
// AUTH
// ════════════════════════════════════════════════════════════════════════════

function handleRegister(body) {
  const { full_name, email, password, phone, university, university_name,
          academic_level, region, nationality, gender } = body;

  if (!full_name || !email || !password || !phone || !university || !academic_level || !region || !nationality)
    return fail('Missing required fields');
  if (password.length < 8)
    return fail('Password must be at least 8 characters');

  const sheet = getSheet(SH.USERS);
  if (findRow(sheet, 'email', email).row)
    return fail('This email is already registered');

  const uni = university === 'non_kau' ? (university_name || 'Other University') : 'King Abdulaziz University (KAU)';

  const user = {
    id: uuid(), full_name, email,
    password_hash: hashPassword(password),
    phone, university: uni, academic_level, region,
    nationality, gender: gender || '',
    roles: 'delegate',
    email_verified: 'yes',
    reset_token: '', reset_expiry: '',
    created_at: new Date().toISOString()
  };
  appendRow(sheet, user);

  sendEmail(email, `Welcome to ${CONF_NAME}!`,
    `<h2>Welcome, ${full_name}!</h2>
     <p>Your MSRC 2026 account has been created. You can now log in to the portal and submit your abstract, register for the conference, and more.</p>
     <p style="margin-top:16px"><a href="${BASE_URL}" style="background:#0f2444;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none">Go to Portal →</a></p>`
  );
  sendEmail(ADMIN_EMAIL, `New User – ${full_name}`,
    `<p>New registration: <b>${full_name}</b> (${email}), ${university}, ${academic_level}, ${region}</p>`
  );

  return ok({ message: 'Account created successfully' });
}

function handleLogin(body) {
  const { email, password } = body;
  if (!email || !password) return fail('Email and password required');

  const { row: user } = findRow(getSheet(SH.USERS), 'email', email);
  if (!user || user.password_hash !== hashPassword(password))
    return fail('Invalid email or password');

  const token = createSession(user.id);
  const safe = sanitizeUser(user);
  return ok({ token, user: safe });
}

function handleLogout(token) {
  if (token) {
    const sheet = getSheet(SH.SESSIONS);
    const { rowIndex } = findRow(sheet, 'token', token);
    if (rowIndex > 0) sheet.deleteRow(rowIndex);
  }
  return ok({ message: 'Logged out' });
}

function handleForgotPassword(body) {
  const { email } = body;
  if (!email) return fail('Email required');

  const sheet = getSheet(SH.USERS);
  const { row: user, rowIndex } = findRow(sheet, 'email', email);

  if (user) {
    const resetToken = generateToken();
    const expiry = new Date(Date.now() + 60 * 60 * 1000); // 1 hour
    updateRow(sheet, rowIndex, { reset_token: resetToken, reset_expiry: expiry.toISOString() });

    const resetLink = `${BASE_URL}?reset_token=${resetToken}`;
    sendEmail(email, `Reset Your MSRC Password`,
      `<h2>Password Reset Request</h2>
       <p>Click the button below to reset your password. This link expires in 1 hour.</p>
       <p style="margin:24px 0"><a href="${resetLink}" style="background:#0f2444;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none">Reset Password →</a></p>
       <p style="color:#6b7280;font-size:13px">If you didn't request this, ignore this email.</p>
       <p style="color:#6b7280;font-size:12px">Or copy this link: ${resetLink}</p>`
    );
  }
  // Always return OK to prevent email enumeration
  return ok({ message: 'If that email is registered, a reset link has been sent.' });
}

function handleResetPassword(body) {
  const { reset_token, new_password } = body;
  if (!reset_token || !new_password) return fail('Token and new password required');
  if (new_password.length < 8) return fail('Password must be at least 8 characters');

  const sheet = getSheet(SH.USERS);
  const { row: user, rowIndex } = findRow(sheet, 'reset_token', reset_token);
  if (!user) return fail('Invalid or expired reset token');
  if (new Date(user.reset_expiry) < new Date()) return fail('Reset token has expired. Please request a new one.');

  updateRow(sheet, rowIndex, {
    password_hash: hashPassword(new_password),
    reset_token: '', reset_expiry: ''
  });
  return ok({ message: 'Password reset successfully. You can now log in.' });
}

// ════════════════════════════════════════════════════════════════════════════
// PROFILE
// ════════════════════════════════════════════════════════════════════════════

function handleGetProfile(token) {
  const user = requireAuth(token);
  return ok({ user: sanitizeUser(user) });
}

function handleUpdateProfile(token, body) {
  const user = requireAuth(token);
  const sheet = getSheet(SH.USERS);
  const { rowIndex } = findRow(sheet, 'id', user.id);
  const allowed = ['phone', 'region', 'gender'];
  const updates = {};
  allowed.forEach(k => { if (body[k] !== undefined) updates[k] = body[k]; });
  updateRow(sheet, rowIndex, updates);
  return ok({ message: 'Profile updated' });
}

function sanitizeUser(user) {
  const safe = { ...user };
  delete safe.password_hash;
  delete safe.reset_token;
  delete safe.reset_expiry;
  return safe;
}

// ════════════════════════════════════════════════════════════════════════════
// ABSTRACTS
// ════════════════════════════════════════════════════════════════════════════

function handleSubmitAbstract(token, body) {
  const user = requireAuth(token);
  const { title, type, specialty, background, methods, results,
          conclusion, iban, coi, coi_text, authors,
          file_base64, file_name, photo_base64, photo_name } = body;

  if (!title || !type || !specialty || !background || !methods || !results || !conclusion)
    return fail('Please fill all required abstract fields');

  const wordCount = [background, methods, results, conclusion]
    .join(' ').trim().split(/\s+/).filter(Boolean).length;
  if (wordCount > 350) return fail(`Abstract exceeds 350 word limit (${wordCount} words)`);

  let fileUrl = '', photoUrl = '';
  if (file_base64 && DRIVE_FOLDER_ID) {
    fileUrl = saveFileToDrive(file_base64, file_name || 'abstract.pdf', 'application/pdf');
  }
  if (photo_base64 && DRIVE_FOLDER_ID) {
    photoUrl = saveFileToDrive(photo_base64, photo_name || 'photo.jpg', 'image/jpeg');
  }

  const absId = uuid();
  appendRow(getSheet(SH.ABSTRACTS), {
    id: absId, user_id: user.id, title, type, specialty,
    background, methods, results, conclusion,
    iban: iban || '', coi: coi ? 'yes' : 'no', coi_text: coi_text || '',
    status: 'submitted', track: '',
    file_url: fileUrl, photo_url: photoUrl, avg_score: '',
    submitted_at: new Date().toISOString()
  });

  // Save co-authors
  if (authors && Array.isArray(authors)) {
    const coSheet = getSheet(SH.CO_AUTHORS);
    authors.forEach(a => {
      if (a.name && a.name.trim()) {
        appendRow(coSheet, { id: uuid(), abstract_id: absId, name: a.name, email: a.email || '', university: a.university || '' });
      }
    });
  }

  sendEmail(user.email, `Abstract Submission Confirmed – ${CONF_NAME}`,
    `<h2>Abstract Received!</h2>
     <p>Dear ${user.full_name}, your abstract has been successfully submitted.</p>
     <table style="background:#f8f9fc;border-radius:8px;padding:16px;width:100%;margin:16px 0">
       <tr><td style="padding:4px 0;color:#6b7280">Title</td><td style="font-weight:600">${title}</td></tr>
       <tr><td style="padding:4px 0;color:#6b7280">Specialty</td><td>${specialty}</td></tr>
       <tr><td style="padding:4px 0;color:#6b7280">Abstract ID</td><td style="font-family:monospace">${absId}</td></tr>
     </table>
     <p>You will be notified by email when a decision has been made. You can also track your status in the portal.</p>`
  );
  sendEmail(ADMIN_EMAIL, `New Abstract – ${title}`,
    `<p><b>${user.full_name}</b> (${user.email}) submitted: <b>${title}</b><br>Specialty: ${specialty} | Type: ${type}</p>`
  );

  return ok({ message: 'Abstract submitted successfully', id: absId });
}

function handleGetMyAbstracts(token) {
  const user = requireAuth(token);
  const allAbs = sheetToObjects(getSheet(SH.ABSTRACTS)).filter(a => a.user_id === user.id);
  const allEvals = sheetToObjects(getSheet(SH.EVALUATIONS));

  const result = allAbs.map(a => {
    const scores = allEvals.filter(e => e.abstract_id === a.id)
      .map(e => parseFloat(e.avg_score)).filter(s => !isNaN(s));
    const avg = scores.length ? (scores.reduce((x,y) => x + y, 0) / scores.length).toFixed(1) : '—';
    return { ...a, avg_score: avg };
  });

  return ok({ abstracts: result });
}

// ════════════════════════════════════════════════════════════════════════════
// EVALUATOR
// ════════════════════════════════════════════════════════════════════════════

function handleApplyEvaluator(token, body) {
  const user = requireAuth(token);
  const { academic_level, specialty, years_exp, prev_exp, prev_exp_detail } = body;
  if (!academic_level || !specialty || years_exp === undefined) return fail('Missing required fields');

  const appSheet = getSheet(SH.EVAL_APPS);
  const alreadyPending = sheetToObjects(appSheet).some(a => a.user_id === user.id && a.status === 'pending');
  if (alreadyPending) return fail('You already have a pending application');

  appendRow(appSheet, {
    id: uuid(), user_id: user.id,
    academic_level, specialty, years_exp,
    prev_exp: prev_exp || 'no',
    prev_exp_detail: prev_exp_detail || '',
    cv_url: '', status: 'pending',
    applied_at: new Date().toISOString()
  });

  sendEmail(user.email, `Evaluator Application Received – ${CONF_NAME}`,
    `<p>Dear ${user.full_name},</p>
     <p>Thank you for applying to be an evaluator for MSRC 2026. The admin team will review your application and you will be notified of the decision by email.</p>`
  );
  sendEmail(ADMIN_EMAIL, `New Evaluator Application – ${user.full_name}`,
    `<p><b>${user.full_name}</b> (${user.email}) applied as evaluator.<br>Specialty: ${specialty} | Level: ${academic_level} | Experience: ${years_exp} yrs</p>`
  );

  return ok({ message: 'Application submitted successfully' });
}

function handleGetEvalStatus(token) {
  const user = requireAuth(token);
  const apps = sheetToObjects(getSheet(SH.EVAL_APPS)).filter(a => a.user_id === user.id);
  const latest = apps.length ? apps.sort((a,b) => new Date(b.applied_at) - new Date(a.applied_at))[0] : null;
  return ok({ application: latest });
}

function handleGetAssignments(token) {
  const user = requireAuth(token);
  requireRole(user, 'evaluator', 'admin', 'super_admin');

  const assignments = sheetToObjects(getSheet(SH.ASSIGNMENTS)).filter(a => a.evaluator_id === user.id);
  const allAbs = sheetToObjects(getSheet(SH.ABSTRACTS));
  const allEvals = sheetToObjects(getSheet(SH.EVALUATIONS));

  const result = assignments.map(a => {
    const abs = allAbs.find(ab => ab.id === a.abstract_id) || null;
    const alreadyEvaluated = allEvals.some(e => e.abstract_id === a.abstract_id && e.evaluator_id === user.id);
    return { ...a, abstract: abs, already_evaluated: alreadyEvaluated };
  });

  return ok({ assignments: result });
}

function handleSubmitEvaluation(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'evaluator', 'admin', 'super_admin');

  const { assignment_id, abstract_id, originality, methodology,
          clarity, relevance, recommendation, comments, track_assigned } = body;
  if (!abstract_id || !originality || !methodology || !clarity || !relevance || !recommendation)
    return fail('All evaluation fields are required');

  const scores = [originality, methodology, clarity, relevance].map(Number);
  if (scores.some(s => isNaN(s) || s < 1 || s > 10)) return fail('Scores must be between 1 and 10');

  const avg = (scores.reduce((a,b) => a + b, 0) / 4).toFixed(2);

  appendRow(getSheet(SH.EVALUATIONS), {
    id: uuid(), abstract_id, evaluator_id: user.id,
    originality, methodology, clarity, relevance, avg_score: avg,
    recommendation, track_assigned: track_assigned || '',
    comments: comments || '',
    submitted_at: new Date().toISOString()
  });

  // Update assignment status
  if (assignment_id) {
    const aSheet = getSheet(SH.ASSIGNMENTS);
    const { rowIndex } = findRow(aSheet, 'id', assignment_id);
    if (rowIndex > 0) updateRow(aSheet, rowIndex, { eval_status: 'evaluated' });
  }

  // Recalculate abstract composite avg
  const allAbsEvals = sheetToObjects(getSheet(SH.EVALUATIONS)).filter(e => e.abstract_id === abstract_id);
  const allScores = allAbsEvals.map(e => parseFloat(e.avg_score)).filter(s => !isNaN(s));
  const compositeAvg = allScores.length ? (allScores.reduce((a,b) => a + b, 0) / allScores.length).toFixed(2) : avg;
  const absSheet = getSheet(SH.ABSTRACTS);
  const { rowIndex: absRow } = findRow(absSheet, 'id', abstract_id);
  if (absRow > 0) updateRow(absSheet, absRow, { avg_score: compositeAvg });

  return ok({ message: 'Evaluation submitted', avg_score: avg });
}

// ════════════════════════════════════════════════════════════════════════════
// CONFERENCE REGISTRATION
// ════════════════════════════════════════════════════════════════════════════

function handleVerifyDiscount(body) {
  const { code } = body;
  if (!code) return fail('Discount code required');

  const { row: disc } = findRow(getSheet(SH.DISC_CODES), 'code', code.toUpperCase());
  if (!disc) return fail('Invalid or expired discount code');

  const isActive = disc.active === true || String(disc.active).toUpperCase() === 'TRUE';
  if (!isActive) return fail('This discount code is no longer active');

  const maxUses = parseInt(disc.max_uses) || 0;
  const usedCount = parseInt(disc.used_count) || 0;
  if (maxUses > 0 && usedCount >= maxUses) return fail('Discount code has reached its usage limit');

  return ok({ discount: { code: disc.code, type: disc.type, amount: parseInt(disc.amount) } });
}

function handleRegisterConference(token, body) {
  const user = requireAuth(token);
  const { phone, national_id, workshop, discount_code } = body;
  if (!phone || !national_id) return fail('Phone and national ID are required');

  const regSheet = getSheet(SH.REGISTRATIONS);
  if (sheetToObjects(regSheet).some(r => r.user_id === user.id))
    return fail('You are already registered for this conference');

  let discAmount = 0;
  let discCode = '';

  if (discount_code) {
    const discSheet = getSheet(SH.DISC_CODES);
    const { row: disc, rowIndex: dRow } = findRow(discSheet, 'code', discount_code.toUpperCase());
    if (disc && (disc.active === true || String(disc.active).toUpperCase() === 'TRUE')) {
      discAmount = parseInt(disc.amount) || 0;
      discCode = disc.code;
      updateRow(discSheet, dRow, { used_count: (parseInt(disc.used_count) || 0) + 1 });
    }
  }

  const workshopAmt = workshop ? 50 : 0;
  const baseAmount = 50;
  const total = Math.max(0, baseAmount - discAmount + workshopAmt);

  const regId = uuid();
  appendRow(regSheet, {
    id: regId, user_id: user.id, phone, national_id,
    workshop: workshop ? 'yes' : 'no',
    discount_code: discCode, base_amount: baseAmount,
    discount_amount: discAmount, total_amount: total,
    payment_status: total === 0 ? 'free' : 'pending',
    payment_ref: '', mode: 'online',
    registered_at: new Date().toISOString()
  });

  sendEmail(user.email, `Conference Registration Confirmed – ${CONF_NAME}`,
    `<h2>You're Registered! 🎉</h2>
     <p>Dear ${user.full_name},</p>
     <p>Your conference registration is confirmed.</p>
     <table style="background:#f8f9fc;border-radius:8px;padding:16px;width:100%;margin:16px 0">
       <tr><td style="padding:4px 0;color:#6b7280">Conference</td><td><b>${CONF_NAME}</b></td></tr>
       <tr><td style="padding:4px 0;color:#6b7280">Workshop</td><td>${workshop ? '✓ Included' : 'Not included'}</td></tr>
       <tr><td style="padding:4px 0;color:#6b7280">Amount</td><td><b>${total === 0 ? 'FREE' : total + ' SAR'}</b></td></tr>
     </table>
     <p>Please keep this confirmation for your records. See you at the conference!</p>`
  );

  return ok({ message: 'Registration successful', total });
}

// ════════════════════════════════════════════════════════════════════════════
// CERTIFICATES
// ════════════════════════════════════════════════════════════════════════════

function handleGetMyCerts(token) {
  const user = requireAuth(token);
  const certs = sheetToObjects(getSheet(SH.CERTIFICATES)).filter(c => c.user_id === user.id);
  return ok({ certificates: certs });
}

function handleVerifyCert(cert_id) {
  if (!cert_id) return fail('Certificate ID required');
  const { row: cert } = findRow(getSheet(SH.CERTIFICATES), 'cert_id', cert_id.toUpperCase());
  if (!cert) return fail('Certificate not found. Please check the ID and try again.');
  const { row: user } = findRow(getSheet(SH.USERS), 'id', cert.user_id);
  return ok({
    verified: true,
    cert: {
      cert_id: cert.cert_id,
      cert_type: cert.cert_type,
      issued_at: cert.issued_at,
      recipient: user ? user.full_name : 'Unknown',
      conference: CONF_NAME
    }
  });
}

function issueCertificate(userId, certType) {
  const certId = 'MSRC-2026-' + Utilities.getUuid().replace(/-/g,'').substr(0, 6).toUpperCase();
  const cert = { id: uuid(), user_id: userId, cert_id: certId, cert_type: certType || 'attendance', issued_at: new Date().toISOString() };
  appendRow(getSheet(SH.CERTIFICATES), cert);
  return cert;
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – STATS
// ════════════════════════════════════════════════════════════════════════════

function handleAdminStats(token) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const abstracts     = sheetToObjects(getSheet(SH.ABSTRACTS));
  const registrations = sheetToObjects(getSheet(SH.REGISTRATIONS));
  const evalApps      = sheetToObjects(getSheet(SH.EVAL_APPS));
  const users         = sheetToObjects(getSheet(SH.USERS));
  const certs         = sheetToObjects(getSheet(SH.CERTIFICATES));

  const revenue = registrations.reduce((s, r) => s + (parseFloat(r.total_amount) || 0), 0);
  const breakdown = {};
  abstracts.forEach(a => { breakdown[a.status] = (breakdown[a.status] || 0) + 1; });

  return ok({
    stats: {
      users:         users.length,
      abstracts:     abstracts.length,
      registrations: registrations.length,
      revenue:       revenue.toLocaleString('en-SA'),
      evaluators:    evalApps.filter(a => a.status === 'approved').length,
      accepted:      abstracts.filter(a => String(a.status).startsWith('accepted')).length,
      certificates:  certs.length,
      with_workshop: registrations.filter(r => r.workshop === 'yes').length,
      status_breakdown: breakdown
    }
  });
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – ABSTRACTS
// ════════════════════════════════════════════════════════════════════════════

function handleAdminAbstracts(token, params) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const users = sheetToObjects(getSheet(SH.USERS));
  let abstracts = sheetToObjects(getSheet(SH.ABSTRACTS)).map(a => {
    const u = users.find(u => u.id === a.user_id);
    return { ...a, presenter_name: u ? u.full_name : '—', presenter_email: u ? u.email : '' };
  });

  if (params && params.status && params.status !== 'all')
    abstracts = abstracts.filter(a => a.status === params.status);
  if (params && params.search) {
    const q = params.search.toLowerCase();
    abstracts = abstracts.filter(a =>
      a.title.toLowerCase().includes(q) ||
      (a.presenter_name || '').toLowerCase().includes(q));
  }

  return ok({ abstracts });
}

function handleAdminUpdateAbstract(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const { abstract_id, status, track } = body;
  if (!abstract_id || !status) return fail('abstract_id and status required');

  const absSheet = getSheet(SH.ABSTRACTS);
  const { row: abs, rowIndex } = findRow(absSheet, 'id', abstract_id);
  if (!abs) return fail('Abstract not found');
  updateRow(absSheet, rowIndex, { status, track: track || abs.track || '' });

  const statusMessages = {
    accepted_oral:   `🎉 Congratulations! Your abstract has been <b>accepted for Oral Presentation</b>.`,
    accepted_poster: `🎉 Congratulations! Your abstract has been <b>accepted for Poster Presentation</b>.`,
    rejected:        `We regret to inform you that your abstract was not selected for presentation.`,
    needs_revision:  `Your abstract has been reviewed and requires <b>revision</b>. Please check the portal for details.`,
    under_review:    `Your abstract is now <b>under review</b> by our panel of evaluators.`,
    submitted:       `Your abstract status has been updated to Submitted.`
  };
  const msg = statusMessages[status] || `Your abstract status has been updated to: ${status}`;

  const { row: presenter } = findRow(getSheet(SH.USERS), 'id', abs.user_id);
  if (presenter) {
    sendEmail(presenter.email, `Abstract Decision – ${CONF_NAME}`,
      `<p>Dear ${presenter.full_name},</p>
       <p>${msg}</p>
       <p><b>Abstract:</b> ${abs.title}</p>
       <p>Log in to the portal for more details.</p>`
    );
  }

  logAdminAction(user.id, 'update_abstract_status', abstract_id, `Status: ${status}`);
  return ok({ message: 'Abstract status updated' });
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – EVALUATORS
// ════════════════════════════════════════════════════════════════════════════

function handleAdminEvalApps(token) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const users = sheetToObjects(getSheet(SH.USERS));
  const apps  = sheetToObjects(getSheet(SH.EVAL_APPS)).map(a => {
    const u = users.find(u => u.id === a.user_id);
    return { ...a, applicant_name: u ? u.full_name : '—', applicant_email: u ? u.email : '' };
  });
  return ok({ applications: apps });
}

function handleAdminProcessEvalApp(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const { app_id, action } = body;
  if (!app_id || !['approve','reject'].includes(action)) return fail('app_id and action (approve|reject) required');

  const appSheet = getSheet(SH.EVAL_APPS);
  const { row: app, rowIndex } = findRow(appSheet, 'id', app_id);
  if (!app) return fail('Application not found');

  const newStatus = action === 'approve' ? 'approved' : 'rejected';
  updateRow(appSheet, rowIndex, { status: newStatus });

  const usersSheet = getSheet(SH.USERS);
  const { row: evalUser, rowIndex: uRow } = findRow(usersSheet, 'id', app.user_id);

  if (action === 'approve' && evalUser) {
    const roles = String(evalUser.roles).split(',').map(r => r.trim()).filter(Boolean);
    if (!roles.includes('evaluator')) {
      roles.push('evaluator');
      updateRow(usersSheet, uRow, { roles: roles.join(',') });
    }
    sendEmail(evalUser.email, `Evaluator Application Approved – ${CONF_NAME}`,
      `<p>Dear ${evalUser.full_name},</p>
       <p>🎉 Congratulations! Your evaluator application has been <b>approved</b>.</p>
       <p>You now have access to the Evaluator Dashboard where you can review and score assigned abstracts. Log in to the portal to get started.</p>`
    );
  } else if (evalUser) {
    sendEmail(evalUser.email, `Evaluator Application Update – ${CONF_NAME}`,
      `<p>Dear ${evalUser.full_name},</p>
       <p>Thank you for your interest in evaluating abstracts for MSRC 2026. After careful review, we were unable to accept your application at this time.</p>
       <p>We appreciate your enthusiasm and hope to see you at the conference.</p>`
    );
  }

  logAdminAction(user.id, `${action}_evaluator_app`, app_id, `User: ${app.user_id}`);
  return ok({ message: `Application ${newStatus}` });
}

function handleAdminAssignAbstracts(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const { evaluator_id, abstract_ids } = body;
  if (!evaluator_id || !abstract_ids || !abstract_ids.length) return fail('evaluator_id and abstract_ids[] required');

  const assignSheet = getSheet(SH.ASSIGNMENTS);
  const existing = sheetToObjects(assignSheet);
  let assigned = 0;

  abstract_ids.forEach(absId => {
    if (!existing.some(a => a.evaluator_id === evaluator_id && a.abstract_id === absId)) {
      appendRow(assignSheet, {
        id: uuid(), evaluator_id, abstract_id: absId,
        assigned_at: new Date().toISOString(), eval_status: 'pending'
      });
      assigned++;
    }
  });

  const { row: evalUser } = findRow(getSheet(SH.USERS), 'id', evaluator_id);
  if (evalUser && assigned > 0) {
    sendEmail(evalUser.email, `New Abstracts Assigned – ${CONF_NAME}`,
      `<p>Dear ${evalUser.full_name},</p>
       <p><b>${assigned} new abstract(s)</b> have been assigned to you for evaluation.</p>
       <p>Log in to the portal and navigate to <b>My Assignments</b> to begin your review.</p>`
    );
  }

  logAdminAction(user.id, 'assign_abstracts', evaluator_id, `${assigned} abstracts assigned`);
  return ok({ message: `${assigned} abstract(s) assigned successfully` });
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – REGISTRATIONS
// ════════════════════════════════════════════════════════════════════════════

function handleAdminRegistrations(token) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const users = sheetToObjects(getSheet(SH.USERS));
  const regs  = sheetToObjects(getSheet(SH.REGISTRATIONS)).map(r => {
    const u = users.find(u => u.id === r.user_id);
    return { ...r,
      user_name:       u ? u.full_name : '—',
      user_email:      u ? u.email : '',
      user_university: u ? u.university : ''
    };
  });

  const revenue     = regs.reduce((s, r) => s + (parseFloat(r.total_amount) || 0), 0);
  const withWshop   = regs.filter(r => r.workshop === 'yes').length;
  const discUsed    = regs.filter(r => r.discount_code).length;

  return ok({ registrations: regs, stats: { total: regs.length, revenue, with_workshop: withWshop, discount_used: discUsed } });
}

function handleAdminAddOnsite(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const { user_email, workshop, notes } = body;
  if (!user_email) return fail('User email required');

  const { row: regUser } = findRow(getSheet(SH.USERS), 'email', user_email);
  if (!regUser) return fail('User not found. They must register an account first.');

  const regSheet = getSheet(SH.REGISTRATIONS);
  if (sheetToObjects(regSheet).some(r => r.user_id === regUser.id))
    return fail('This user is already registered');

  const workshopAmt = workshop ? 50 : 0;
  appendRow(regSheet, {
    id: uuid(), user_id: regUser.id, phone: regUser.phone || '',
    national_id: '', workshop: workshop ? 'yes' : 'no',
    discount_code: 'ONSITE', base_amount: 50, discount_amount: 50,
    total_amount: workshopAmt,
    payment_status: 'paid_onsite',
    payment_ref: notes || 'Onsite registration',
    mode: 'onsite', registered_at: new Date().toISOString()
  });

  // Auto-issue attendance certificate
  const cert = issueCertificate(regUser.id, 'attendance');
  sendEmail(regUser.email, `Conference Registration & Certificate – ${CONF_NAME}`,
    `<p>Dear ${regUser.full_name},</p>
     <p>You have been registered for the conference onsite. Your attendance certificate has been generated.</p>
     <p>Certificate ID: <code>${cert.cert_id}</code></p>`
  );

  logAdminAction(user.id, 'add_onsite_registration', regUser.id, user_email);
  return ok({ message: 'Onsite registration added and certificate generated', cert_id: cert.cert_id });
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – USERS
// ════════════════════════════════════════════════════════════════════════════

function handleAdminUsers(token) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const users = sheetToObjects(getSheet(SH.USERS)).map(sanitizeUser);
  return ok({ users });
}

function handleAdminUpdateUserRole(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'super_admin');
  const { user_id, roles } = body;
  if (!user_id || !roles) return fail('user_id and roles required');

  const sheet = getSheet(SH.USERS);
  const { row: target, rowIndex } = findRow(sheet, 'id', user_id);
  if (!target) return fail('User not found');

  updateRow(sheet, rowIndex, { roles });
  logAdminAction(user.id, 'update_user_roles', user_id, `Roles: ${roles}`);
  return ok({ message: 'User roles updated' });
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN – CERTIFICATES
// ════════════════════════════════════════════════════════════════════════════

function handleAdminIssueCert(token, body) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');
  const { user_email, cert_type } = body;
  if (!user_email) return fail('User email required');

  const { row: targetUser } = findRow(getSheet(SH.USERS), 'email', user_email);
  if (!targetUser) return fail('User not found');

  const cert = issueCertificate(targetUser.id, cert_type || 'attendance');
  sendEmail(targetUser.email, `Your Certificate – ${CONF_NAME}`,
    `<p>Dear ${targetUser.full_name},</p>
     <p>Your <b>${cert_type || 'attendance'}</b> certificate for ${CONF_NAME} has been issued.</p>
     <p>Certificate ID: <code>${cert.cert_id}</code></p>
     <p>You can verify your certificate at: <a href="${BASE_URL}">${BASE_URL}</a></p>`
  );
  logAdminAction(user.id, 'issue_certificate', targetUser.id, user_email);
  return ok({ message: 'Certificate issued successfully', cert_id: cert.cert_id });
}

function handleAdminBulkCerts(token) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const regs = sheetToObjects(getSheet(SH.REGISTRATIONS))
    .filter(r => r.payment_status === 'paid' || r.payment_status === 'free' || r.payment_status === 'paid_onsite');
  const existingCerts = sheetToObjects(getSheet(SH.CERTIFICATES));

  let count = 0;
  regs.forEach(r => {
    if (!existingCerts.some(c => c.user_id === r.user_id && c.cert_type === 'attendance')) {
      issueCertificate(r.user_id, 'attendance');
      count++;
    }
  });

  logAdminAction(user.id, 'bulk_generate_certificates', 'all', `${count} generated`);
  return ok({ message: `${count} certificate(s) generated successfully` });
}

// ════════════════════════════════════════════════════════════════════════════
// EXPORT CSV
// ════════════════════════════════════════════════════════════════════════════

function handleExportCSV(token, type) {
  const user = requireAuth(token);
  requireRole(user, 'admin', 'super_admin');

  const sheetMap = { abstracts: SH.ABSTRACTS, registrations: SH.REGISTRATIONS, users: SH.USERS };
  const sheetName = sheetMap[type];
  if (!sheetName) return fail('Unknown export type. Valid: abstracts, registrations, users');

  const data = getSheet(sheetName).getDataRange().getValues();
  // Exclude sensitive columns for users export
  let rows = data;
  if (type === 'users') {
    const headers = data[0];
    const exclude = ['password_hash','reset_token','reset_expiry'];
    const keepIdx = headers.map((h,i) => exclude.includes(h) ? -1 : i).filter(i => i >= 0);
    rows = data.map(row => keepIdx.map(i => row[i]));
  }

  const csv = rows.map(row =>
    row.map(cell => `"${String(cell).replace(/"/g,'""')}"`).join(',')
  ).join('\n');

  return ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Content-Disposition', `attachment; filename="msrc_2026_${type}.csv"`);
}

// ════════════════════════════════════════════════════════════════════════════
// DRIVE FILE UPLOAD
// ════════════════════════════════════════════════════════════════════════════

function saveFileToDrive(base64Data, filename, mimeType) {
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, filename);
    let folder;
    if (DRIVE_FOLDER_ID) {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } else {
      const ss  = SpreadsheetApp.openById(SS_ID);
      const pid = DriveApp.getFileById(SS_ID).getParents().next().getId();
      // Create MSRC uploads subfolder
      const root = DriveApp.getFolderById(pid);
      const iter = root.getFoldersByName('MSRC_Uploads');
      folder = iter.hasNext() ? iter.next() : root.createFolder('MSRC_Uploads');
    }
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    console.error('Drive upload error:', e.message);
    return '';
  }
}

// ════════════════════════════════════════════════════════════════════════════
// EMAIL
// ════════════════════════════════════════════════════════════════════════════

function sendEmail(to, subject, htmlBody) {
  try {
    if (!to || !to.includes('@')) return;
    GmailApp.sendEmail(to, `[MSRC 2026] ${subject}`, '', {
      htmlBody: `
      <!DOCTYPE html><html><body>
      <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)">
        <div style="background:#0f2444;padding:28px 36px;display:flex;align-items:center;gap:12px">
          <div>
            <div style="font-family:Georgia,serif;font-size:26px;font-weight:900;color:#c8a84b;letter-spacing:1px">MSRC</div>
            <div style="font-size:11px;color:rgba(255,255,255,.5);letter-spacing:1.5px;text-transform:uppercase;margin-top:2px">Medical Students Research Conference 2026</div>
          </div>
        </div>
        <div style="padding:36px">${htmlBody}</div>
        <div style="background:#f8f9fc;padding:20px 36px;border-top:1px solid #eef0f6;font-size:12px;color:#9ca3af">
          This email was sent by the MSRC 2026 Conference Portal. Please do not reply to this email.
        </div>
      </div></body></html>`,
      name: 'MSRC 2026 Portal'
    });
  } catch(e) {
    console.error('Email error to ' + to + ': ' + e.message);
  }
}

// ════════════════════════════════════════════════════════════════════════════
// ADMIN LOG
// ════════════════════════════════════════════════════════════════════════════

function logAdminAction(adminId, action, target, details) {
  try {
    appendRow(getSheet(SH.LOGS), {
      id: uuid(), admin_id: adminId, action, target, details,
      timestamp: new Date().toISOString()
    });
  } catch(e) {}
}

// ════════════════════════════════════════════════════════════════════════════
// SETUP – initialise all sheets + seed discount codes
// ════════════════════════════════════════════════════════════════════════════

function handleSetup() {
  Object.keys(COLUMNS).forEach(name => getSheet(name));

  // Seed discount codes if empty
  const discSheet = getSheet(SH.DISC_CODES);
  if (discSheet.getLastRow() < 2) {
    const codes = [
      { code:'KAU50',    type:'conference_free', amount:50, max_uses:2000, used_count:0, active:true },
      { code:'STAFF25',  type:'discount_fixed',  amount:25, max_uses:100,  used_count:0, active:true },
      { code:'WORKSHOP', type:'workshop_free',    amount:50, max_uses:50,   used_count:0, active:true }
    ];
    codes.forEach(c => appendRow(discSheet, c));
  }

  return ok({ message: '✅ MSRC 2026 backend setup complete. All sheets initialised.' });
}
