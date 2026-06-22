/**
 * ═══════════════════════════════════════════════════════════════
 *  REAC HUB PORTAL — Google Apps Script Backend  v5.2
 *  Deploy as Web App: Execute as "Me", Access "Anyone"
 * ═══════════════════════════════════════════════════════════════
 *
 *  SETUP INSTRUCTIONS:
 *  1. Open script.google.com → New Project → paste this code
 *  2. Confirm SHEET_ID below matches your spreadsheet ID
 *  3. Deploy → New Deployment → Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Copy the Web App URL → paste into index.html GAS_URL variable
 *
 *  v5.2 IMPROVEMENTS:
 *  - Fixed email sender to use noreply@reacrevenuesolutions.com instead of personal account
 *  - Added rate limiting for OTP requests (max 3 per hour)
 *  - Improved HTML email template with better branding
 *  - Added security notice in OTP emails
 *
 *  v5.1 FIX:
 *  - Month labels now correctly read from Row 1 of 'Scores' tab
 *  - Scores sheet layout (corrected):
 *      Row 1: month group labels (January 2026, February 2026 …)
 *      Row 2: column headers (Email, Team Lead, Employee No,
 *              Agent Name, Count, TPS, Attendance %, QA Score,
 *              Overall Performance, RANK  …repeated per month)
 *      Row 3+: data rows
 *  - Login sheet layout unchanged: col A=Email, B=Password, C=Name, D=Role
 *  - All frontend calls use GET (avoids CORS preflight with GAS)
 * ═══════════════════════════════════════════════════════════════
 */

// ── CONFIG ──────────────────────────────────────────────────────
const SHEET_ID  = '13t7zgokHi_7c7UM2bEVvPCmMs55q62GcVRB66ytd4Fw';

// Email configuration
const FROM_NAME = 'REAC Hub Portal';
const FROM_EMAIL = 'noreply@reacrevenuesolutions.com'; // v5.2: Organizational email
const SUPPORT_EMAIL = 'support@reacrevenuesolutions.com'; // v5.2: Support contact
const ORG_NAME = 'REAC Revenue Solutions';

// Tab names — must match exactly what's in the spreadsheet
const TAB_SCORES = 'Scores';   // agent performance data
const TAB_LOGIN  = 'Login';    // all user credentials & roles

// OTP expiry
const OTP_EXPIRY_MINUTES = 10;
const OTP_RATE_LIMIT_MINUTES = 60;  // v5.2: Rate limiting window
const OTP_MAX_ATTEMPTS = 3;         // v5.2: Max OTP requests per window

// ── SCRIPT-PROPERTIES HELPERS ────────────────────────────────────
function getProp(key)      { return PropertiesService.getScriptProperties().getProperty(key); }
function setProp(key, val) { PropertiesService.getScriptProperties().setProperty(key, val); }
function delProp(key)      { PropertiesService.getScriptProperties().deleteProperty(key); }

// ── RESPONSE HELPER ──────────────────────────────────────────────
function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ROUTE: GET (all frontend calls go here) ──────────────────────
function doGet(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'login':                   return jsonResp(actionLogin(e.parameter));
      case 'checkEmail':              return jsonResp(actionCheckEmail(e.parameter));
      case 'sendOtp':                 return jsonResp(actionSendOtp(e.parameter));
      case 'verifyOtp':               return jsonResp(actionVerifyOtp(e.parameter));
      case 'verifyOtpAndSetPassword': return jsonResp(actionVerifyOtpAndSetPassword(e.parameter));
      case 'getData':                 return jsonResp(actionGetData());
      default:                        return jsonResp({ success: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResp({ success: false, message: 'Server error: ' + err.message });
  }
}

// ── ROUTE: POST (kept for completeness; frontend uses GET) ────────
function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse(e.postData.contents); } catch(ex) {
      e.postData.contents.split('&').forEach(p => {
        const parts = p.split('=');
        if (parts.length === 2)
          body[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1].replace(/\+/g,' '));
      });
    }
    switch (body.action) {
      case 'verifyOtpAndSetPassword': return jsonResp(actionVerifyOtpAndSetPassword(body));
      default: return jsonResp({ success: false, error: 'Unknown action: ' + body.action });
    }
  } catch(err) {
    return jsonResp({ success: false, message: 'Server error: ' + err.message });
  }
}

// ── SHEET HELPERS ────────────────────────────────────────────────
function getSheet(tabName) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Sheet tab not found: "' + tabName + '". Check tab name spelling.');
  return sheet;
}

/**
 * Login sheet structure (TAB_LOGIN = 'Login')
 *   Row 1 = header:  Email | Password | Name | Role
 *   Row 2+= data
 */
function getAllLoginRows() {
  const sheet = getSheet(TAB_LOGIN);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues();
}

/**
 * Scores sheet structure (TAB_SCORES = 'Scores')  ← v5.1 CORRECTED
 *
 *   Row 1 = month group labels   (col E onwards, repeating every 6 cols)
 *             e.g. "January 2026", "February 2026", …
 *   Row 2 = column headers
 *             Email | Team Lead | Employee No | Agent Name |
 *             Count | TPS | Attendance % | QA Score | Overall Performance | RANK
 *             (the last 6 columns repeat for every month)
 *   Row 3+= agent data rows
 *
 *  Fixed columns (0-indexed):
 *    0 = Email
 *    1 = Team Lead email
 *    2 = Employee Number
 *    3 = Agent Name
 *  Repeating groups of 6 per month (starting at col index 4):
 *    +0 Count | +1 TPS | +2 Attendance % | +3 QA Score | +4 Overall Performance | +5 RANK
 */
function getAllScoreRows() {
  const sheet = getSheet(TAB_SCORES);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  // Data starts at row 3 (rows 1 & 2 are headers)
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
}

// ── ACTION: LOGIN ────────────────────────────────────────────────
function actionLogin(params) {
  const email    = (params.email    || '').toLowerCase().trim();
  const password = (params.password || '').trim();

  if (!email || !password)
    return { success: false, message: 'Email and password are required.' };

  const rows = getAllLoginRows();
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = String(rows[i][0] || '').toLowerCase().trim();
    if (rowEmail !== email) continue;

    const rowPass = String(rows[i][1] || '').trim();
    const rowName = String(rows[i][2] || '').trim();
    const rowRole = String(rows[i][3] || '').toLowerCase().trim();

    if (!rowPass)
      return { success: false, message: 'No password set. Please use Sign Up or Forgot Password.' };
    if (rowPass !== password)
      return { success: false, message: 'Incorrect password.' };

    // Determine portal role
    let role = 'agent';
    if (rowRole === 'master admin' || rowRole === 'admin') role = 'admin';
    else if (rowRole === 'team lead' || rowRole === 'teamlead') role = 'teamlead';

    return {
      success: true,
      user: { email, name: rowName, role, rawRole: rowRole }
    };
  }

  return { success: false, message: 'Email not found in the system.' };
}

// ── ACTION: CHECK EMAIL ──────────────────────────────────────────
function actionCheckEmail(params) {
  const email = (params.email || '').toLowerCase().trim();
  if (!email) return { exists: false };

  const rows = getAllLoginRows();
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = String(rows[i][0] || '').toLowerCase().trim();
    if (rowEmail !== email) continue;
    const hasPass = String(rows[i][1] || '').trim() !== '';
    return { exists: true, hasPassword: hasPass };
  }
  return { exists: false };
}

// ── v5.2: RATE LIMITING HELPER ──────────────────────────────────
function checkOtpRateLimit(email) {
  const key = 'otp_attempts_' + email;
  const stored = getProp(key);
  let attempts = [];
  
  if (stored) {
    try {
      attempts = JSON.parse(stored);
      // Filter out attempts older than the rate limit window
      const now = new Date().getTime();
      attempts = attempts.filter(t => now - t < OTP_RATE_LIMIT_MINUTES * 60 * 1000);
    } catch(e) {
      attempts = [];
    }
  }

  if (attempts.length >= OTP_MAX_ATTEMPTS) {
    return {
      allowed: false,
      message: `Too many OTP requests. Please try again in ${OTP_RATE_LIMIT_MINUTES} minutes.`
    };
  }

  // Add current attempt
  attempts.push(new Date().getTime());
  setProp(key, JSON.stringify(attempts));
  
  return { allowed: true };
}

// ── ACTION: SEND OTP ─────────────────────────────────────────────
function actionSendOtp(params) {
  const email = (params.email || '').toLowerCase().trim();
  if (!email) return { success: false, message: 'Email is required.' };

  const check = actionCheckEmail({ email });
  if (!check.exists) return { success: false, message: 'Email not found in the system.' };

  // v5.2: Check rate limit
  const rateLimitCheck = checkOtpRateLimit(email);
  if (!rateLimitCheck.allowed) {
    return { success: false, message: rateLimitCheck.message };
  }

  const otp    = String(Math.floor(100000 + Math.random() * 900000));
  const expiry = new Date().getTime() + OTP_EXPIRY_MINUTES * 60 * 1000;
  setProp('otp_' + email, JSON.stringify({ otp, expiry }));

  // v5.2: Enhanced HTML email template with organizational branding
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; margin: 0; padding: 0; background: #f5f5f5; }
    .container { max-width: 520px; margin: 0 auto; background: #ffffff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: linear-gradient(135deg, #00c8ff 0%, #378add 100%); padding: 32px 24px; text-align: center; color: white; }
    .header-title { font-size: 24px; font-weight: 900; letter-spacing: 2px; margin: 0; }
    .header-subtitle { font-size: 12px; letter-spacing: 1px; margin: 8px 0 0 0; opacity: 0.9; }
    .content { padding: 32px 24px; }
    .greeting { font-size: 14px; color: #333333; margin: 0 0 20px 0; }
    .otp-section { background: #f0f8ff; border: 2px solid #00c8ff; border-radius: 8px; padding: 24px; text-align: center; margin: 24px 0; }
    .otp-label { font-size: 12px; letter-spacing: 1px; color: #666666; text-transform: uppercase; margin-bottom: 12px; }
    .otp-code { font-size: 48px; font-weight: 900; letter-spacing: 12px; color: #00c8ff; font-family: 'Courier New', monospace; margin: 0; font-variant: tabular-nums; }
    .otp-expiry { font-size: 12px; color: #666666; margin-top: 12px; }
    .security-notice { background: #fff3cd; border-left: 4px solid #ffc107; padding: 12px 16px; margin: 20px 0; border-radius: 4px; font-size: 13px; color: #856404; }
    .security-notice strong { display: block; margin-bottom: 4px; }
    .footer { background: #f5f5f5; padding: 24px; text-align: center; border-top: 1px solid #eeeeee; font-size: 12px; color: #666666; }
    .footer-org { font-weight: 600; color: #333333; margin-bottom: 8px; }
    .footer-text { margin: 4px 0; }
    .divider { height: 1px; background: #eeeeee; margin: 20px 0; }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header">
      <div class="header-title">REAC HUB</div>
      <div class="header-subtitle">PERFORMANCE PORTAL</div>
    </div>

    <!-- Content -->
    <div class="content">
      <p class="greeting">Hello,</p>
      
      <p style="font-size: 14px; color: #555555; line-height: 1.6; margin: 0 0 24px 0;">
        You requested a verification code to access your <strong>REAC Hub Portal</strong> account. 
        Use the code below to complete your request:
      </p>

      <!-- OTP Code -->
      <div class="otp-section">
        <div class="otp-label">Your One-Time Password</div>
        <div class="otp-code">${otp}</div>
        <div class="otp-expiry">⏱ Expires in ${OTP_EXPIRY_MINUTES} minutes</div>
      </div>

      <!-- Security Notice -->
      <div class="security-notice">
        <strong>⚠️ Security Notice:</strong>
        Never share this code with anyone. ${ORG_NAME} staff will never ask for your OTP code. If you didn't request this, please contact <a href="mailto:${SUPPORT_EMAIL}" style="color: #856404; text-decoration: none;">${SUPPORT_EMAIL}</a> immediately.
      </div>

      <p style="font-size: 13px; color: #666666; line-height: 1.6; margin: 24px 0 0 0;">
        Questions? Contact our support team at <a href="mailto:${SUPPORT_EMAIL}" style="color: #00c8ff; text-decoration: none;">${SUPPORT_EMAIL}</a>
      </p>
    </div>

    <!-- Footer -->
    <div class="footer">
      <div class="footer-org">${ORG_NAME}</div>
      <div class="footer-text">Automated notification — please do not reply to this email</div>
      <div class="footer-text" style="margin-top: 12px; font-size: 11px; color: #999999;">
        This is a secure message sent from REAC Hub Portal
      </div>
    </div>
  </div>
</body>
</html>
  `;

  const plainBody = `
REAC HUB — PERFORMANCE PORTAL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Your One-Time Password (OTP):

  ${otp}

⏱ Expires in ${OTP_EXPIRY_MINUTES} minutes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SECURITY NOTICE:
Never share this code with anyone. ${ORG_NAME} staff will never ask for your OTP code.

Questions? Contact: ${SUPPORT_EMAIL}

${ORG_NAME}
Automated notification — please do not reply to this email
  `;

  try {
    // v5.2: Use Gmail alias (organizational email) with proper sender configuration
    // Note: This requires the organizational email to be configured as a send-as alias
    // If the send-as alias doesn't exist, GAS will fall back to the script owner's email
    GmailApp.sendEmail(
      email,
      'REAC Hub Portal — Your Verification Code',
      plainBody,
      {
        htmlBody: htmlBody,
        from: FROM_EMAIL,        // Organizational email
        replyTo: SUPPORT_EMAIL,  // Support team email for replies
        name: FROM_NAME          // Display name
      }
    );
    return { success: true };
  } catch(e) {
    // Fallback if organizational email is not configured
    try {
      MailApp.sendEmail(
        email,
        'REAC Hub Portal — Your Verification Code',
        plainBody,
        {
          htmlBody: htmlBody,
          replyTo: SUPPORT_EMAIL,
          name: FROM_NAME
        }
      );
      return { success: true };
    } catch(fallbackErr) {
      return { success: false, message: 'Failed to send email: ' + fallbackErr.message };
    }
  }
}

// ── ACTION: VERIFY OTP (check only — does NOT consume the OTP) ───
function actionVerifyOtp(params) {
  const email = (params.email || '').toLowerCase().trim();
  const otp   = String(params.otp || '').trim();

  if (!email || !otp) return { success: false, message: 'Email and OTP are required.' };

  const stored = getProp('otp_' + email);
  if (!stored) return { success: false, message: 'No OTP found. Please request a new one.' };

  let parsed;
  try { parsed = JSON.parse(stored); } catch(e) {
    delProp('otp_' + email);
    return { success: false, message: 'OTP data corrupted. Please request a new one.' };
  }

  if (new Date().getTime() > parsed.expiry) {
    delProp('otp_' + email);
    return { success: false, message: 'OTP has expired. Please request a new one.' };
  }

  if (otp !== parsed.otp)
    return { success: false, message: 'Incorrect OTP code. Please check and try again.' };

  // ✓ Valid — intentionally do NOT consume here; consumed in verifyOtpAndSetPassword
  return { success: true };
}

// ── ACTION: VERIFY OTP & SET PASSWORD ────────────────────────────
function actionVerifyOtpAndSetPassword(params) {
  const email    = (params.email    || '').toLowerCase().trim();
  const otp      = String(params.otp || '').trim();
  const password = (params.password || '').trim();

  if (!email)    return { success: false, message: 'Email is required.' };
  if (!otp)      return { success: false, message: 'OTP is required.' };
  if (!password) return { success: false, message: 'Password is required.' };

  // Validate OTP
  const stored = getProp('otp_' + email);
  if (!stored) return { success: false, message: 'Session expired. Please start the process again.' };

  let parsed;
  try { parsed = JSON.parse(stored); } catch(e) {
    delProp('otp_' + email);
    return { success: false, message: 'OTP data corrupted. Please start over.' };
  }

  if (new Date().getTime() > parsed.expiry) {
    delProp('otp_' + email);
    return { success: false, message: 'OTP has expired. Please start over.' };
  }

  if (otp !== parsed.otp)
    return { success: false, message: 'OTP mismatch. Please start over.' };

  // Consume OTP
  delProp('otp_' + email);

  // Find user in Login sheet and update password (column B = index 2 in 1-based)
  try {
    const loginSheet = getSheet(TAB_LOGIN);
    const rows       = getAllLoginRows();
    for (let i = 0; i < rows.length; i++) {
      const rowEmail = String(rows[i][0] || '').toLowerCase().trim();
      if (rowEmail !== email) continue;
      // Row in sheet = i + 2  (1 header row + 1-based index)
      loginSheet.getRange(i + 2, 2).setValue(password);
      SpreadsheetApp.flush();
      return { success: true };
    }
  } catch(e) {
    return { success: false, message: 'Error writing to Login sheet: ' + e.message };
  }

  return { success: false, message: 'Email not found while saving password. Contact your administrator.' };
}

// ── ACTION: GET DATA ─────────────────────────────────────────────
function actionGetData() {
  const agents     = buildAgentsData();
  const management = buildManagementData();
  return { agents, management };
}

// ── BUILD: MANAGEMENT DATA ───────────────────────────────────────
// Returns all non-agent users (admins & team leads) from Login sheet
function buildManagementData() {
  const rows = getAllLoginRows();
  return rows
    .map(row => ({
      email: String(row[0] || '').toLowerCase().trim(),
      name:  String(row[2] || '').trim(),
      role:  String(row[3] || '').toLowerCase().trim()
    }))
    .filter(u => u.email && u.role !== 'agent');
}

// ── BUILD: AGENTS DATA ───────────────────────────────────────────
/**
 * Scores sheet column layout (0-indexed):
 *   0  = Email
 *   1  = Team Lead email
 *   2  = Employee Number
 *   3  = Agent Name
 *   4  = [Month 1] Count
 *   5  = [Month 1] TPS
 *   6  = [Month 1] Attendance %
 *   7  = [Month 1] QA Score
 *   8  = [Month 1] Overall Performance
 *   9  = [Month 1] RANK
 *   10 = [Month 2] Count  … and so on every 6 columns
 *
 *  v5.1: Month labels are read from ROW 1 (not row 2).
 *        Data rows start at ROW 3 (not row 4).
 */
function buildAgentsData() {
  const sheet   = getSheet(TAB_SCORES);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Need at least row 1 (month labels) + row 2 (headers) + 1 data row
  if (lastRow < 3) return [];

  // ── v5.1 FIX: Read month group labels from ROW 1 ──────────────
  const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Collect month labels from row 1 starting at column index 4 (0-based)
  // Each month label spans 6 columns: Count, TPS, Attendance, QA, Overall, Rank
  const months = [];
  const DATA_START_COL = 4; // 0-based index where month data begins
  const COLS_PER_MONTH = 6;

  for (let c = DATA_START_COL; c < lastCol; c += COLS_PER_MONTH) {
    // The month label cell — scan a few cells in the group to find a non-empty label
    let label = '';
    for (let look = 0; look < COLS_PER_MONTH; look++) {
      const v = String(row1[c + look] || '').trim();
      // Accept any non-empty string that isn't purely numeric (avoids picking up numbers)
      if (v && !v.match(/^\d+$/) && v.length > 3) { label = v; break; }
    }
    if (!label) continue; // skip groups with no label (trailing empty columns)
    months.push({
      label,
      cols: {
        count:      c,
        tps:        c + 1,
        attendance: c + 2,
        qa:         c + 3,
        overall:    c + 4,
        rank:       c + 5
      }
    });
  }

  // ── v5.1 FIX: Data rows start at ROW 3 ────────────────────────
  const allRows = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const agents  = [];

  allRows.forEach(row => {
    const name = String(row[3] || '').trim();
    if (!name || name.toLowerCase() === 'total') return;

    const email   = String(row[0] || '').toLowerCase().trim();
    const tlEmail = String(row[1] || '').toLowerCase().trim();
    const empNo   = String(row[2] || '').trim();

    const monthData = months.map(m => {
      const get = key => {
        const idx = m.cols[key];
        if (idx === undefined || idx >= row.length) return null;
        const val = row[idx];
        if (val === null || val === undefined || val === '') return null;
        if (String(val).trim() === 'Attrite') return 'Attrite';
        return val;
      };
      const overall = get('overall');
      return {
        label:      m.label,
        count:      get('count'),
        tps:        formatPct(get('tps')),
        attendance: formatPct(get('attendance')),
        qa:         formatPct(get('qa')),
        overall:    overall === 'Attrite' ? 'Attrite' : formatPct(overall),
        rank:       get('rank')
      };
    });

    agents.push({ email, name, empNo, teamLeadEmail: tlEmail, months: monthData });
  });

  return agents;
}

// ── HELPER: FORMAT PERCENTAGE ────────────────────────────────────
function formatPct(val) {
  if (val === null || val === undefined || val === '') return null;
  if (String(val) === 'Attrite') return 'Attrite';
  if (typeof val === 'number') return (val * 100).toFixed(4) + '%';
  const s = String(val).trim();
  if (s.endsWith('%')) return s;
  const n = parseFloat(s);
  if (!isNaN(n)) return (n < 1 ? n * 100 : n).toFixed(4) + '%';
  return null;
}
