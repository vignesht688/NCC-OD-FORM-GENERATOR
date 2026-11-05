const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const dotenv = require('dotenv');
const rootEnvPath = path.resolve(__dirname, '..', '.env');

function refreshEnv() {
  try {
    dotenv.config({ path: rootEnvPath, override: true });
  } catch (err) {
    // safely ignore file read errors
  }
}

// In-memory reference used strictly during automated test runs
let latestTestEmail = null;

function getLatestTestEmail() {
  return latestTestEmail;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getTransporter() {
  refreshEnv();
  const host = (process.env.SMTP_HOST || '').trim();
  const user = (process.env.SMTP_USER || '').trim();
  const pass = (process.env.SMTP_PASS || '').trim();
  const port = Number(process.env.SMTP_PORT) || 587;
  const secure = process.env.SMTP_SECURE === 'true' || port === 465;

  // Strict check: host, user, and pass must all be non-empty strings
  if (!host || !user || !pass) {
    return null;
  }

  return nodemailer.createTransport({
    host,
    port,
    secure,
    auth: {
      user,
      pass
    },
    tls: {
      rejectUnauthorized: process.env.NODE_ENV === 'production'
    }
  });
}

/**
 * Checks whether SMTP environment variables are properly configured
 * @returns {{ configured: boolean, missing: string[] }}
 */
function checkSmtpConfig() {
  refreshEnv();
  const missing = [];
  if (!process.env.SMTP_HOST || !process.env.SMTP_HOST.trim()) missing.push('SMTP_HOST');
  if (!process.env.SMTP_USER || !process.env.SMTP_USER.trim()) missing.push('SMTP_USER');
  if (!process.env.SMTP_PASS || !process.env.SMTP_PASS.trim()) missing.push('SMTP_PASS');

  return {
    configured: missing.length === 0,
    missing
  };
}

/**
 * Verifies SMTP connection configuration with the mail server
 * @returns {Promise<{ ok: boolean, error?: string }>}
 */
async function verifySmtpConnection() {
  const { configured, missing } = checkSmtpConfig();
  if (!configured) {
    return { ok: false, error: `Missing in .env: ${missing.join(', ')}` };
  }
  const transporter = getTransporter();
  if (!transporter) {
    return { ok: false, error: 'Could not create Nodemailer transporter.' };
  }
  try {
    await transporter.verify();
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Sends a 6-digit OTP email for account verification
 * @param {string} userEmail - User's registered email
 * @param {string} userName - User's name
 * @param {string} otpCode - 6-digit numeric OTP
 * @returns {Promise<{ delivered: boolean, error?: string, messageId?: string }>}
 */
async function sendOtpEmail(userEmail, userName, otpCode) {
  refreshEnv();
  const safeName = escapeHtml(userName || userEmail.split('@')[0]);
  let fromAddress = (process.env.SMTP_FROM || '').trim();
  if (!fromAddress || fromAddress.includes('yourgmail@gmail.com')) {
    fromAddress = `"NCC OD Portal" <${process.env.SMTP_USER || 'noreply@srcas.ac.in'}>`;
  }
  const subject = `${otpCode} is your NCC OD Form Generator Verification Code`;

  const plainText = `Hello ${safeName},

Thank you for registering on the NCC OD Form Generator Portal (Sri Ramakrishna College of Arts & Science).

Your 6-digit One-Time Password (OTP) for account verification is:

    ${otpCode}

IMPORTANT:
- This OTP expires in 10 minutes.
- It can be used only once.
- Never share this code with anyone.

If you did not attempt to register on the NCC OD Form Generator, please disregard this email.

---
National Cadet Corps (NCC) - Sri Ramakrishna College of Arts & Science
Website: NCC OD Form Generator
`;

  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${subject}</title>
  <style>
    body { margin: 0; padding: 0; background-color: #0d1117; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #e6edf3; }
    .container { max-width: 560px; margin: 30px auto; background: #161b22; border-radius: 14px; border: 1px solid #30363d; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.6); }
    .header { background: linear-gradient(135deg, #0b1f3a 0%, #1e3a8a 100%); padding: 28px 24px; text-align: center; border-bottom: 2px solid #3b82f6; }
    .badge { display: inline-block; background: #f59e0b; color: #000; font-weight: 800; font-size: 13px; letter-spacing: 2px; padding: 4px 12px; border-radius: 4px; margin-bottom: 10px; }
    .title { color: #ffffff; font-size: 22px; font-weight: 700; margin: 0; }
    .subtitle { color: #93c5fd; font-size: 13px; margin: 6px 0 0; }
    .content { padding: 32px 28px; line-height: 1.6; }
    .greeting { font-size: 17px; font-weight: 600; color: #f0f6fc; margin-bottom: 12px; }
    .text { color: #8b949e; font-size: 14px; margin-bottom: 20px; }
    .otp-box { background: #0d1117; border: 2px dashed #3b82f6; border-radius: 12px; padding: 24px 16px; text-align: center; margin: 24px 0; }
    .otp-label { font-size: 12px; color: #8b949e; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; margin-bottom: 8px; }
    .otp-code { font-size: 40px; font-weight: 800; letter-spacing: 10px; color: #60a5fa; font-family: 'Courier New', Courier, monospace; margin: 6px 0; }
    .otp-expiry { font-size: 13px; color: #f59e0b; font-weight: 600; margin-top: 8px; }
    .notice { background: rgba(59, 130, 246, 0.1); border-left: 3px solid #3b82f6; padding: 12px 14px; font-size: 13px; color: #93c5fd; border-radius: 0 6px 6px 0; margin-bottom: 20px; }
    .footer { background: #0d1117; padding: 20px 24px; text-align: center; font-size: 12px; color: #6e7681; border-top: 1px solid #21262d; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="badge">NCC</div>
      <h1 class="title">NCC OD Form Generator</h1>
      <p class="subtitle">Sri Ramakrishna College of Arts &amp; Science</p>
    </div>
    <div class="content">
      <div class="greeting">Hello ${safeName},</div>
      <p class="text">
        Thank you for creating an account on the <strong>NCC OD Form Generator</strong>. To complete your registration and verify your email address, please enter the one-time verification code below:
      </p>

      <div class="otp-box">
        <div class="otp-label">One-Time Password (OTP)</div>
        <div class="otp-code">${otpCode}</div>
        <div class="otp-expiry">&#9201; Expires in 10 minutes</div>
      </div>

      <div class="notice">
        <strong>Security Notice:</strong> This code is strictly single-use. Do not share your OTP with anyone. If you did not request this code, you can safely ignore this email.
      </div>
    </div>
    <div class="footer">
      National Cadet Corps &bull; Sri Ramakrishna College of Arts &amp; Science<br>
      Website: NCC OD Form Generator &bull; Automated System Message
    </div>
  </div>
</body>
</html>
`;

  // Fast-path: simulated delivery in automated test mode
  if (process.env.AUTOMATED_TEST === '1') {
    latestTestEmail = {
      to: userEmail,
      name: safeName,
      otp: otpCode,
      sentAt: new Date().toISOString()
    };
    return {
      delivered: true,
      messageId: 'test-mock-msg-id'
    };
  }

  // Development OTP Mode: logs OTP directly to terminal console
  if (process.env.DEV_OTP_MODE === 'true') {
    console.log('\n============================================================');
    console.log(' 🔐 [OTP DISPATCH] NCC OD PORTAL EMAIL VERIFICATION CODE');
    console.log('============================================================');
    console.log(` Recipient:    ${userEmail}`);
    console.log(` 6-Digit OTP:  ${otpCode}`);
    console.log(` Expires In:   10 minutes`);
    console.log('============================================================\n');
    return {
      delivered: true,
      devMode: true,
      messageId: 'dev-mode-' + Date.now()
    };
  }

  const { configured, missing } = checkSmtpConfig();

  if (!configured) {
    console.warn(`[EMAIL SERVICE NOTICE] SMTP credentials not set (${missing.join(', ')}). Logging verification OTP to console/Render logs.`);
    console.log('\n============================================================');
    console.log(' 🔐 [OTP DISPATCH - SERVER LOG] NCC OD PORTAL VERIFICATION CODE');
    console.log('============================================================');
    console.log(` Recipient:    ${userEmail}`);
    console.log(` 6-Digit OTP:  ${otpCode}`);
    console.log(` Status:       Printed to server logs (Add SMTP_USER & SMTP_PASS in Render for live emails)`);
    console.log(` Expires In:   10 minutes`);
    console.log('============================================================\n');
    return {
      delivered: true,
      devMode: true,
      messageId: 'server-log-' + Date.now()
    };
  }

  try {
    const transporter = getTransporter();
    const info = await transporter.sendMail({
      from: fromAddress,
      to: userEmail,
      subject,
      text: plainText,
      html: htmlContent
    });

    console.log(`[EMAIL SERVICE] Verification OTP successfully sent to ${userEmail} (Message ID: ${info.messageId})`);
    return {
      delivered: true,
      messageId: info.messageId
    };
  } catch (err) {
    console.error(`[EMAIL SERVICE ERROR] Failed to send email to ${userEmail}: ${err.message}`);
    return {
      delivered: false,
      error: `Failed to deliver email: ${err.message}`
    };
  }
}

/**
 * Returns safe environment & SMTP diagnostic details without printing passwords
 * @returns {object}
 */
function getSafeSmtpDiagnostics() {
  refreshEnv();
  const fileExists = fs.existsSync(rootEnvPath);
  const host = (process.env.SMTP_HOST || '').trim();
  const port = process.env.SMTP_PORT || '465';
  const secure = process.env.SMTP_SECURE || 'true';
  const user = (process.env.SMTP_USER || '').trim();
  const pass = (process.env.SMTP_PASS || '').trim();
  const from = (process.env.SMTP_FROM || '').trim();
  const devMode = process.env.DEV_OTP_MODE === 'true';

  return {
    envPath: rootEnvPath,
    envExists: fileExists,
    devOtpMode: devMode,
    smtpHost: host || '[EMPTY]',
    smtpPort: port,
    smtpSecure: secure,
    smtpUser: user ? user : '[EMPTY - Missing in .env]',
    smtpPass: pass ? `[CONFIGURED (${pass.length} chars)]` : '[EMPTY - Missing in .env]',
    smtpFrom: from || '[EMPTY]',
    isConfigured: Boolean(host && user && pass),
    isReady: Boolean(host && user && pass && !devMode)
  };
}

module.exports = {
  sendOtpEmail,
  checkSmtpConfig,
  verifySmtpConnection,
  getLatestTestEmail,
  getSafeSmtpDiagnostics
};
