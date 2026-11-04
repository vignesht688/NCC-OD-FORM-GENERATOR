// 1. Load Root .env via dotenv first before other modules
const path = require('path');
const rootEnvPath = path.resolve(__dirname, '..', '.env');
require('dotenv').config({ path: rootEnvPath, override: true });

// 2. Import dependencies after environment variables are loaded
const express = require('express');
const fs = require('fs');
const crypto = require('crypto');
const { sendOtpEmail, checkSmtpConfig, verifySmtpConnection, getLatestTestEmail, getSafeSmtpDiagnostics } = require('./mailer');
const { exec } = require('child_process');
const docxLib = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, VerticalAlign, BorderStyle, XmlComponent } = docxLib;

// ============================================================
// HELPER: Build a proper vMerge element
// For OpenXML: RESTART = <w:vMerge w:val="restart"/>
//              CONTINUE = <w:vMerge/>  (NO val attribute!)
// ============================================================
class VMergeContinue extends XmlComponent {
    constructor() {
        super('w:vMerge'); // empty element, no val = CONTINUE
    }
}

// ============================================================
// SHARED: Build OD form table rows
// ============================================================
function buildODTableRows(cadets, particularsVal, odReasonVal) {
    const rows = [];
    // Header row
    rows.push(new TableRow({
        children: [
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'S.NO', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'REGISTER NO', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NAME', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DEPARTMENT', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'SHIFT', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'PARTICULARS', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ON DUTY', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
        ]
    }));

    const N = cadets.length;
    let P = 1;
    if (N > 27) {
        P = Math.ceil((N - 27) / 32) + 1;
    }
    const midPage = Math.ceil(P / 2);
    const startRow = midPage === 1 ? 0 : 27 + (midPage - 2) * 32;

    // Data rows
    cadets.forEach((cadet, index) => {
        let isPartRestart = false;
        let partVal = "";
        let isOdRestart = false;
        let odVal = "";

        if (startRow > 0) {
            if (index === 0) {
                isPartRestart = true;
                partVal = "";
                isOdRestart = true;
                odVal = "";
            } else if (index === startRow) {
                isPartRestart = true;
                partVal = particularsVal;
                isOdRestart = true;
                odVal = odReasonVal;
            }
        } else {
            if (index === 0) {
                isPartRestart = true;
                partVal = particularsVal;
                isOdRestart = true;
                odVal = odReasonVal;
            }
        }

        const partMerge = startRow > 0
            ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
            : (index === 0 ? 'restart' : 'continue');

        const odMerge = startRow > 0
            ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
            : (index === 0 ? 'restart' : 'continue');

        const partCell = new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isPartRestart ? partVal : "", font: 'Cambria', size: 24 })] })],
            verticalAlign: VerticalAlign.CENTER,
            verticalMerge: partMerge
        });

        const odCell = new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isOdRestart ? odVal : "", font: 'Cambria', size: 24 })] })],
            verticalAlign: VerticalAlign.CENTER,
            verticalMerge: odMerge
        });

        // Fix continue cells: remove wrong val attribute, keep element empty
        if (partMerge === 'continue') {
            const tcPr = partCell.root.find(el => el && el.rootKey === 'w:tcPr');
            if (tcPr) {
                const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                if (vMergeEl) vMergeEl.root = [];
            }
        }
        if (odMerge === 'continue') {
            const tcPr = odCell.root.find(el => el && el.rootKey === 'w:tcPr');
            if (tcPr) {
                const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                if (vMergeEl) vMergeEl.root = [];
            }
        }

        rows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || '', font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: ' ' + cadet.name, font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: ' ' + (cadet.department || ''), font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || '', font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                partCell,
                odCell,
            ]
        }));
    });
    return rows;
}

// Helper to sort cadets by year (III first, then II, then I) and name
function getYearWeight(department) {
    const dept = (department || '').trim().toUpperCase();
    if (dept.startsWith('III')) return 3;
    if (dept.startsWith('II')) return 2;
    if (dept.startsWith('I')) return 1;
    return 0;
}

function sortCadetsForOD(a, b) {
    // 1. Sort by Year (III first, then II, then I)
    const weightA = getYearWeight(a.department);
    const weightB = getYearWeight(b.department);
    if (weightB !== weightA) {
        return weightB - weightA;
    }

    // 2. Sort by Name (alphabetically)
    const nameA = (a.name || '').trim().toUpperCase();
    const nameB = (b.name || '').trim().toUpperCase();
    if (nameA !== nameB) {
        return nameA.localeCompare(nameB);
    }

    // 3. Fallback to Register Number
    const regA = (a.register_no || '').trim();
    const regB = (b.register_no || '').trim();
    return regA.localeCompare(regB);
}

// ============================================================
// SHARED: Build complete OD Document
// ============================================================
async function buildODDocument(cadets, particularsVal, odReasonVal, docConfig) {
    if (Array.isArray(cadets)) {
        cadets.sort(sortCadetsForOD);
    }
    const college = (docConfig && docConfig.college) || 'SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)';
    const unit = (docConfig && docConfig.unit) || 'NATIONAL CADET CORPS';
    const formType = (docConfig && docConfig.formType) || 'OD FORM';
    const sigLeft = (docConfig && docConfig.sigLeft) || 'NCC Officer';
    const sigRight = (docConfig && docConfig.sigRight) || 'Principal';

    const tableRows = buildODTableRows(cadets, particularsVal, odReasonVal);

    const doc = new Document({
        sections: [{
            properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
            children: [
                new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: college, font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: unit, font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: formType, font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                new Paragraph({ spacing: { before: 2160 }, children: [new TextRun('')] }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0 },
                        bottom: { style: BorderStyle.NONE, size: 0 },
                        left: { style: BorderStyle.NONE, size: 0 },
                        right: { style: BorderStyle.NONE, size: 0 },
                        insideHorizontal: { style: BorderStyle.NONE, size: 0 },
                        insideVertical: { style: BorderStyle.NONE, size: 0 },
                    },
                    rows: [new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: sigLeft, font: 'Bookman Old Style', size: 24 })] })] }),
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sigRight, font: 'Bookman Old Style', size: 24 })] })] }),
                        ]
                    })]
                })
            ]
        }]
    });
    return await Packer.toBuffer(doc);
}

// Root .env loaded via dotenv at top of file

const app = express();
const PORT = process.env.PORT || 3000;
const DATABASE_DIR = path.join(__dirname, '..', 'database');
const FRONTEND_DIR = path.join(__dirname, '..', 'frontend');
const CADETS_FILE = path.join(DATABASE_DIR, 'cadets.json');
const ATTENDANCE_FILE = path.join(DATABASE_DIR, 'attendance.json');
const DELETED_CADETS_FILE = path.join(DATABASE_DIR, 'deleted_cadets.json');

// Credential Data Directory: stores email and password records
const CREDENTIAL_DIR = path.join(__dirname, '..', 'credential data');
if (!fs.existsSync(CREDENTIAL_DIR)) {
    fs.mkdirSync(CREDENTIAL_DIR, { recursive: true });
}
const CREDENTIALS_FILE = path.join(CREDENTIAL_DIR, 'credentials.json');
const CREDENTIALS_SUMMARY_FILE = path.join(CREDENTIAL_DIR, 'credentials_summary.txt');
const USERS_FILE = CREDENTIALS_FILE;
const AUTHORIZED_EMAILS_FILE = path.join(CREDENTIAL_DIR, 'authorized_emails.json');

// Helper to safely read JSON from disk
function readJsonSafely(filePath, defaultValue = []) {
    try {
        if (fs.existsSync(filePath)) {
            const data = fs.readFileSync(filePath, 'utf8');
            return JSON.parse(data || '[]');
        }
    } catch (err) {
        console.error(`Error reading ${filePath}:`, err);
    }
    return defaultValue;
}

// Helper to safely write JSON to disk
function writeJsonSafely(filePath, data) {
    try {
        fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
        return true;
    } catch (err) {
        console.error(`Error writing ${filePath}:`, err);
        return false;
    }
}

// Ensure deleted_cadets.json exists in database
if (!fs.existsSync(DELETED_CADETS_FILE)) {
    writeJsonSafely(DELETED_CADETS_FILE, []);
}

// ============================================================
// AUTHENTICATION & SESSION MANAGEMENT
// ============================================================
function hashPassword(password, salt) {
    if (!salt) salt = crypto.randomBytes(16).toString('hex');
    const hash = crypto.pbkdf2Sync(password, salt, 1000, 64, 'sha512').toString('hex');
    return { salt, hash };
}

function verifyPassword(password, salt, hash) {
    if (!salt || !hash) return false;
    const checkHash = crypto.pbkdf2Sync(password, salt, 1000, 64, 'sha512').toString('hex');
    return checkHash === hash;
}

// Helper to check if an email is authorized per credential data/authorized_emails.json
function isEmailAuthorized(email) {
    if (!email) return { authorized: false, reason: 'Email address is required.' };
    const cleanEmail = email.trim().toLowerCase();

    // 1. Basic format validation
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail)) {
        return { authorized: false, reason: 'Please enter a valid email address.' };
    }

    const authConfig = readJsonSafely(AUTHORIZED_EMAILS_FILE, {
        allowedDomains: ['@srcas.ac.in'],
        authorizedEmails: ['vignesh@srcas.ac.in', 'ncc@srcas.ac.in', 'admin@ncc.org']
    });

    const allowedDomains = (authConfig.allowedDomains || ['@srcas.ac.in']).map(d => d.trim().toLowerCase());
    const authorizedEmails = (authConfig.authorizedEmails || []).map(e => e.trim().toLowerCase());

    // 2. Check specifically listed authorized emails
    if (authorizedEmails.includes(cleanEmail)) {
        return { authorized: true, cleanEmail };
    }

    // 3. Check allowed domain (e.g. @srcas.ac.in)
    // Strictly checks that the email ends with the exact domain (e.g. @srcas.ac.in), rejecting fake domains
    const isDomainAllowed = allowedDomains.some(domain => {
        const d = domain.startsWith('@') ? domain : '@' + domain;
        return cleanEmail.endsWith(d);
    });

    if (isDomainAllowed) {
        return { authorized: true, cleanEmail };
    }

    return {
        authorized: false,
        reason: 'Access Denied: Registration is restricted to official @srcas.ac.in institutional emails and authorized accounts. Unauthorized email domains (e.g. Gmail, Yahoo, Outlook) are blocked.'
    };
}

// Generate a random cryptographically secure 6-digit OTP
function generateSixDigitOtp() {
    return crypto.randomInt(100000, 1000000).toString();
}

// Hash OTP with SHA-256 (raw OTP is never stored in DB)
function hashOtp(otp) {
    if (!otp) return '';
    return crypto.createHash('sha256').update(String(otp).trim()).digest('hex');
}

function parseJwtPayload(token) {
    try {
        const parts = token.split('.');
        if (parts.length < 2) return null;
        let base64Url = parts[1];
        let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        while (base64.length % 4) {
            base64 += '=';
        }
        const payload = Buffer.from(base64, 'base64').toString('utf8');
        return JSON.parse(payload);
    } catch (e) {
        return null;
    }
}

// In-memory session store: token -> { user, expiresAt }
const activeSessions = new Map();

// In-memory pending password changes for first-time login: tempToken -> { userId, email, expiresAt }
const pendingPasswordChanges = new Map();

function createSession(user) {
    const token = crypto.randomBytes(32).toString('hex');
    const safeUser = {
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role || 'officer',
        avatarUrl: user.avatarUrl || '',
        provider: user.provider || 'local'
    };
    // 7-day session validity
    activeSessions.set(token, {
        user: safeUser,
        expiresAt: Date.now() + 7 * 24 * 60 * 60 * 1000
    });
    return { token, user: safeUser };
}

function getSessionUser(token) {
    if (!token) return null;
    const session = activeSessions.get(token);
    if (!session) return null;
    if (session.expiresAt < Date.now()) {
        activeSessions.delete(token);
        return null;
    }
    return session.user;
}

// Authentication Middleware to protect administrative routes
function requireAuth(req, res, next) {
    const authHeader = req.headers.authorization;
    let token = null;
    if (authHeader && authHeader.startsWith('Bearer ')) {
        token = authHeader.substring(7);
    } else if (req.query.token) {
        token = req.query.token;
    }
    const user = getSessionUser(token);
    if (!user) {
        return res.status(401).json({ error: 'Authentication required. Please sign in.' });
    }
    req.user = user;
    next();
}

// Ensure credentials.json exists and migrate schema for OTP system
function migrateCredentialsSchema() {
    if (!fs.existsSync(USERS_FILE)) {
        writeJsonSafely(USERS_FILE, []);
        syncCredentialsSummary([]);
        return;
    }
    const users = readJsonSafely(USERS_FILE, []);
    let modified = false;
    users.forEach(u => {
        if (typeof u.email_verified !== 'boolean') {
            u.email_verified = false;
            modified = true;
        }
        if (!('otp_hash' in u)) {
            u.otp_hash = null;
            modified = true;
        }
        if (!('otp_expires_at' in u)) {
            u.otp_expires_at = null;
            modified = true;
        }
        if (!('otp_attempts' in u)) {
            u.otp_attempts = 0;
            modified = true;
        }
        if (!('last_otp_sent_at' in u)) {
            u.last_otp_sent_at = null;
            modified = true;
        }
        // Remove legacy plaintext password if salted PBKDF2 hash exists
        if (u.password && u.passwordHash && u.salt) {
            delete u.password;
            modified = true;
        }
    });
    if (modified) {
        writeJsonSafely(USERS_FILE, users);
    }
    syncCredentialsSummary(users);
}
migrateCredentialsSchema();

// Middleware to parse JSON and URL-encoded bodies
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Auth Config
app.get('/api/auth/config', (req, res) => {
    res.json({
        googleClientId: process.env.GOOGLE_CLIENT_ID || ''
    });
});

// Helper to synchronize plain text summary in credential data folder
function syncCredentialsSummary(users) {
    try {
        let content = "======================================================================\n";
        content += " NCC OD FORM GENERATOR - REGISTERED CREDENTIALS DIRECTORY\n";
        content += "======================================================================\n";
        content += "Directory Location: credential data/credentials.json\n";
        content += "Total Registered Accounts: " + users.length + "\n";
        content += "Policy: Authorized @srcas.ac.in & Listed Emails (6-Digit OTP Verified)\n";
        content += "Last Updated:  " + new Date().toLocaleString() + "\n\n";

        if (users.length === 0) {
            content += "----------------------------------------------------------------------\n";
            content += "No credentials currently stored.\n";
            content += "Authorized users must register and complete 6-digit OTP verification\n";
            content += "before signing in.\n";
            content += "----------------------------------------------------------------------\n\n";
        } else {
            users.forEach((u, i) => {
                const statusStr = u.email_verified 
                    ? "VERIFIED [ACTIVE]" 
                    : "UNVERIFIED (Pending OTP Verification)";
                content += "----------------------------------------------------------------------\n";
                content += "[" + (i + 1) + "] NAME:     " + (u.name || 'User') + "\n";
                content += "    Email:    " + u.email + "\n";
                content += "    Status:   " + statusStr + "\n";
                content += "    Role:     " + ((u.role || 'Officer').charAt(0).toUpperCase() + (u.role || 'Officer').slice(1)) + "\n";
                content += "    Created:  " + (u.createdAt || 'N/A') + "\n";
                content += "    Verified: " + (u.otp_verified_at || (u.email_verified ? 'Yes' : 'No (Pending)')) + "\n\n";
            });
        }

        content += "======================================================================\n";
        content += " INSTRUCTIONS:\n";
        content += " 1. Only @srcas.ac.in emails & authorized accounts can register.\n";
        content += " 2. Accounts require 6-digit OTP verification before first sign-in.\n";
        content += " 3. Passwords and OTPs are securely hashed.\n";
        content += "======================================================================\n";

        fs.writeFileSync(CREDENTIALS_SUMMARY_FILE, content, 'utf8');
    } catch (err) {
        console.error('Error syncing credentials summary:', err);
    }
}

// Auth Login - Strictly Email-Verified Users Only
app.post('/api/auth/login', (req, res) => {
    const { email, password } = req.body;
    if (!email || !password) {
        return res.status(400).json({ error: 'Please enter both your email address and password.' });
    }

    const cleanEmail = email.trim().toLowerCase();

    // Check whether email domain or address is authorized
    const authCheck = isEmailAuthorized(cleanEmail);
    if (!authCheck.authorized) {
        return res.status(403).json({ error: authCheck.reason });
    }

    const users = readJsonSafely(USERS_FILE, []);
    const user = users.find(u => (u.email || '').toLowerCase() === cleanEmail);

    if (!user) {
        return res.status(401).json({
            error: 'No registered account found with email "' + cleanEmail + '". Please switch to the "Create Account" tab to register.'
        });
    }

    const isPasswordValid = verifyPassword(password, user.salt, user.passwordHash) || 
                            (user.password && user.password === password);

    if (!isPasswordValid) {
        return res.status(401).json({
            error: 'Incorrect password. Please enter the password you created for this account.'
        });
    }

    // UPGRADE: Hash plain password legacy record if needed
    if (!user.passwordHash && isPasswordValid) {
        const { salt, hash } = hashPassword(password);
        user.salt = salt;
        user.passwordHash = hash;
        delete user.password;
        writeJsonSafely(USERS_FILE, users);
    }

    // LOGIN RESTRICTION: Check whether email verification is completed
    if (!user.email_verified) {
        return res.status(403).json({
            unverified: true,
            email: user.email,
            error: 'Your email address is not verified. Please verify your email before logging in.'
        });
    }

    user.lastLoginAt = new Date().toISOString();
    writeJsonSafely(USERS_FILE, users);
    syncCredentialsSummary(users);

    const session = createSession(user);
    res.json({ success: true, token: session.token, user: session.user });
});

// Auth Register - Strictly Authorizes Email & Dispatches 6-Digit OTP
app.post('/api/auth/register', async (req, res) => {
    try {
        const { name, email, password } = req.body;
        if (!email || !password) {
            return res.status(400).json({ error: 'Email and password are required.' });
        }

        const cleanEmail = email.trim().toLowerCase();

        // 1. Authorize Email against authorized_emails.json
        const authCheck = isEmailAuthorized(cleanEmail);
        if (!authCheck.authorized) {
            return res.status(403).json({ error: authCheck.reason });
        }

        if (password.length < 5) {
            return res.status(400).json({ error: 'Password must be at least 5 characters long.' });
        }

        const users = readJsonSafely(USERS_FILE, []);
        const existingIndex = users.findIndex(u => (u.email || '').toLowerCase() === cleanEmail);

        if (existingIndex >= 0) {
            const existing = users[existingIndex];
            if (existing.email_verified) {
                return res.status(400).json({
                    error: 'An account with email "' + cleanEmail + '" is already registered. Please switch to the Sign In tab.'
                });
            }

            // Account exists but is unverified: enforce 60s cooldown or issue fresh OTP
            const elapsed = Date.now() - new Date(existing.last_otp_sent_at || 0).getTime();
            if (elapsed < 60000) {
                const remainingSec = Math.ceil((60000 - elapsed) / 1000);
                return res.status(429).json({
                    unverified: true,
                    email: cleanEmail,
                    error: 'A verification OTP was recently sent. Please wait ' + remainingSec + ' seconds before requesting a new code.',
                    remainingSeconds: remainingSec
                });
            }

            // Issue new OTP and update password
            const { salt, hash } = hashPassword(password);
            const otp = generateSixDigitOtp();
            existing.name = (name || existing.name).trim();
            existing.salt = salt;
            existing.passwordHash = hash;
            delete existing.password;
            existing.otp_hash = hashOtp(otp);
            existing.otp_expires_at = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 minutes
            existing.otp_attempts = 0;
            existing.last_otp_sent_at = new Date().toISOString();

            writeJsonSafely(USERS_FILE, users);
            syncCredentialsSummary(users);

            const mailResult = await sendOtpEmail(cleanEmail, existing.name, otp);

            if (!mailResult.delivered) {
                return res.status(503).json({
                    success: false,
                    delivered: false,
                    smtpDelivered: false,
                    email: cleanEmail,
                    error: mailResult.error || 'Failed to deliver OTP email. SMTP service is not configured or unavailable.'
                });
            }

            return res.status(200).json({
                success: true,
                delivered: true,
                smtpDelivered: true,
                devMode: Boolean(mailResult.devMode),
                
                email_verified: false,
                email: cleanEmail,
                message: 'A fresh 6-digit verification code has been generated. Check your server terminal.',
            cooldownSeconds: 60,
            otpExpiresInMinutes: 10
            });
        }

        // New user registration
        const { salt, hash } = hashPassword(password);
        const otp = generateSixDigitOtp();
        const otpHash = hashOtp(otp);
        const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 minutes

        const newUser = {
            id: 'usr-' + Date.now().toString(36) + '-' + crypto.randomBytes(3).toString('hex'),
            name: (name || cleanEmail.split('@')[0]).trim(),
            email: cleanEmail,
            salt,
            passwordHash: hash,
            role: users.length === 0 ? 'admin' : 'officer',
            provider: 'local',
            mustChangePassword: false,
            avatarUrl: '',
            createdAt: new Date().toISOString(),
            lastLoginAt: null,
            email_verified: false,
            otp_hash: otpHash,
            otp_expires_at: expiresAt,
            otp_attempts: 0,
            last_otp_sent_at: new Date().toISOString()
        };

        // 1. Dispatch OTP email FIRST
        const mailResult = await sendOtpEmail(cleanEmail, newUser.name, otp);

        // 2. If email delivery fails, DO NOT add user to database - report error immediately
        if (!mailResult.delivered) {
            return res.status(503).json({
                success: false,
                delivered: false,
                smtpDelivered: false,
                email: cleanEmail,
                error: mailResult.error || 'Failed to deliver verification email. Please check your SMTP configuration in .env.'
            });
        }

        // 3. Email delivered successfully: Commit user to credentials database
        users.push(newUser);
        writeJsonSafely(USERS_FILE, users);
        syncCredentialsSummary(users);

        res.status(201).json({
            success: true,
            delivered: true,
            smtpDelivered: true,
            devMode: Boolean(mailResult.devMode),
            
            email_verified: false,
            email: cleanEmail,
            message: 'A 6-digit verification code has been generated. Check your server terminal.',
            cooldownSeconds: 60,
            otpExpiresInMinutes: 10
        });
    } catch (err) {
        console.error('Registration error:', err);
        res.status(500).json({ error: 'Registration error: ' + err.message });
    }
});

// Verify 6-Digit OTP - Validates OTP, enforces expiration (10m) & attempts limit, marks account as verified
app.post('/api/auth/verify-otp', (req, res) => {
    const { email, otp } = req.body;
    if (!email || !otp) {
        return res.status(400).json({ error: 'Email and 6-digit OTP code are required.' });
    }

    const cleanEmail = email.trim().toLowerCase();
    const cleanOtp = String(otp).trim();

    if (!/^\d{6}$/.test(cleanOtp)) {
        return res.status(400).json({ error: 'Invalid OTP format. OTP must be exactly 6 digits.' });
    }

    const users = readJsonSafely(USERS_FILE, []);
    const user = users.find(u => (u.email || '').toLowerCase() === cleanEmail);

    if (!user) {
        return res.status(404).json({ error: 'No registered account found with email "' + cleanEmail + '".' });
    }

    if (user.email_verified) {
        return res.status(400).json({ error: 'This account is already verified. You can sign in immediately.' });
    }

    if (!user.otp_hash || !user.otp_expires_at) {
        return res.status(400).json({ error: 'No active OTP verification session found. Please request a new OTP.' });
    }

    // 1. Check expiration (10 minutes)
    const now = Date.now();
    const expiresAt = new Date(user.otp_expires_at).getTime();
    if (now > expiresAt) {
        return res.status(400).json({
            error: 'Verification OTP has expired (codes expire in 10 minutes). Please request a new OTP.',
            expired: true
        });
    }

    // 2. Limit repeated incorrect OTP attempts (max 5 attempts)
    if ((user.otp_attempts || 0) >= 5) {
        return res.status(429).json({
            error: 'Too many incorrect OTP attempts. Please click "Resend OTP" to generate a fresh code.',
            maxAttempts: true
        });
    }

    // 3. Compare hashed OTP
    const inputHash = hashOtp(cleanOtp);
    if (inputHash !== user.otp_hash) {
        user.otp_attempts = (user.otp_attempts || 0) + 1;
        writeJsonSafely(USERS_FILE, users);
        const remainingAttempts = 5 - user.otp_attempts;
        return res.status(400).json({
            error: 'Incorrect OTP. Please check your email and try again (' + remainingAttempts + ' attempts remaining).'
        });
    }

    // 4. Mark verified and invalidate OTP (strictly single-use)
    user.email_verified = true;
    user.otp_hash = null;
    user.otp_expires_at = null;
    user.otp_attempts = 0;
    user.otp_verified_at = new Date().toISOString();

    writeJsonSafely(USERS_FILE, users);
    syncCredentialsSummary(users);

    console.log(`[AUTH] User verified successfully via OTP: ${user.email}`);

    res.json({
        success: true,
        message: 'Email verified successfully! You can now sign in with your credentials.',
        email: user.email
    });
});

// Resend 6-Digit OTP - with 60-second cooldown, invalidates previous OTP
app.post('/api/auth/resend-otp', async (req, res) => {
    try {
        const { email } = req.body;
        if (!email) {
            return res.status(400).json({ error: 'Email address is required.' });
        }

        const cleanEmail = email.trim().toLowerCase();
        const users = readJsonSafely(USERS_FILE, []);
        const user = users.find(u => (u.email || '').toLowerCase() === cleanEmail);

        if (!user) {
            return res.status(404).json({ error: 'No account found with email "' + cleanEmail + '". Please register first.' });
        }

        if (user.email_verified) {
            return res.status(400).json({ error: 'This email is already verified. You can sign in immediately.' });
        }

        // 1. Enforce 60-second cooldown
        if (user.last_otp_sent_at) {
            const elapsedMs = Date.now() - new Date(user.last_otp_sent_at).getTime();
            const cooldownMs = 60 * 1000;
            if (elapsedMs < cooldownMs) {
                const remainingSec = Math.ceil((cooldownMs - elapsedMs) / 1000);
                return res.status(429).json({
                    error: 'Please wait ' + remainingSec + ' second' + (remainingSec === 1 ? '' : 's') + ' before requesting another OTP.',
                    remainingSeconds: remainingSec
                });
            }
        }

        // 2. Generate fresh 6-digit OTP & invalidate previous
        const freshOtp = generateSixDigitOtp();
        user.otp_hash = hashOtp(freshOtp);
        user.otp_expires_at = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 minutes
        user.otp_attempts = 0;
        user.last_otp_sent_at = new Date().toISOString();

        writeJsonSafely(USERS_FILE, users);
        syncCredentialsSummary(users);

        const mailResult = await sendOtpEmail(user.email, user.name, freshOtp);

        if (!mailResult.delivered) {
            return res.status(503).json({
                success: false,
                delivered: false,
                smtpDelivered: false,
                email: user.email,
                error: mailResult.error || 'Failed to deliver OTP email. SMTP service is not configured or unavailable.'
            });
        }

        res.json({
            success: true,
            delivered: true,
            smtpDelivered: true,
            devMode: Boolean(mailResult.devMode),
            
            email: user.email,
            message: 'A fresh 6-digit verification code has been generated. Check your server terminal.',
            cooldownSeconds: 60,
            otpExpiresInMinutes: 10
        });
    } catch (err) {
        console.error('Error resending OTP:', err);
        res.status(500).json({ error: 'Failed to resend OTP: ' + err.message });
    }
});

// Endpoint to inspect latest dispatched email in development/testing
app.get('/api/auth/dev-latest-email', (req, res) => {
    if (process.env.NODE_ENV === 'production') {
        return res.status(403).json({ error: 'Not available in production' });
    }
    res.json(getLatestDevEmail() || { message: 'No email dispatched yet' });
});

// Auth Status / Current User
app.get('/api/auth/me', (req, res) => {
    const authHeader = req.headers.authorization;
    let token = null;
    if (authHeader && authHeader.startsWith('Bearer ')) {
        token = authHeader.substring(7);
    } else if (req.query.token) {
        token = req.query.token;
    }

    const user = getSessionUser(token);
    if (user) {
        return res.json({ authenticated: true, user });
    }
    return res.status(401).json({ authenticated: false, error: 'Unauthorized or session expired' });
});

// Auth Logout
app.post('/api/auth/logout', (req, res) => {
    const authHeader = req.headers.authorization;
    let token = null;
    if (authHeader && authHeader.startsWith('Bearer ')) {
        token = authHeader.substring(7);
    } else if (req.body && req.body.token) {
        token = req.body.token;
    }
    if (token) {
        activeSessions.delete(token);
    }
    res.json({ success: true });
});

// Endpoint to check voice configuration
app.get('/api/voice-config', (req, res) => {
    res.json({
        provider: 'CadetAssist',
        projectId: process.env.CADET_ASSIST_PROJECT_ID || 'p-3ce6a60b',
        configured: Boolean(process.env.CADET_ASSIST_API_KEY)
    });
});

// Serve static frontend files with caching disabled for local development
app.use(express.static(FRONTEND_DIR, {
    setHeaders: (res, path) => {
        res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    }
}));

// Endpoint to fetch active cadets from cadets.json
app.get('/api/cadets', (req, res) => {
    fs.readFile(CADETS_FILE, 'utf8', (err, data) => {
        if (err) {
            if (err.code === 'ENOENT') {
                // If file doesn't exist, return empty database
                return res.json([]);
            }
            console.error('Error reading database file:', err);
            return res.status(500).json({ error: 'Failed to read database file' });
        }
        try {
            const json = JSON.parse(data || '[]');
            res.json(json);
        } catch (parseErr) {
            console.error('Error parsing database JSON:', parseErr);
            res.status(500).json({ error: 'Invalid database format' });
        }
    });
});

// Endpoint to fetch deleted / archived cadets from deleted_cadets.json
app.get('/api/cadets/deleted', (req, res) => {
    const deletedCadets = readJsonSafely(DELETED_CADETS_FILE, []);
    res.json(deletedCadets);
});

// Endpoint to manually delete a cadet: removes from cadets.json and permanently preserves in deleted_cadets.json
app.post('/api/cadets/manual-delete', (req, res) => {
    const { cadet, reason } = req.body;
    if (!cadet || (!cadet.register_no && !cadet.name)) {
        return res.status(400).json({ error: 'Cadet details required' });
    }

    const currentCadets = readJsonSafely(CADETS_FILE, []);
    const cadetKey = (cadet.register_no || cadet.name || '').trim().toLowerCase();

    // Find and remove from active cadets
    const targetIdx = currentCadets.findIndex(c => (c.register_no || c.name || '').trim().toLowerCase() === cadetKey);
    const deletedCadet = targetIdx !== -1 ? currentCadets.splice(targetIdx, 1)[0] : cadet;

    // Archive in deleted_cadets.json
    const deletedList = readJsonSafely(DELETED_CADETS_FILE, []);
    const archiveRecord = {
        name: deletedCadet.name,
        register_no: deletedCadet.register_no || '',
        department: deletedCadet.department || '',
        shift: deletedCadet.shift || '',
        deleted_at: new Date().toISOString(),
        deleted_by: 'manual_admin',
        reason: reason || 'Manual deletion by user'
    };
    deletedList.unshift(archiveRecord);

    writeJsonSafely(CADETS_FILE, currentCadets);
    writeJsonSafely(DELETED_CADETS_FILE, deletedList);

    console.log(`Cadet "${deletedCadet.name}" manually removed and safely archived in database/deleted_cadets.json.`);
    res.json({
        success: true,
        message: `Cadet "${deletedCadet.name}" removed from active list and preserved in database archive`,
        archived: archiveRecord,
        remainingCount: currentCadets.length
    });
});

// Endpoint to restore a previously deleted cadet back to active list
app.post('/api/cadets/restore', (req, res) => {
    const { register_no, name } = req.body;
    const deletedList = readJsonSafely(DELETED_CADETS_FILE, []);
    const key = (register_no || name || '').trim().toLowerCase();

    const idx = deletedList.findIndex(c => (c.register_no || c.name || '').trim().toLowerCase() === key);
    if (idx === -1) {
        return res.status(404).json({ error: 'Cadet not found in database archive' });
    }

    const [restored] = deletedList.splice(idx, 1);
    const activeCadets = readJsonSafely(CADETS_FILE, []);

    // Check if cadet is already back in active list
    const alreadyExists = activeCadets.some(c => (c.register_no || c.name || '').trim().toLowerCase() === key);
    if (!alreadyExists) {
        activeCadets.push({
            name: restored.name,
            register_no: restored.register_no,
            department: restored.department,
            shift: restored.shift
        });
    }

    writeJsonSafely(CADETS_FILE, activeCadets);
    writeJsonSafely(DELETED_CADETS_FILE, deletedList);

    console.log(`Cadet "${restored.name}" restored to active list.`);
    res.json({
        success: true,
        message: `Cadet "${restored.name}" successfully restored from database archive`,
        cadets: activeCadets
    });
});

// Endpoint to write cadets to cadets.json with automatic preservation of removed records in deleted_cadets.json
app.post('/api/cadets', (req, res) => {
    const newCadets = req.body;
    if (!Array.isArray(newCadets)) {
        return res.status(400).json({ error: 'Data must be an array of cadets' });
    }

    // Safety check: detect any cadets that were removed and archive them in deleted_cadets.json
    const previousCadets = readJsonSafely(CADETS_FILE, []);
    const newKeys = new Set(newCadets.map(c => (c.register_no || c.name || '').trim().toLowerCase()));
    const removedCadets = previousCadets.filter(c => !newKeys.has((c.register_no || c.name || '').trim().toLowerCase()));

    if (removedCadets.length > 0) {
        const deletedList = readJsonSafely(DELETED_CADETS_FILE, []);
        removedCadets.forEach(rc => {
            const rKey = (rc.register_no || rc.name || '').trim().toLowerCase();
            const existsInArchive = deletedList.some(d => (d.register_no || d.name || '').trim().toLowerCase() === rKey);
            if (!existsInArchive) {
                deletedList.unshift({
                    name: rc.name,
                    register_no: rc.register_no || '',
                    department: rc.department || '',
                    shift: rc.shift || '',
                    deleted_at: new Date().toISOString(),
                    deleted_by: 'manual_sync_removal',
                    reason: 'Removed during cadet list update'
                });
            }
        });
        writeJsonSafely(DELETED_CADETS_FILE, deletedList);
        console.log(`Preserved ${removedCadets.length} removed cadet(s) into database/deleted_cadets.json.`);
    }

    fs.writeFile(CADETS_FILE, JSON.stringify(newCadets, null, 2), 'utf8', (err) => {
        if (err) {
            console.error('Error writing database file:', err);
            return res.status(500).json({ error: 'Failed to save database file' });
        }
        console.log(`Database updated successfully with ${newCadets.length} cadets.`);
        res.json({ success: true, count: newCadets.length, archivedCount: removedCadets.length });
    });
});

// Endpoint to generate and download Word document on the server (bypasses browser security limits)
app.post('/api/generate-docx', async (req, res) => {
    try {
        let payload = req.body;
        if (req.body.payload) {
            payload = JSON.parse(req.body.payload);
        }

        const { selectedCadets, particularsVal, odReasonVal, selectedDate, docConfig } = payload;

        if (!selectedCadets || selectedCadets.length === 0) {
            return res.status(400).send("No cadets selected to export.");
        }

        selectedCadets.sort(sortCadetsForOD);

        // Build Table Rows
        const tableRows = [];

        // 1. Table Header Row
        tableRows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ]
        }));

        // 2. Add Cadet Rows
        selectedCadets.forEach((cadet, index) => {
            const rowCells = [
                // Col 0: S.No
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 1: Reg No
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 2: Name (left aligned)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 3: Dept (left aligned)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department ? " " + cadet.department : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 4: Shift
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 5: Particulars (vertically merged)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? particularsVal : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                    verticalMerge: index === 0 ? "restart" : "continue"
                }),
                // Col 6: On Duty (vertically merged)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? odReasonVal : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                    verticalMerge: index === 0 ? "restart" : "continue"
                })
            ];

            tableRows.push(new TableRow({ children: rowCells }));
        });

        // Date String helper
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(selectedDate);
        const dayStr = String(d.getDate()).padStart(2, '0');
        const monthStr = months[d.getMonth()];
        const yearStr = d.getFullYear();
        const formattedDateStr = `${dayStr} ${monthStr} ${yearStr}`;

        // 3. Build Document Structure
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                        }
                    }
                },
                children: [
                    // College Header
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.college,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // NCC Subtitle
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.unit,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Title Header
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 360 },
                        children: [
                            new TextRun({
                                text: docConfig.formType,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Main Grid Table
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        rows: tableRows
                    }),
                    // Space before signatures
                    new Paragraph({
                        spacing: { before: 1440 },
                        children: [new TextRun("")]
                    }),
                    // Signatures Grid
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            left: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            right: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.LEFT,
                                            children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating docx on server:", err);
        res.status(500).send("Error generating document: " + err.message);
    }
});

// ============================================================
// /api/generate-and-open  — saves OD form to Desktop & opens it in MS Word
// ============================================================
app.post('/api/generate-and-open', async (req, res) => {
    try {
        if (process.env.VERCEL || process.env.RENDER || process.env.NODE_ENV === 'production') {
            return res.status(400).json({
                error: 'The "Open in Word" feature is only supported when running the app locally on your computer. Please use the "Download (.docx)" button instead.'
            });
        }
        const { selectedCadets, particularsVal, odReasonVal, selectedDate, docConfig: cfg } = req.body;

        if (!selectedCadets || selectedCadets.length === 0) {
            return res.status(400).json({ error: 'No cadets selected.' });
        }

        const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        const d = new Date(selectedDate);
        const formattedDate = `${String(d.getDate()).padStart(2, '0')} ${months[d.getMonth()]} ${d.getFullYear()}`;

        const buffer = await buildODDocument(selectedCadets, particularsVal, odReasonVal, cfg);

        // Save to Desktop
        const desktopPath = path.join(require('os').homedir(), 'OneDrive', 'Desktop');
        const filename = `od-form-${formattedDate.toLowerCase().replace(/ /g, '-')}.docx`;
        const filePath = path.join(desktopPath, filename);
        fs.writeFileSync(filePath, buffer);

        // Auto-open in MS Word
        exec(`start "" "${filePath}"`, { shell: true }, (err) => {
            if (err) console.error('Auto-open failed:', err);
        });

        console.log(`OD Form saved and opened: ${filePath}`);
        res.json({ success: true, filename, path: filePath });

    } catch (err) {
        console.error('Error in generate-and-open:', err);
        res.status(500).json({ error: err.message });
    }
});


// Duplicate attendance routes removed - using consolidated routes below

// Endpoint to generate attendance sheet docx
app.post('/api/generate-attendance-sheet-docx', async (req, res) => {
    try {
        let payload = req.body;
        if (req.body.payload) {
            payload = JSON.parse(req.body.payload);
        }

        const { date, records, docConfig } = payload;

        if (!records || records.length === 0) {
            return res.status(400).send("No records to export.");
        }

        // Build Table Rows
        const tableRows = [];

        // 1. Table Header Row
        tableRows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ]
        }));

        // 2. Add Cadet Records Rows
        records.forEach((cadet, index) => {
            const statusText = (cadet.status || "PRESENT").toUpperCase();
            const rowCells = [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department ? " " + cadet.department : "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusText, font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ];

            tableRows.push(new TableRow({ children: rowCells }));
        });

        // Date String helper
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(date);
        const dayStr = String(d.getDate()).padStart(2, '0');
        const monthStr = months[d.getMonth()];
        const yearStr = d.getFullYear();
        const formattedDateStr = `${dayStr} ${monthStr} ${yearStr}`;

        // 3. Build Document Structure
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                        }
                    }
                },
                children: [
                    // College Header
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.college,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // NCC Subtitle
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.unit,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Title Header
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 360 },
                        children: [
                            new TextRun({
                                text: `ATTENDANCE SHEET - ${formattedDateStr}`,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Main Grid Table
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        rows: tableRows
                    }),
                    // Space before signatures
                    new Paragraph({
                        spacing: { before: 1440 },
                        children: [new TextRun("")]
                    }),
                    // Signatures Grid
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            left: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            right: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.LEFT,
                                            children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `attendance-${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating attendance docx:", err);
        res.status(500).send("Error generating document: " + err.message);
    }
});

// ============================================================
// ATTENDANCE ROUTES
// ============================================================

// Helper: read attendance database
function readAttendanceDB() {
    try {
        if (fs.existsSync(ATTENDANCE_FILE)) {
            return JSON.parse(fs.readFileSync(ATTENDANCE_FILE, 'utf8') || '{}');
        }
    } catch (e) {
        console.error("Error reading attendance DB:", e);
    }
    return {};
}

// Helper: write attendance database
function writeAttendanceDB(data) {
    fs.writeFileSync(ATTENDANCE_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// GET /api/attendance?date=YYYY-MM-DD  - get one day's record
app.get('/api/attendance', (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date query param required' });
    const db = readAttendanceDB();
    const record = db[date] || { date, event_name: '', od_reason: '', records: {} };
    res.json(record);
});

// POST /api/attendance - save or update one day's record
app.post('/api/attendance', (req, res) => {
    const { date, event_name, od_reason, records } = req.body;
    if (!date) return res.status(400).json({ error: 'date required' });
    const db = readAttendanceDB();
    db[date] = { date, event_name: event_name || '', od_reason: od_reason || '', records: records || {} };
    writeAttendanceDB(db);
    console.log(`Attendance saved for ${date}`);
    res.json({ success: true });
});

// GET /api/attendance/history - get summary list of all recorded sessions
app.get('/api/attendance/history', (req, res) => {
    const db = readAttendanceDB();
    const history = Object.keys(db)
        .sort((a, b) => b.localeCompare(a)) // newest first
        .map(date => {
            const rec = db[date];
            const counts = { present: 0, absent: 0, od: 0, total: 0 };
            if (rec.records) {
                for (const key in rec.records) {
                    counts.total++;
                    const s = rec.records[key];
                    if (s === 'present') counts.present++;
                    else if (s === 'absent') counts.absent++;
                    else if (s === 'od') counts.od++;
                }
            }
            return {
                date,
                event_name: rec.event_name,
                od_reason: rec.od_reason,
                records: rec.records,
                counts
            };
        });
    res.json(history);
});

// DELETE /api/attendance?date=YYYY-MM-DD - remove a day's record
app.delete('/api/attendance', (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date query param required' });
    const db = readAttendanceDB();
    if (db[date]) {
        delete db[date];
        writeAttendanceDB(db);
        console.log(`Attendance record deleted for ${date}`);
        res.json({ success: true });
    } else {
        res.status(404).json({ error: 'No record found for that date' });
    }
});

// ============================================================
// DIRECT DOWNLOAD SHORTCUTS (accessible via browser URL bar)
// ============================================================

// GET /download/attendance?date=YYYY-MM-DD
// Directly generates and downloads the attendance sheet Word doc
app.get('/download/attendance', async (req, res) => {
    try {
        const date = req.query.date || new Date().toISOString().split('T')[0];

        // Load cadets
        let cadets = [];
        try { cadets = JSON.parse(fs.readFileSync(CADETS_FILE, 'utf8') || '[]'); } catch (e) { }

        if (cadets.length === 0) {
            return res.status(400).send('No cadets in database. Please add cadets first.');
        }

        // Load attendance records for this date
        const db = readAttendanceDB();
        const attRecord = db[date] || { records: {} };

        // Build records array with status
        const records = cadets.map(c => ({
            register_no: c.register_no,
            name: c.name,
            department: c.department,
            shift: c.shift,
            status: attRecord.records[c.register_no ? c.register_no.trim() : c.name.trim().toLowerCase()] || 'present'
        }));

        // Date formatting
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(date);
        const formattedDateStr = `${String(d.getDate()).padStart(2, '0')} ${months[d.getMonth()]} ${d.getFullYear()}`;

        // Build table rows
        const tableRows = [];

        // Header row
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            ]
        }));

        // Data rows
        records.forEach((cadet, index) => {
            const statusText = (cadet.status || 'PRESENT').toUpperCase();
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + (cadet.department || ""), font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusText, font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                ]
            }));
        });

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: "SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)", font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "NATIONAL CADET CORPS", font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: `ATTENDANCE SHEET - ${formattedDateStr}`, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 },
                            left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 },
                        },
                        rows: [new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "NCC Officer", font: "Bookman Old Style", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Principal", font: "Bookman Old Style", size: 24 })] })] }),
                            ]
                        })]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `attendance-${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating direct download attendance:", err);
        res.status(500).send("Error: " + err.message);
    }
});

// ============================================================
// /api/generate-for-date  — Quick-generate OD form for ALL cadets for a date
// ============================================================
app.post('/api/generate-for-date', async (req, res) => {
    try {
        if (process.env.VERCEL || process.env.RENDER || process.env.NODE_ENV === 'production') {
            return res.status(400).json({
                error: 'The "Quick Gen" (save to Desktop) feature is only supported when running the app locally on your computer. Please select cadets and use the "Download (.docx)" button instead.'
            });
        }
        const { date, particulars, odReason } = req.body;
        const selectedDate = date || new Date().toISOString().split('T')[0];

        // Load all cadets
        let allCadets = [];
        try { allCadets = JSON.parse(fs.readFileSync(CADETS_FILE, 'utf8') || '[]'); } catch (e) { }

        if (allCadets.length === 0) {
            return res.status(400).json({ error: 'No cadets in database. Please add cadets first.' });
        }

        // Date formatting
        const months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
        const monthsShort = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(selectedDate);
        const day = d.getDate();
        const suffix = day === 1 ? 'ST' : day === 2 ? 'ND' : day === 3 ? 'RD' : 'TH';
        const dateStr = `${day}${suffix} ${months[d.getMonth()]} ${d.getFullYear()}`;
        const formattedDate = `${String(day).padStart(2, '0')} ${monthsShort[d.getMonth()]} ${d.getFullYear()}`;

        const particularsVal = particulars || 'NCC ACTIVITY';
        const odReasonVal = odReason || `EXEMPTION FROM ALL ACADEMIC WORKS AND REGULAR CLASSES ON ${dateStr}`;

        const college = 'SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)';
        const unit = 'NATIONAL CADET CORPS';
        const formType = 'OD FORM';
        const sigLeft = 'NCC Officer';
        const sigRight = 'Principal';

        // Build table rows
        const tableRows = [];

        // Header
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
            ]
        }));

        const N = allCadets.length;
        let P = 1;
        if (N > 27) {
            P = Math.ceil((N - 27) / 32) + 1;
        }
        const midPage = Math.ceil(P / 2);
        const startRow = midPage === 1 ? 0 : 27 + (midPage - 2) * 32;

        // Data rows
        allCadets.forEach((cadet, index) => {
            let isPartRestart = false;
            let partVal = "";
            let isOdRestart = false;
            let odVal = "";

            if (startRow > 0) {
                if (index === 0) {
                    isPartRestart = true;
                    partVal = "";
                    isOdRestart = true;
                    odVal = "";
                } else if (index === startRow) {
                    isPartRestart = true;
                    partVal = particularsVal;
                    isOdRestart = true;
                    odVal = odReasonVal;
                }
            } else {
                if (index === 0) {
                    isPartRestart = true;
                    partVal = particularsVal;
                    isOdRestart = true;
                    odVal = odReasonVal;
                }
            }

            const partMerge = startRow > 0
                ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                : (index === 0 ? 'restart' : 'continue');

            const odMerge = startRow > 0
                ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                : (index === 0 ? 'restart' : 'continue');

            const partCell = new TableCell({
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isPartRestart ? partVal : "", font: "Cambria", size: 24 })] })],
                verticalAlign: VerticalAlign.CENTER,
                verticalMerge: partMerge
            });

            const odCell = new TableCell({
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isOdRestart ? odVal : "", font: "Cambria", size: 24 })] })],
                verticalAlign: VerticalAlign.CENTER,
                verticalMerge: odMerge
            });

            // Fix continue cells: remove wrong val attribute, keep element empty
            if (partMerge === 'continue') {
                const tcPr = partCell.root.find(el => el && el.rootKey === 'w:tcPr');
                if (tcPr) {
                    const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                    if (vMergeEl) vMergeEl.root = [];
                }
            }
            if (odMerge === 'continue') {
                const tcPr = odCell.root.find(el => el && el.rootKey === 'w:tcPr');
                if (tcPr) {
                    const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                    if (vMergeEl) vMergeEl.root = [];
                }
            }

            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + (cadet.department || ""), font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    partCell,
                    odCell,
                ]
            }));
        });

        // Build document
        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: college, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: unit, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: formType, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 },
                            left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 },
                        },
                        rows: [new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: sigLeft, font: "Bookman Old Style", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sigRight, font: "Bookman Old Style", size: 24 })] })] }),
                            ]
                        })]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);

        // Save to Desktop
        const desktopPath = path.join(require('os').homedir(), 'OneDrive', 'Desktop');
        const filename = `od-form-${formattedDate.toLowerCase().replace(/ /g, '-')}.docx`;
        const filePath = path.join(desktopPath, filename);
        fs.writeFileSync(filePath, buffer);

        // Auto-open in MS Word
        exec(`start "" "${filePath}"`, { shell: true }, (err) => {
            if (err) console.error('Auto-open failed:', err);
        });

        console.log(`Quick OD Form generated for ${allCadets.length} cadets: ${filePath}`);
        res.json({ success: true, filename, path: filePath, cadetCount: allCadets.length });

    } catch (err) {
        console.error("Error in generate-for-date:", err);
        res.status(500).json({ error: err.message });
    }
});

// Serve index.html for the root route
app.get('/', (req, res) => {
    res.sendFile(path.join(FRONTEND_DIR, 'index.html'));
});

// Safe SMTP Diagnostics API Endpoint
app.get('/api/auth/smtp-status', (req, res) => {
    const diag = getSafeSmtpDiagnostics();
    res.json(diag);
});

// Start the server (only if not running on serverless Vercel environment)
if (!process.env.VERCEL) {
    app.listen(PORT, () => {
        const url = `http://localhost:${PORT}`;
        const diag = getSafeSmtpDiagnostics();

        console.log(`==================================================`);
        console.log(` NCC OD Form Automator server running at: ${url}`);
        console.log(`--------------------------------------------------`);
        console.log(` [ENV] Config file: ${diag.envPath}`);
        console.log(` [ENV] File exists: ${diag.envExists ? 'YES' : 'NO'}`);
        console.log(` [ENV] Mode: ${diag.devOtpMode ? 'DEV MODE (Terminal Console Output)' : 'LIVE EMAIL (SMTP Dispatch)'}`);
        console.log(` [SMTP] Host: ${diag.smtpHost}:${diag.smtpPort} (SSL: ${diag.smtpSecure})`);
        console.log(` [SMTP] User: ${diag.smtpUser}`);
        console.log(` [SMTP] Password: ${diag.smtpPass}`);
        console.log(`==================================================`);

        if (diag.isConfigured) {
            verifySmtpConnection().then(res => {
                if (res.ok) {
                    console.log(' [SMTP] Mail server connection verified successfully.');
                } else {
                    console.warn(` [SMTP WARNING] Mail server connection test failed: ${res.error}`);
                }
            }).catch(err => {
                console.warn(` [SMTP WARNING] Mail server connection test failed: ${err.message}`);
            });
        } else {
            console.log(` [SMTP] Live email delivery is awaiting SMTP_USER & SMTP_PASS in .env.`);
        }
    });
}

module.exports = app;
