#!/usr/bin/env node
/**
 * NCC OD Form Generator - Safe SMTP Configuration & Diagnostics CLI
 * Run: node backend/check_smtp.js
 */
const path = require('path');
const rootEnvPath = path.resolve(__dirname, '..', '.env');
require('dotenv').config({ path: rootEnvPath, override: true });

const { getSafeSmtpDiagnostics, verifySmtpConnection } = require('./mailer');

async function main() {
    console.log('======================================================================');
    console.log('       NCC OD FORM GENERATOR - SAFE SMTP DIAGNOSTIC AUDIT');
    console.log('======================================================================');

    const diag = getSafeSmtpDiagnostics();

    console.log('\n[1] ENVIRONMENT FILE LOCATION:');
    console.log(`    Target Path: ${diag.envPath}`);
    console.log(`    File Exists: ${diag.envExists ? 'YES (Detected)' : 'NO (Missing)'}`);

    console.log('\n[2] ENVIRONMENT VARIABLES DETECTED:');
    console.log(`    DEV_OTP_MODE: ${diag.devOtpMode ? 'true (Local Test Mode)' : 'false (Live Email Mode)'}`);
    console.log(`    SMTP_HOST:    ${diag.smtpHost}`);
    console.log(`    SMTP_PORT:    ${diag.smtpPort}`);
    console.log(`    SMTP_SECURE:  ${diag.smtpSecure}`);
    console.log(`    SMTP_USER:    ${diag.smtpUser}`);
    console.log(`    SMTP_PASS:    ${diag.smtpPass}`);
    console.log(`    SMTP_FROM:    ${diag.smtpFrom}`);

    console.log('\n[3] CONFIGURATION STATUS:');
    if (!diag.isConfigured) {
        console.log('    STATUS: ❌ UNCONFIGURED');
        console.log('    REASON: SMTP_USER and/or SMTP_PASS are empty in .env');
        console.log('\n    NEXT STEP:');
        console.log(`    1. Open: ${diag.envPath}`);
        console.log('    2. Fill in SMTP_USER=your-email@gmail.com');
        console.log('    3. Fill in SMTP_PASS=your-16-char-app-password');
        console.log('    4. Save the file and rerun: node backend/check_smtp.js');
    } else {
        console.log('    STATUS: ✅ CREDENTIALS FOUND IN .ENV');
        console.log('\n[4] TESTING SMTP CONNECTION TO MAIL SERVER...');
        try {
            const result = await verifySmtpConnection();
            if (result.ok) {
                console.log('    RESULT: ✅ SMTP Server connected and verified successfully!');
                console.log('    The backend is ready to send verification emails.');
            } else {
                console.log(`    RESULT: ❌ Connection test failed: ${result.error}`);
            }
        } catch (err) {
            console.log(`    RESULT: ❌ Unexpected error: ${err.message}`);
        }
    }
    console.log('======================================================================\n');
}

main();
