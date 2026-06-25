/**
 * myDoc - Authentication & RBAC Functions
 */

function getUserAccess(clientArg) {
  const activeUser = Session.getActiveUser();
  let email = activeUser ? activeUser.getEmail() : '';

  // Fallback if session email is empty (Execute as Me)
  if (!email) {
    // Try to validate as token first
    const tokenEmail = validateSessionToken(clientArg);
    if (tokenEmail) {
      email = tokenEmail;
      Logger.log('Using validated session token. Email: ' + email);
    } else {
      // Fallback to raw email (Manual Login Debug)
      email = clientArg || 'NO_EMAIL';
      Logger.log('Using raw client arg (Token invalid): ' + email);
    }
  }

  // STRICT SECURITY CHECK: Reject "NO_EMAIL" immediately
  if (!email || email === 'NO_EMAIL') {
    Logger.log('Identity Unknown. Returning NONE role.');
    return {
      email: '',
      role: 'NONE',
      modules: []
    };
  }
  const cleanEmail = email.trim().toLowerCase();

  let role = getUserRole(email);

  // Emergency Fallback: If email is in AUTHORIZED_USERS as ADMIN, force role to ADMIN
  for (let i = 0; i < AUTHORIZED_USERS.length; i++) {
    const authUser = AUTHORIZED_USERS[i];
    const checkEmail = authUser.email.trim().toLowerCase();

    if (checkEmail === cleanEmail) {
      if (authUser.role.trim().toUpperCase() === 'ADMIN' && authUser.status.trim().toLowerCase() === 'active') {
        Logger.log('!!! EMERGENCY OVERRIDE: Forcing ADMIN role for ' + cleanEmail);
        role = 'ADMIN';
      }
      break;
    }
  }

  // Normalize role for comparison
  const cleanRole = (role || 'NONE').toString().trim().toUpperCase();
  Logger.log('=== getUserAccess === Email: ' + email + ' | Role: ' + cleanRole);

  // Define module access by role
  const rolePermissions = {
    'ADMIN': ['dashboard', 'registration', 'triage', 'diagnosis', 'lab', 'lab-inventory', 'finance', 'reports', 'family-planning', 'patient-history'],
    'DOCTOR': ['dashboard', 'triage', 'diagnosis', 'reports', 'family-planning', 'patient-history'],
    'NURSE': ['registration', 'triage'],
    'LAB TECH': ['lab', 'lab-inventory'],
    'FINANCE': ['finance', 'patient-history']
  };

  const modules = rolePermissions[cleanRole] || [];

  return {
    email: email,
    role: cleanRole,
    modules: modules
  };
}

function getUserRole(email) {
  Logger.log('=== getUserRole called for: ' + email);
  if (!email) return null;

  const cleanEmail = email.trim().toLowerCase();

  // PRIORITY 1: Check hardcoded AUTHORIZED_USERS list first
  for (let i = 0; i < AUTHORIZED_USERS.length; i++) {
    const user = AUTHORIZED_USERS[i];
    if (user.email.trim().toLowerCase() === cleanEmail) {
      if (user.status.trim().toLowerCase() === 'active') {
        Logger.log('✓ Found user in hardcoded list. Role: ' + user.role);
        return user.role.trim();
      } else {
        Logger.log('⚠ Found user in hardcoded list BUT status is: ' + user.status);
      }
    }
  }
  Logger.log('User not found in active hardcoded list for: ' + cleanEmail);

  // PRIORITY 2: Check spreadsheet (fallback)
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Authorized_Users');

    if (!sheet) {
      Logger.log('Sheet "Authorized_Users" not found.');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0] || '').trim().toLowerCase();
      const rowStatus = String(data[i][5] || '').trim().toLowerCase();

      if (rowEmail === cleanEmail && rowStatus === 'active') {
        Logger.log('✓ Found user in spreadsheet. Role: ' + data[i][2]);
        return String(data[i][2] || '').trim();
      }
    }
  } catch (error) {
    Logger.log('⚠️ Spreadsheet check failed: ' + error.message);
  }

  return null;
}

function isAuthorizedUser(email) {
  Logger.log('=== isAuthorizedUser called ===');
  const cleanEmail = email.trim().toLowerCase();
  Logger.log('Checking email: ' + cleanEmail);

  // PRIORITY 1: Check hardcoded AUTHORIZED_USERS list
  for (let i = 0; i < AUTHORIZED_USERS.length; i++) {
    const user = AUTHORIZED_USERS[i];
    const userEmail = user.email.trim().toLowerCase();

    if (userEmail === cleanEmail && user.status.trim() === 'Active') {
      Logger.log('✓ User authorized via hardcoded list!');
      return true;
    }
  }

  // PRIORITY 2: Check spreadsheet
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Authorized_Users');

    if (!sheet) {
      // Logic to create sheet remains same...
      return false; // For now return false if not found
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0] || '').trim().toLowerCase();
      const rowStatus = String(data[i][5] || '').trim();

      if (rowEmail === cleanEmail && rowStatus === 'Active') {
        Logger.log('✓ User authorized via spreadsheet!');
        return true;
      }
    }

    Logger.log('✗ User NOT authorized: ' + cleanEmail);
    return false;
  } catch (error) {
    Logger.log('⚠️ Spreadsheet check failed: ' + error.message);
    return false;
  }
}

function authenticateUser(clientEmail) {
  Logger.log('=== authenticateUser called ===');
  let userEmail = Session.getActiveUser().getEmail();

  // Fallback to client email
  if (!userEmail) {
    userEmail = clientEmail || '';
    Logger.log('Session email empty. Using client email: ' + userEmail);
  } else {
    Logger.log('Using trusted Session email: ' + userEmail);
  }

  if (!userEmail) {
    return { success: false, message: '⛔ Identity Not Found. Please try the manual login option.' };
  }

  Logger.log('Authenticating user: ' + userEmail);

  try {
    if (isAuthorizedUser(userEmail)) {
      Logger.log('User is authorized. Setting session property.');

      // Store authorization in user properties (legacy support)
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('authorized', 'true');
      userProperties.setProperty('authorizedEmail', userEmail);
      userProperties.setProperty('authorizedTime', new Date().getTime().toString());

      // Generate Session Token (for client-side storage)
      const token = createSessionToken(userEmail);

      Logger.log('Session properties set. Returning redirect URL and Token.');

      return {
        success: true,
        redirectUrl: ScriptApp.getService().getUrl() + '?authorized=true',
        token: token,
        email: userEmail
      };
    } else {
      Logger.log('User is NOT authorized. Returning error.');
      return {
        success: false,
        message: '⛔ Access Denied: Your account is not authorized. Please contact the administrator.'
      };
    }
  } catch (error) {
    Logger.log('❌ Error during authentication: ' + error.message);
    Logger.log('Error stack: ' + error.stack);

    // Check if it's a spreadsheet configuration error
    if (error.message.includes('YOUR_SPREADSHEET_ID_HERE') ||
      error.message.includes('Invalid') ||
      error.message.includes('not found') ||
      error.message.includes('Spreadsheet')) {
      return {
        success: false,
        message: '⚠️ Configuration Error: SPREADSHEET_ID not configured in Config.js. Please update it with your actual spreadsheet ID from the spreadsheet URL.'
      };
    }

    // Generic error
    return {
      success: false,
      message: '❌ Authentication Error: ' + error.message + '. Please contact the administrator.'
    };
  }
}

// === SESSION MANAGEMENT ===

function createSessionToken(email) {
  // Generate a random token
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();

  // Store token -> email mapping for 6 hours (21600 seconds)
  cache.put(token, email, 21600);

  return token;
}

function validateSessionToken(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const email = cache.get(token);
  return email;
}

/**
 * Generates a 6-digit code, stores it in cache, and emails it to the user.
 * @param {string} email
 * @return {Object} result
 */
function sendVerificationCode(email) {
  try {
    if (!email) throw new Error("Email is required");
    email = email.trim();

    // 1. Check if user is authorized at all
    if (!isAuthorizedUser(email)) {
      return { success: false, message: 'This email is not authorized to access the system.' };
    }

    // 2. Generate 6-digit code
    const code = Math.floor(100000 + Math.random() * 900000).toString();

    // 3. Store in Cache (valid for 10 minutes)
    const cache = CacheService.getScriptCache();
    cache.put('otp_' + email, code, 600);

    // 4. Send Email
    MailApp.sendEmail({
      to: email,
      subject: 'Your Login Verification Code - myDoc App',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <h2 style="color: #2370b8;">Login Verification</h2>
          <p>Hello,</p>
          <p>Your verification code for myDoc Clinic App is:</p>
          <div style="background: #f4f4f4; padding: 15px; font-size: 24px; font-weight: bold; letter-spacing: 5px; text-align: center; border-radius: 8px; margin: 20px 0;">
            ${code}
          </div>
          <p>This code will expire in 10 minutes.</p>
          <p>If you did not request this code, please ignore this email.</p>
        </div>
      `
    });

    return { success: true, message: 'Verification code sent to ' + email };

  } catch (error) {
    Logger.log("Error sending code: " + error.toString());
    return { success: false, message: 'Failed to send code: ' + error.message };
  }
}

/**
 * Verifies the code entered by the user.
 * @param {string} email
 * @param {string} code
 * @return {Object} Auth result (same as authenticateUser)
 */
function verifyLoginCode(email, code) {
  try {
    if (!email || !code) throw new Error("Email and Code are required");
    email = email.trim();
    code = code.trim();

    // 1. Retrieve code from Cache
    const cache = CacheService.getScriptCache();
    const storedCode = cache.get('otp_' + email);

    // 2. Validate
    if (!storedCode) {
      return { success: false, message: 'Code expired or invalid. Please try again.' };
    }

    if (storedCode !== code) {
      return { success: false, message: 'Incorrect code.' };
    }

    // 3. Code matches! Proceed to authenticate
    cache.remove('otp_' + email); // Consume the code
    return authenticateUser(email);

  } catch (error) {
    Logger.log("Error verifying code: " + error.toString());
    return { success: false, message: 'Verification failed: ' + error.message };
  }
}

/**
 * Run this function ONCE manually in the script editor to authorize email permissions.
 */
function authorizeEmailService() {
  console.log("Checking email permissions...");
  // This line triggers the auth prompt without sending an email
  const quota = MailApp.getRemainingDailyQuota();
  console.log("Email service authorized! Remaining Daily Quota: " + quota);
}
