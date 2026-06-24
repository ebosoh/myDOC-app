/**
 * myDoc - Web App Entry Points
 */

function doGet(e) {
  // ⚠️ TEMPORARY: Authentication disabled for debugging
  // TODO: Re-enable authentication once issue is resolved

  Logger.log('=== doGet called - AUTH DISABLED ===');
  Logger.log('User: ' + Session.getActiveUser().getEmail());

  // Skip all auth checks and go straight to main app
  return showMainApp();

  /* ORIGINAL AUTH CODE - COMMENTED OUT FOR NOW
  const userEmail = Session.getActiveUser().getEmail();
  const userProperties = PropertiesService.getUserProperties();
  const authorized = userProperties.getProperty('authorized');
  const authorizedEmail = userProperties.getProperty('authorizedEmail');

  Logger.log('=== doGet called ===');
  Logger.log('User email: ' + userEmail);
  Logger.log('Session authorized: ' + authorized);
  Logger.log('Session email: ' + authorizedEmail);

  // PRIORITY 1: Check session properties (fast, reliable, works even if SPREADSHEET_ID not configured)
  if (authorized === 'true' && authorizedEmail === userEmail) {
    Logger.log('✓ User authorized via session properties. Showing main app.');
    return showMainApp();
  }

  // PRIORITY 2: Check spreadsheet (requires SPREADSHEET_ID to be configured)
  try {
    Logger.log('Checking authorization via spreadsheet...');
    if (isAuthorizedUser(userEmail)) {
      Logger.log('✓ User authorized via spreadsheet check. Showing main app.');
      return showMainApp();
    }
  } catch (error) {
    Logger.log('⚠️ Spreadsheet authorization check failed: ' + error.message);
    Logger.log('⚠️ Make sure SPREADSHEET_ID is configured in Config.js');
    // Continue to show login page
  }

  Logger.log('✗ User not authorized. Showing login page.');
  return showLoginPage();
  */
}

function showMainApp() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('myDoc')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showLoginPage() {
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('Login - Makele Digital Clinic')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
