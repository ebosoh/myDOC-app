/**
 * RemoteBridge.gs
 * 
 * COPY THIS CONTENT INTO A NEW FILE IN YOUR GOOGLE APPS SCRIPT EDITOR.
 * NAMING: Name the file "RemoteBridge.gs"
 * 
 * This script handles incoming HTTP POST requests from your GitHub-hosted frontend.
 */

function doPost(e) {
    // enable CORS
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    try {
        const requestData = JSON.parse(e.postData.contents);
        const functionName = requestData.functionName;
        const args = requestData.args || [];
        const userEmail = requestData.userEmail;

        // Security Check: Ideally, validate a token here. 
        // For now, we trust the email passed, but this isn't secure for production 
        // without a shared secret or proper OAuth flow.
        // In strict mode, we'd check requestData.authToken against the cache.

        // Whitelist allowed functions for security (optional but recommended)
        // const ALLOWED_FUNCTIONS = ['authenticateUser', 'registerPatient', 'searchPatient', 'getPatientHistory', ...];
        // if (!ALLOWED_FUNCTIONS.includes(functionName)) throw new Error("Function not authorized");

        // Execute the requested function
        if (typeof this[functionName] === 'function') {

            // MOCK SESSION: Since we are technically 'anonymous' in a web app,
            // we must override how functions get the current user.
            // We will attach the userEmail to a global property that other functions can read if needed,
            // OR we update those functions to accept email as a param.
            // For this bridge, we'll try to execute directly. 
            // NOTE: Standard `Session.getActiveUser()` will return BLANK or the OWNER's email depending on settings.
            // We need to refactor core functions to rely on the passed email.

            // Temporary Hack: Set a global variable for this execution context
            // (Note: This defaults to unsafe in concurrent executions, but GAS isolates executions usually)
            CacheService.getScriptCache().put('current_execution_user_' + Session.getTemporaryActiveUserKey(), userEmail, 60);

            const result = this[functionName].apply(this, args);

            return output.setContent(JSON.stringify({
                status: 'success',
                result: result
            }));
        } else {
            throw new Error(`Function '${functionName}' not found`);
        }

    } catch (error) {
        return output.setContent(JSON.stringify({
            status: 'error',
            error: error.message
        }));
    }
}

/**
 * Handle OPTIONS requests for CORS (if browsers preflight)
 * Note: GAS Web Apps often struggle with true OPTIONS handling, 
 * but `doGet` can sometimes act as a fallback or health check.
 */
function doGet(e) {
    const output = ContentService.createTextOutput(JSON.stringify({ status: 'running', message: 'Backend v1' }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}
