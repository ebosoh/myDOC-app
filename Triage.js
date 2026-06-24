/**
 * myDoc - Triage Module Functions
 */

function saveTriageRecord(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Triage_Records');

  // Generate Encounter ID: TRI-YYYY-XXXXX
  const year = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  const idSuffix = (lastRow).toString().padStart(5, '0');
  const encounterId = `TRI-${year}-${idSuffix}`;

  const timestamp = new Date();
  const staffId = Session.getActiveUser().getEmail();

  const rowData = [
    encounterId,
    data.patient_id,
    timestamp,
    data.chief_complain,
    data.complain_duration || '',
    data.bp_systolic,
    data.bp_diastolic,
    data.heart_rate_bpm,
    data.resp_rate_breaths,
    data.temp_c,
    data.oxygen_saturation || '',
    data.weight,
    data.height,
    data.BMI,
    data.triage_priority,
    staffId,
    data.allergies || '',
    data.chronic_diseases || '',
    data.service_point || 'General'
  ];

  sheet.appendRow(rowData);

  // Create Google Calendar event with full patient details
  try {
    // Get patient details from Registration
    const regSheet = ss.getSheetByName('Registration_Records');

    if (regSheet) {
      const regData = regSheet.getDataRange().getValues();
      let patientName = 'Unknown Patient';
      let phoneNumber = 'N/A';

      // Find patient
      for (let i = 1; i < regData.length; i++) {
        if (regData[i][0] === data.patient_id) {
          const firstName = regData[i][4] || '';
          const middleName = regData[i][5] || '';
          const lastName = regData[i][6] || '';

          // Build full name (First + Middle + Last)
          patientName = `${firstName} ${middleName} ${lastName}`.replace(/\s+/g, ' ').trim();

          // Get phone number (column index 10)
          phoneNumber = regData[i][10] || 'N/A';

          break;
        }
      }

      // Use configured CALENDAR_ID
      const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

      if (calendar) {
        // Create event with full details
        const eventTitle = `📋 ${patientName}`;
        const eventDescription = `
👤 Patient: ${patientName}
📞 Phone: ${phoneNumber}
🆔 Patient ID: ${data.patient_id}
🩺 Chief Complaint: ${data.chief_complain}
📅 Triage Date: ${timestamp.toLocaleString('en-GB')}
🏥 Location: Makele Digital Clinic
Note: Patient has been triaged and is waiting for consultation.
        `.trim();

        // Set event time (30 minutes from triage time for consultation)
        const startTime = new Date(timestamp);
        const endTime = new Date(startTime.getTime() + (30 * 60 * 1000)); // 30 minutes later

        calendar.createEvent(eventTitle, startTime, endTime, {
          description: eventDescription,
          location: 'Makele Digital Clinic, Kitui'
        });

        Logger.log('Calendar event created for patient: ' + patientName + ' (' + phoneNumber + ')');
      } else {
        Logger.log('Calendar not found. Check CALENDAR_ID: ' + CALENDAR_ID);
      }
    }
  } catch (error) {
    Logger.log('Error creating calendar event: ' + error.toString());
    // Don't fail the triage if calendar creation fails
  }

  return { success: true, message: 'Triage record saved and calendar event created.', encounterId: encounterId };
}

function getPendingTriage() {
  try {
    Logger.log('getPendingTriage called');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const triageSheet = ss.getSheetByName('Triage_Records');
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const regSheet = ss.getSheetByName('Registration_Records');

    if (!triageSheet) {
      Logger.log('ERROR: Triage_Records sheet not found');
      return [];
    }

    if (!regSheet) {
      Logger.log('ERROR: Registration_Records sheet not found');
      return [];
    }

    const triageLastRow = triageSheet.getLastRow();
    Logger.log('Triage sheet has ' + triageLastRow + ' rows');

    if (triageLastRow < 2) {
      Logger.log('No triage records found');
      return [];
    }

    // OPTIMIZATION: Only get the last 1000 rows of triage records
    // Assuming patients triaged > 1000 records ago are no longer "waiting"
    const limit = 1000;
    const startRow = Math.max(2, triageLastRow - limit + 1);
    const numRows = triageLastRow - startRow + 1;

    // Get headers to map indexes correctly if needed, but we know the structure
    // Row 1 is headers. Data starts at startRow.

    const triageData = triageSheet.getRange(startRow, 1, numRows, triageSheet.getLastColumn()).getValues();
    Logger.log('Processing ' + triageData.length + ' triage records (Optimization: Last ' + limit + ' rows)');

    // Get list of encounter IDs already treated
    // OPTIMIZATION: Only fetch the first 2 columns of Treatment (ID and EncounterID) to save memory
    let treatedEncounters = new Set();
    if (treatmentSheet && treatmentSheet.getLastRow() > 1) {
      const tLastRow = treatmentSheet.getLastRow();
      // Column B is Encounter ID (index 2 in 1-based, index 1 in 0-based result of getValues)
      // fetch columns A and B (1 and 2)
      const treatmentData = treatmentSheet.getRange(2, 2, tLastRow - 1, 1).getValues();
      treatmentData.forEach(r => treatedEncounters.add(r[0]));
      Logger.log('Found ' + treatedEncounters.size + ' treated encounters');
    } else {
      Logger.log('No treatment records found');
    }

    // Get Patient Map for names
    // optimized to map ID -> Name
    const regData = regSheet.getDataRange().getValues();
    regData.shift();
    const patientMap = {};
    regData.forEach(r => {
      if (r[0]) {
        patientMap[r[0]] = (r[4] || '') + ' ' + (r[6] || '');
      }
    });

    // Filter Triage records not in Treatment
    const pending = triageData.filter(r => {
      if (!r || !r[0]) return false;
      // Filter out Family Planning patients (handled in FP module)
      if (r[18] === 'Family Planning') return false;
      return !treatedEncounters.has(r[0]);
    }).map(r => {
      const triageTime = r[2];
      return {
        encounter_id: r[0],
        patient_id: r[1],
        patient_name: patientMap[r[1]] || 'Unknown',
        triage_time: triageTime instanceof Date ? triageTime.toISOString() : triageTime,
        priority: r[14],
        chief_complain: r[3]
      };
    });

    // Sort by priority (Emergency > Urgent > ...) or Time? 
    // Usually FIFO but priority overrides. 
    // For now, let's keep array order (which is chronological by default due to read order)

    Logger.log('Returning ' + pending.length + ' pending triage records');

    return pending;
  } catch (error) {
    Logger.log('ERROR in getPendingTriage: ' + error.toString());
    return [];
  }
}
