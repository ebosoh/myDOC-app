/**
 * myDoc - Treatment/Diagnosis Module Functions
 */

function getRecentTreatments() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Treatment_Records');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const startRow = Math.max(2, lastRow - 19);
  const numRows = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const headers = SHEET_CONFIG['Treatment_Records'];

  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      let value = (index < row.length) ? row[index] : '';
      if (value instanceof Date) value = value.toISOString();
      obj[header] = value;
    });
    return obj;
  }).reverse();
}

function saveTreatmentRecord(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Treatment_Records');

  // Generate Treatment ID: TXR-YYYY-XXXXX
  const year = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  const idSuffix = (lastRow).toString().padStart(5, '0');
  const treatmentId = `TXR-${year}-${idSuffix}`;

  const timestamp = new Date();
  const doctorId = Session.getActiveUser().getEmail();

  // Determine consultation stage automatically
  const hasLabOrders = data.labd_referral_ids && data.labd_referral_ids.trim() !== '';
  const hasPrescriptions = data.prescribed_medications && data.prescribed_medications.trim() !== '';

  // Logic:
  // - If lab orders exist → 'initial' (waiting for lab results), labResultsReviewed = false
  // - If no lab orders → 'final' (complete consultation), labResultsReviewed = true
  let consultationStage = 'final';
  let labResultsReviewed = true;
  if (hasLabOrders) {
    consultationStage = 'initial';
    labResultsReviewed = false;
  }







  const rowData = [
    treatmentId,
    data.encounter_id,
    data.patient_id,
    doctorId,
    timestamp,
    data.history_present_illness,
    data.physical_exam_findings || '',
    data.diagnosis_primary_code,
    data.diagnosis_primary_desc,
    data.diagnosis_secondary_code || '', // JSON string or comma sep
    data.prescribed_medications || '', // JSON string
    data.labd_referral_ids || '',
    data.treatment_plan || '',
    data.follow_up_date || '',
    data.disposition || '',
    consultationStage, // Column P
    labResultsReviewed, // Column Q
    labResultsReviewed ? timestamp : '', // Column R - lab_review_date
    timestamp // Column S - initial_consultation_date
  ];

  sheet.appendRow(rowData);

  // Create Lab Orders if any were selected
  if (data.labd_referral_ids && data.labd_referral_ids.trim() !== '') {
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    const labTests = data.labd_referral_ids.split(',');
    const doctorId = Session.getActiveUser().getEmail();

    labTests.forEach(testName => {
      if (testName.trim()) {
        // Generate Lab Order ID
        const labLastRow = labSheet.getLastRow();
        const labIdSuffix = (labLastRow).toString().padStart(5, '0');
        const labOrderId = `LAB-${year}-${labIdSuffix}`;

        const labRowData = [
          labOrderId,                    // lab_order_id
          data.patient_id,               // patient_id
          data.encounter_id,             // encounter_id - THIS IS CRITICAL!
          doctorId,                      // ordering_doctor_id
          timestamp,                     // order_date_time
          testName.trim(),               // test_name
          '',                            // test_code
          'Ordered',                     // test_status
          '',                            // result_value
          '',                            // result_units
          '',                            // reference_range
          '',                            // abnormal_flag
          '',                            // lab_technologist_id
          '',                            // detailed_description
          ''                             // file_url
        ];

        labSheet.appendRow(labRowData);
        Logger.log('Created lab order: ' + labOrderId + ' for test: ' + testName.trim());
      }
    });
  }

  return { success: true, message: 'Treatment record saved.', treatmentId: treatmentId };
}

// Get ICD-11 codes for autocomplete
function getICD11Codes() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ICD11_Codes');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('ICD11_Codes');
      var headers = ['code', 'description', 'category', 'common_name'];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');

      // Populate with sample codes
      populateICD11Codes(sheet);
    }

    if (sheet.getLastRow() < 2) {
      return [];
    }

    var data = sheet.getDataRange().getValues();
    var codes = [];

    for (var i = 1; i < data.length; i++) {
      codes.push({
        code: String(data[i][0]),
        description: String(data[i][1]),
        category: String(data[i][2]),
        common_name: String(data[i][3])
      });
    }

    // Sort alphabetically by description
    codes.sort(function (a, b) {
      return a.description.localeCompare(b.description);
    });

    Logger.log('Returning ' + codes.length + ' ICD-11 codes');
    return codes;

  } catch (error) {
    Logger.log('ERROR in getICD11Codes: ' + error.toString());
    return [];
  }
}

// NEW FUNCTIONS FOR TWO-QUEUE WORKFLOW

function getPendingLabResultsQueue() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    const regSheet = ss.getSheetByName('Registration_Records');

    if (!treatmentSheet || !labSheet || !regSheet) {
      return [];
    }

    const treatmentData = treatmentSheet.getDataRange().getValues();
    const labData = labSheet.getDataRange().getValues();
    const regData = regSheet.getDataRange().getValues();

    // Remove headers
    treatmentData.shift();
    labData.shift();
    regData.shift();

    // Create patient map
    const patientMap = {};
    regData.forEach(r => {
      if (r[0]) {
        patientMap[r[0]] = {
          name: (r[4] || '') + ' ' + (r[6] || ''),
          phone: r[10] || ''
        };
      }
    });

    // OPTIMIZATION: Index lab orders by encounter_id to avoid nested loop O(N*M)
    const labsByEncounter = {};
    labData.forEach(lab => {
      // Lab Structure: [lab_order_id, patient_id, encounter_id, ...]
      // encounter_id is at index 2
      const encId = lab[2];
      if (encId) {
        if (!labsByEncounter[encId]) {
          labsByEncounter[encId] = [];
        }
        labsByEncounter[encId].push(lab);
      }
    });

    // Find consultations with lab orders that have results
    const pendingQueue = [];

    Logger.log('Total treatment records: ' + treatmentData.length);
    Logger.log('Total lab records: ' + labData.length);

    treatmentData.forEach(treatment => {
      const encounterId = treatment[1];
      const labResultsReviewed = treatment[16]; // Column Q (index 16)

      // FAST LOOKUP: Get labs from map
      const encounterLabs = labsByEncounter[encounterId] || [];

      if (encounterLabs.length > 0) {
        // Check if any labs are resulted
        // Lab Status is at index 7
        const anyResulted = encounterLabs.some(lab => lab[7] === 'Resulted');

        // Only show if has lab results AND consultation not yet finalized
        if (anyResulted && labResultsReviewed === false) {
          const allResulted = encounterLabs.every(lab => lab[7] === 'Resulted');
          const patientId = treatment[2];
          const patientInfo = patientMap[patientId] || { name: 'Unknown', phone: '' };

          // Convert Date to string to avoid serialization issues
          const consultDate = treatment[4];
          const consultDateStr = consultDate instanceof Date ? consultDate.toISOString() : consultDate;

          pendingQueue.push({
            treatment_id: treatment[0],
            encounter_id: encounterId,
            patient_id: patientId,
            patient_name: patientInfo.name,
            consultation_date: consultDateStr,
            diagnosis_code: treatment[7],
            diagnosis_desc: treatment[8],
            lab_count: encounterLabs.length,
            resulted_count: encounterLabs.filter(lab => lab[7] === 'Resulted').length,
            all_resulted: allResulted
          });
        }
      }
    });

    Logger.log('Final queue size: ' + pendingQueue.length);

    // Sort by consultation date (newest first)
    pendingQueue.sort((a, b) => new Date(b.consultation_date) - new Date(a.consultation_date));

    return pendingQueue;

  } catch (error) {
    Logger.log('Error in getPendingLabResultsQueue: ' + error.toString());
    return [];
  }
}

function getLabResultsForConsultation(encounterId) {
  try {
    Logger.log('=== getLabResultsForConsultation START ===');
    Logger.log('Received encounterId: ' + encounterId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const labSheet = ss.getSheetByName('Laboratory_Orders');

    if (!labSheet) {
      Logger.log('ERROR: Lab sheet not found');
      return { success: false, message: 'Lab sheet not found', results: [] };
    }

    // OPTIMIZATION: Use TextFinder to find rows with this encounterId instead of reading the whole sheet
    // Encounter ID is in Column C (Index 3)
    // We search the entire Column C.
    // NOTE: TextFinder finds "cells". We need to map them to rows.

    // Check if encounterId is valid to avoid getting everything for empty string
    if (!encounterId || encounterId.trim() === '') {
      return { success: false, message: 'Invalid Encounter ID', results: [] };
    }

    // OPTIMIZATION: Read the data once and filter in memory.
    // TextFinder was causing issues with whitespace (exact match) or performance (partial match N+1 reads).
    // Since getPendingLabResultsQueue successfully reads the whole sheet, we can do it here too.

    // Get all data (headers are in row 1)
    const data = labSheet.getDataRange().getValues();
    data.shift(); // Remove headers

    // Filter matching encounter ID (index 2)
    // Using loose equality and trimming for maximum robustness
    const targetId = String(encounterId).trim();

    const results = data.filter(row => String(row[2]).trim() === targetId).map(lab => {
      // Map columns based on known index from previous code
      // logic was: lab[0] = id, lab[5]=name, lab[6]=code, lab[7]=status, lab[8]=result, ...
      return {
        lab_order_id: String(lab[0] || ''),
        test_name: String(lab[5] || ''),
        test_code: String(lab[6] || ''),
        status: String(lab[7] || ''),
        result_value: String(lab[8] || ''),
        result_units: String(lab[9] || ''),
        reference_range: String(lab[10] || ''),
        abnormal_flag: String(lab[11] || 'Normal'),
        detailed_description: String(lab[13] || ''),
        file_url: String(lab[14] || '')
      };
    });

    const response = {
      success: true,
      results: results,
      total_count: results.length,
      resulted_count: results.filter(r => r.status === 'Resulted').length
    };

    Logger.log('Found ' + results.length + ' results for encounter: ' + targetId);
    Logger.log('Returning response: ' + JSON.stringify(response));
    Logger.log('=== getLabResultsForConsultation END ===');
    return response;

  } catch (error) {
    Logger.log('ERROR in getLabResultsForConsultation: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return { success: false, message: error.message, results: [] };
  }
}

function finalizeConsultation(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Treatment_Records');
    const rows = sheet.getDataRange().getValues();

    // Find row by treatment_id
    const rowIndex = rows.findIndex(r => r[0] === data.treatment_id);
    if (rowIndex === -1) {
      return { success: false, message: 'Treatment record not found' };
    }

    // Update columns:
    // prescribed_medications (10), treatment_plan (12), follow_up_date (13), 
    // disposition (14), consultation_stage (15), lab_results_reviewed (16), lab_review_date (17)

    const now = new Date();

    sheet.getRange(rowIndex + 1, 8).setValue(data.final_diagnosis_code || ''); // Column H - diagnosis_primary_code
    sheet.getRange(rowIndex + 1, 9).setValue(data.final_diagnosis_desc || ''); // Column I - diagnosis_primary_desc
    sheet.getRange(rowIndex + 1, 11).setValue(data.prescribed_medications || ''); // Column K
    sheet.getRange(rowIndex + 1, 13).setValue(data.treatment_plan || ''); // Column M
    sheet.getRange(rowIndex + 1, 14).setValue(data.follow_up_date || ''); // Column N
    sheet.getRange(rowIndex + 1, 15).setValue(data.disposition || ''); // Column O
    sheet.getRange(rowIndex + 1, 16).setValue('final'); // Column P - consultation_stage
    sheet.getRange(rowIndex + 1, 17).setValue(true); // Column Q - lab_results_reviewed
    sheet.getRange(rowIndex + 1, 18).setValue(now); // Column R - lab_review_date

    return { success: true, message: 'Consultation finalized successfully' };

  } catch (error) {
    Logger.log('Error in finalizeConsultation: ' + error.toString());
    return { success: false, message: error.message };
  }
}

