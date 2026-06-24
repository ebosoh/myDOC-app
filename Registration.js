/**
 * myDoc - Patient Registration Functions
 */

function registerPatient(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Registration_Records');

  // Generate Patient ID: KCL-YYYY-XXXXX
  const year = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  const idSuffix = (lastRow).toString().padStart(5, '0'); // Simple increment
  const patientId = `KCL-${year}-${idSuffix}`;

  const timestamp = new Date();
  const staffId = Session.getActiveUser().getEmail(); // Using email as Staff ID for now

  const rowData = [
    patientId,
    timestamp,
    data.national_id || '',
    data.NHIF_no || '',
    data.first_name,
    data.middle_name || '',
    data.last_name,
    data.date_of_birth,
    data.gender,
    data.marital_status || '',
    data.phone_number,
    data.email || '',
    data.county,
    data.sub_county,
    data.ward || '',
    data.language || '',
    data.next_of_kin,
    data.NOK_relationship,
    data.NOK_phone,
    data.blood_group || '',
    data.known_allergies || '',
    true, // is_active
    staffId,
    data.disability_status || '',
    data.village || ''
  ];

  sheet.appendRow(rowData);
  return { success: true, message: 'Patient registered successfully!', patientId: patientId };
}

function getRecentPatients() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Registration_Records');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Only headers

  // Get last 20 records
  const startRow = Math.max(2, lastRow - 19);
  const numRows = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  // Use SHEET_CONFIG for headers to ensure correct mapping even if sheet headers are missing/changed
  const headers = SHEET_CONFIG['Registration_Records'];

  // Map to object and reverse to show newest first
  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      let value = (index < row.length) ? row[index] : '';
      // Convert Date objects to string to ensure successful serialization
      if (value instanceof Date) {
        value = value.toISOString();
      }
      obj[header] = value;
    });
    return obj;
  }).reverse();
}

function searchPatient(query) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Registration_Records');

    if (!sheet) {
      Logger.log('ERROR: Registration_Records sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    Logger.log('Sheet has ' + lastRow + ' rows');

    if (lastRow < 2) {
      Logger.log('No data rows in sheet');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove headers

    Logger.log('Processing ' + data.length + ' data rows');

    // Simple search implementation
    const q = query.toString().toLowerCase().trim();
    if (!q) {
      Logger.log('Empty query');
      return [];
    }

    Logger.log('Searching for: ' + q);

    const results = data.filter(row => {
      if (!row || row.length < 1) return false;
      // Search by ID, Name, Phone, National ID
      const pid = String(row[0] || '').toLowerCase();
      const name = (String(row[4] || '') + ' ' + String(row[6] || '')).toLowerCase();
      const phone = String(row[10] || '').toLowerCase();
      const nid = String(row[2] || '').toLowerCase();

      return pid.includes(q) || name.includes(q) || phone.includes(q) || nid.includes(q);
    }).slice(0, 10); // Limit to 10 results

    Logger.log('Found ' + results.length + ' results');

    const mapped = results.map(row => ({
      patient_id: row[0],
      first_name: row[4],
      last_name: row[6],
      gender: row[8],
      date_of_birth: row[7],
      phone_number: row[10]
    }));

    Logger.log('Returning: ' + JSON.stringify(mapped));
    return mapped;
  } catch (error) {
    Logger.log('ERROR in searchPatient: ' + error.toString());
    return [];
  }
}

function getPatientsForTriage() {
  try {
    Logger.log('getPatientsForTriage called');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const triageSheet = ss.getSheetByName('Triage_Records');
    const recentPatients = getRecentPatients();

    if (!triageSheet) {
      Logger.log('Triage_Records sheet not found, returning all recent patients');
      return recentPatients;
    }

    // Get list of patient IDs already triaged
    const triageData = triageSheet.getDataRange().getValues();
    const triagedPatientIds = new Set();
    for (let i = 1; i < triageData.length; i++) {
      if (triageData[i][1]) { // patient_id is at index 1
        triagedPatientIds.add(String(triageData[i][1]).trim());
      }
    }

    // Filter out patients who are already in the triage records
    const filtered = recentPatients.filter(p => !triagedPatientIds.has(String(p.patient_id).trim()));

    Logger.log('Returning ' + filtered.length + ' patients for triage after filtering');
    return filtered;
  } catch (error) {
    Logger.log('ERROR in getPatientsForTriage: ' + error.toString());
    return [];
  }
}

function getPatientPhone(patientId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const regSheet = ss.getSheetByName('Registration_Records');

    if (!regSheet) return 'N/A';

    const data = regSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === patientId) {
        return data[i][10] || 'N/A'; // phone_number at index 10
      }
    }

    return 'N/A';
  } catch (error) {
    Logger.log('ERROR in getPatientPhone: ' + error.toString());
    return 'N/A';
  }
}

// ========================================
// PATIENT HISTORY VIEW FUNCTIONS
// ========================================

/**
 * Search patients by multiple criteria
 * @param {string} searchQuery - The search term
 * @param {string} searchType - Type: 'national_id', 'sha_no', 'phone', 'name'
 * @returns {Array} Array of matching patients with full demographics
 */
function searchPatients(searchQuery, searchType) {
  try {
    Logger.log('=== searchPatients START ===');
    Logger.log('Search query: ' + searchQuery + ', Type: ' + searchType);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Registration_Records');

    if (!sheet) {
      Logger.log('ERROR: Registration_Records sheet not found');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove headers

    const query = String(searchQuery || '').toLowerCase().trim();
    if (!query) {
      Logger.log('Empty search query');
      return [];
    }

    Logger.log('Searching ' + data.length + ' patient records');

    // Filter based on search type
    const results = data.filter(row => {
      if (!row || row.length < 1) return false;

      switch (searchType) {
        case 'national_id':
          return String(row[2] || '').toLowerCase().includes(query);
        case 'sha_no':
          return String(row[3] || '').toLowerCase().includes(query);
        case 'phone':
          return String(row[10] || '').toLowerCase().includes(query);
        case 'name':
          const fullName = (String(row[4] || '') + ' ' + String(row[5] || '') + ' ' + String(row[6] || '')).toLowerCase();
          return fullName.includes(query);
        default:
          // Search all fields if no type specified
          const pid = String(row[0] || '').toLowerCase();
          const nid = String(row[2] || '').toLowerCase();
          const sha = String(row[3] || '').toLowerCase();
          const name = (String(row[4] || '') + ' ' + String(row[5] || '') + ' ' + String(row[6] || '')).toLowerCase();
          const phone = String(row[10] || '').toLowerCase();
          return pid.includes(query) || nid.includes(query) || sha.includes(query) || name.includes(query) || phone.includes(query);
      }
    }).slice(0, 20); // Limit to 20 results

    Logger.log('Found ' + results.length + ' matching patients');

    // Map to simplified patient objects (for testing)
    const patients = results.map(row => ({
      patient_id: String(row[0] || ''),
      first_name: String(row[4] || ''),
      middle_name: String(row[5] || ''),
      last_name: String(row[6] || ''),
      gender: String(row[8] || ''),
      phone_number: String(row[10] || ''),
      national_id: String(row[2] || ''),
      sha_no: String(row[3] || '')
    }));

    Logger.log('=== searchPatients END ===');
    Logger.log('Returning ' + patients.length + ' patients');

    // Check each patient for serialization issues
    for (var i = 0; i < patients.length; i++) {
      Logger.log('=== Checking Patient ' + (i + 1) + ' ===');
      var patient = patients[i];

      // Check each field
      for (var key in patient) {
        var value = patient[key];
        var valueType = typeof value;

        if (value === null) {
          Logger.log('WARNING: ' + key + ' is NULL');
        } else if (value === undefined) {
          Logger.log('WARNING: ' + key + ' is UNDEFINED');
        } else if (valueType === 'object') {
          Logger.log('WARNING: ' + key + ' is OBJECT: ' + value);
        }
      }

      // Try to serialize this patient
      try {
        var json = JSON.stringify(patient);
        Logger.log('Patient ' + (i + 1) + ' serializes OK, length: ' + json.length);
      } catch (e) {
        Logger.log('ERROR: Patient ' + (i + 1) + ' FAILS serialization: ' + e.toString());
      }
    }

    // Try to serialize the entire array
    try {
      var allJson = JSON.stringify(patients);
      Logger.log('SUCCESS: All patients array serializes OK, length: ' + allJson.length);
    } catch (e) {
      Logger.log('ERROR: Patients array FAILS serialization: ' + e.toString());
    }

    // Wrap in object to ensure proper transmission to client
    // Google Apps Script sometimes fails to send plain arrays
    var response = {
      success: true,
      patients: patients,
      count: patients.length
    };

    Logger.log('Returning wrapped response with ' + response.count + ' patients');

    // TEST: Return hardcoded sample data
    return {
      success: true,
      patients: [
        { patient_id: 'TEST1', first_name: 'John', middle_name: '', last_name: 'Doe', gender: 'Male', phone_number: '123456789', national_id: '11111111', sha_no: '222' },
        { patient_id: 'TEST2', first_name: 'Jane', middle_name: 'Mary', last_name: 'Smith', gender: 'Female', phone_number: '987654321', national_id: '22222222', sha_no: '444' },
        { patient_id: 'TEST3', first_name: 'Bob', middle_name: '', last_name: 'Johnson', gender: 'Male', phone_number: '555555555', national_id: '33333333', sha_no: '666' }
      ],
      count: 3
    };

    // Real data (commented out for testing)
    // return response;

  } catch (error) {
    Logger.log('ERROR in searchPatients: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return [];
  }
}

/**
 * Get complete medical history for a patient
 * @param {string} patientId - The patient ID
 * @returns {Object} Complete patient history object
 */
function getPatientHistory(patientId) {
  try {
    Logger.log('=== getPatientHistory START ===');
    Logger.log('Patient ID: ' + patientId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Get patient demographics
    const regSheet = ss.getSheetByName('Registration_Records');
    const triageSheet = ss.getSheetByName('Triage_Records');
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    const billingSheet = ss.getSheetByName('Billing_Records');

    const history = {
      patient: null,
      visits: [],
      consultations: [],
      lab_results: [],
      billing: []
    };

    // Get patient demographics
    if (regSheet) {
      const regData = regSheet.getDataRange().getValues();
      for (let i = 1; i < regData.length; i++) {
        if (regData[i][0] === patientId) {
          const row = regData[i];
          history.patient = {
            patient_id: String(row[0] || ''),
            registration_date: row[1] instanceof Date ? row[1].toISOString() : String(row[1] || ''),
            national_id: String(row[2] || ''),
            sha_no: String(row[3] || ''),
            first_name: String(row[4] || ''),
            middle_name: String(row[5] || ''),
            last_name: String(row[6] || ''),
            date_of_birth: row[7] instanceof Date ? row[7].toISOString() : String(row[7] || ''),
            gender: String(row[8] || ''),
            marital_status: String(row[9] || ''),
            phone_number: String(row[10] || ''),
            email: String(row[11] || ''),
            county: String(row[12] || ''),
            sub_county: String(row[13] || ''),
            ward: String(row[14] || ''),
            blood_group: String(row[19] || ''),
            known_allergies: String(row[20] || '')
          };
          break;
        }
      }
    }

    if (!history.patient) {
      Logger.log('Patient not found');
      return { success: false, message: 'Patient not found' };
    }

    // Get triage/visit history
    if (triageSheet) {
      const triageData = triageSheet.getDataRange().getValues();
      triageData.shift(); // Remove header

      triageData.forEach(row => {
        if (row[1] === patientId) {
          history.visits.push({
            encounter_id: String(row[0] || ''),
            triage_date: row[2] instanceof Date ? row[2].toISOString() : String(row[2] || ''),
            chief_complaint: String(row[3] || ''),
            temperature: String(row[4] || ''),
            blood_pressure: String(row[5] || ''),
            pulse_rate: String(row[6] || ''),
            respiratory_rate: String(row[7] || ''),
            weight: String(row[8] || ''),
            height: String(row[9] || ''),
            priority: String(row[10] || '')
          });
        }
      });
    }

    // Get consultation/treatment history
    if (treatmentSheet) {
      const treatmentData = treatmentSheet.getDataRange().getValues();
      treatmentData.shift(); // Remove header

      treatmentData.forEach(row => {
        if (row[2] === patientId) {
          history.consultations.push({
            treatment_id: String(row[0] || ''),
            encounter_id: String(row[1] || ''),
            consultation_date: row[4] instanceof Date ? row[4].toISOString() : String(row[4] || ''),
            diagnosis_code: String(row[7] || ''),
            diagnosis_desc: String(row[8] || ''),
            prescribed_medications: String(row[10] || ''),
            treatment_plan: String(row[12] || ''),
            disposition: String(row[14] || ''),
            consultation_stage: String(row[15] || '')
          });
        }
      });
    }

    // Get lab results
    if (labSheet) {
      const labData = labSheet.getDataRange().getValues();
      labData.shift(); // Remove header

      labData.forEach(row => {
        if (row[1] === patientId) {
          history.lab_results.push({
            lab_order_id: String(row[0] || ''),
            encounter_id: String(row[2] || ''),
            order_date: row[4] instanceof Date ? row[4].toISOString() : String(row[4] || ''),
            test_name: String(row[5] || ''),
            status: String(row[7] || ''),
            result_value: String(row[8] || ''),
            result_units: String(row[9] || ''),
            reference_range: String(row[10] || ''),
            abnormal_flag: String(row[11] || ''),
            detailed_description: String(row[13] || '')
          });
        }
      });
    }

    // Get billing history
    if (billingSheet) {
      const billingData = billingSheet.getDataRange().getValues();
      billingData.shift(); // Remove header

      billingData.forEach(row => {
        if (row[2] === patientId) {
          history.billing.push({
            bill_id: String(row[0] || ''),
            encounter_id: String(row[1] || ''),
            billing_date: row[3] instanceof Date ? row[3].toISOString() : String(row[3] || ''),
            total_charge: String(row[4] || ''),
            amount_paid: String(row[5] || ''),
            payment_method: String(row[6] || ''),
            payment_date: row[7] instanceof Date ? row[7].toISOString() : String(row[7] || '')
          });
        }
      });
    }

    // Sort by date (newest first)
    history.visits.sort((a, b) => new Date(b.triage_date) - new Date(a.triage_date));
    history.consultations.sort((a, b) => new Date(b.consultation_date) - new Date(a.consultation_date));
    history.lab_results.sort((a, b) => new Date(b.order_date) - new Date(a.order_date));
    history.billing.sort((a, b) => new Date(b.billing_date) - new Date(a.billing_date));

    Logger.log('History retrieved - Visits: ' + history.visits.length + ', Consultations: ' + history.consultations.length);
    Logger.log('=== getPatientHistory END ===');

    return {
      success: true,
      history: history
    };

  } catch (error) {
    Logger.log('ERROR in getPatientHistory: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return { success: false, message: error.message };
  }
}

/**
 * Simple patient search - returns minimal data to avoid serialization issues
 */
function simpleSearchPatients(query) {
  try {
    Logger.log('=== simpleSearchPatients START ===');
    Logger.log('Query: ' + query);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Registration_Records');

    if (!sheet) {
      Logger.log('Sheet not found');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    data.shift(); // Remove header

    const q = String(query || '').toLowerCase().trim();
    if (!q) {
      Logger.log('Empty query');
      return [];
    }

    Logger.log('Searching ' + data.length + ' records');

    // Search in patient ID, name, phone, national ID
    const results = data.filter(function (row) {
      if (!row || row.length < 1) return false;

      const pid = String(row[0] || '').toLowerCase();
      const firstName = String(row[4] || '').toLowerCase();
      const lastName = String(row[6] || '').toLowerCase();
      const phone = String(row[10] || '').toLowerCase();
      const nationalId = String(row[2] || '').toLowerCase();

      return pid.includes(q) || firstName.includes(q) || lastName.includes(q) ||
        phone.includes(q) || nationalId.includes(q);
    }).slice(0, 10); // Limit to 10 results

    Logger.log('Found ' + results.length + ' results');

    // Return ONLY strings - no dates, no booleans, no objects
    const patients = results.map(function (row) {
      return {
        id: String(row[0] || ''),
        name: String(row[4] || '') + ' ' + String(row[6] || ''),
        phone: String(row[10] || ''),
        gender: String(row[8] || ''),
        national_id: String(row[2] || ''),
        sha_no: String(row[3] || '')
      };
    });

    Logger.log('Returning ' + patients.length + ' patients');
    return patients;

  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    return [];
  }
}

