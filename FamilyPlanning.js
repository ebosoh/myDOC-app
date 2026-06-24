/**
 * myDoc - Family Planning Module Functions
 */

function registerFPClient(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('FP_Client_Records');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('FP_Client_Records');
      const headers = ['client_number', 'registration_date', 'client_name', 'client_type', 'first_ever_user', 'age', 'sex', 'disability_status', 'telephone', 'county', 'subcounty', 'village_estate', 'registered_by'];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');
    }

    // Generate Client Number: MDC-MMM-YYYY-###
    const now = new Date();
    const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const month = monthNames[now.getMonth()];
    const year = now.getFullYear();
    const prefix = 'MDC-' + month + '-' + year + '-';

    // Find max sequence for this month/year
    const allData = sheet.getDataRange().getValues();
    let maxSeq = 0;
    for (let i = 1; i < allData.length; i++) {
      const clientNum = allData[i][0];
      if (clientNum && clientNum.toString().indexOf(prefix) === 0) {
        const seq = parseInt(clientNum.split('-')[3]);
        if (seq > maxSeq) maxSeq = seq;
      }
    }

    let newSeq = (maxSeq + 1).toString();
    while (newSeq.length < 3) {
      newSeq = '0' + newSeq;
    }
    const clientNumber = prefix + newSeq;

    const timestamp = new Date();
    const staffId = Session.getActiveUser().getEmail();

    const rowData = [
      clientNumber,
      timestamp,
      data.client_name,
      data.client_type,
      data.first_ever_user,
      parseInt(data.age),
      data.sex,
      data.disability_status,
      data.telephone,
      data.county,
      data.subcounty || '',
      data.village_estate || '',
      staffId
    ];

    sheet.appendRow(rowData);

    Logger.log('FP Client registered: ' + clientNumber);
    return { success: true, message: 'Client registered successfully', clientNumber: clientNumber };
  } catch (error) {
    Logger.log('ERROR in registerFPClient: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function saveFPDispense(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trackingSheet = ss.getSheetByName('FP_Tracking_Records');
    let clientSheet = ss.getSheetByName('FP_Client_Records');

    // 1. Ensure FP_Tracking_Records exists
    if (!trackingSheet) {
      trackingSheet = ss.insertSheet('FP_Tracking_Records');
      const headers = ['tracking_id', 'client_number', 'client_name', 'phone_number', 'fp_type', 'dispense_date', 'dosage', 'duration_days', 'return_date', 'notes', 'status', 'dispensed_by', 'client_type', 'first_ever_user', 'encounter_id'];
      trackingSheet.appendRow(headers);
      trackingSheet.setFrozenRows(1);
      trackingSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');
    }

    // 2. Ensure FP_Client_Records exists (Statistics Source)
    if (!clientSheet) {
      clientSheet = ss.insertSheet('FP_Client_Records');
      const headers = ['client_number', 'registration_date', 'client_name', 'client_type', 'first_ever_user', 'age', 'sex', 'disability_status', 'telephone', 'county', 'subcounty', 'village_estate', 'registered_by'];
      clientSheet.appendRow(headers);
      clientSheet.setFrozenRows(1);
      clientSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');
    }

    // 3. Check if Client Exists in FP_Client_Records
    const clientData = clientSheet.getDataRange().getValues();
    let clientExists = false;
    for (let i = 1; i < clientData.length; i++) {
      if (clientData[i][0] == data.client_number) {
        clientExists = true;
        break;
      }
    }

    // 4. Auto-Register if Missing
    if (!clientExists) {
      Logger.log('Auto-registering FP Client: ' + data.client_number);
      const regSheet = ss.getSheetByName('Registration_Records');
      if (regSheet) {
        const regData = regSheet.getDataRange().getValues();
        let patientRow = null;
        for (let i = 1; i < regData.length; i++) {
          if (regData[i][0] == data.client_number) {
            patientRow = regData[i];
            break;
          }
        }

        if (patientRow) {
          // Extract & Map Data
          // Registration_Records indices: 4=First, 6=Last, 7=DOB, 8=Sex, 10=Phone, 31=County, 32=Sub, 42=Disability, 43=Village
          const dob = new Date(patientRow[7]);
          const age = new Date().getFullYear() - dob.getFullYear();
          let disabilityCode = '0'; // Default None
          const disText = String(patientRow[42] || '').toLowerCase(); // Column 42 is disability_status in Registration

          if (disText.includes('visual')) disabilityCode = '1';
          else if (disText.includes('hearing')) disabilityCode = '2';
          else if (disText.includes('cognitive')) disabilityCode = '3';
          else if (disText.includes('physical') || disText.includes('other')) disabilityCode = '4';
          // 'None' stays '0'

          const newClientRow = [
            data.client_number,
            new Date(),
            data.client_name,
            data.client_type,         // From Dispense Form
            data.first_ever_user,     // From Dispense Form
            age,
            patientRow[8],            // Sex
            disabilityCode,           // Mapped Disability
            patientRow[10],           // Phone
            patientRow[31],           // County
            patientRow[32],           // SubCounty
            patientRow[43],           // Village
            Session.getActiveUser().getEmail()
          ];
          clientSheet.appendRow(newClientRow);
        } else {
          Logger.log('Warning: Patient not found in Registration_Records for auto-registration.');
        }
      }
    }

    // 5. Generate Tracking ID: FPT-YYYY-#####
    const year = new Date().getFullYear();
    const lastRow = trackingSheet.getLastRow();
    let idSuffix = lastRow.toString();
    while (idSuffix.length < 5) {
      idSuffix = '0' + idSuffix;
    }
    const trackingId = 'FPT-' + year + '-' + idSuffix;

    // Calculate return date
    const dispenseDate = new Date(data.dispense_date);
    const durationDays = parseInt(data.duration_days);
    const returnDate = new Date(dispenseDate);
    returnDate.setDate(returnDate.getDate() + durationDays);

    const staffId = Session.getActiveUser().getEmail();

    const rowData = [
      trackingId,
      data.client_number,
      data.client_name,
      data.phone_number,
      data.fp_type,
      dispenseDate,
      data.dosage || '',
      durationDays,
      returnDate,
      data.notes || '',
      'Active',
      staffId,
      data.client_type || '',
      data.first_ever_user || '',
      data.encounter_id || ''
    ];

    trackingSheet.appendRow(rowData);

    Logger.log('FP Dispense saved: ' + trackingId);
    return { success: true, message: 'FP service recorded successfully (and client registered if new)', trackingId: trackingId };
  } catch (error) {
    Logger.log('ERROR in saveFPDispense: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getFPWaitingList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const triageSheet = ss.getSheetByName('Triage_Records');
    const fpSheet = ss.getSheetByName('FP_Tracking_Records');

    if (!triageSheet) return [];

    // Get Triage Data
    const triageData = triageSheet.getDataRange().getValues();
    // Headers: encounter_id(0), patient_id(1), timestamp(2), ..., triage_priority(14), ..., service_point(18)

    // Get Processed Encounters (from FP Tracking)
    const processedEncounters = new Set();
    if (fpSheet && fpSheet.getLastRow() > 1) {
      const fpData = fpSheet.getDataRange().getValues();
      // encounter_id is at index 14 (15th column) based on my update above
      for (let i = 1; i < fpData.length; i++) {
        if (fpData[i][14]) processedEncounters.add(String(fpData[i][14]));
      }
    }

    const waitingList = [];
    // Start from 1 to skip header
    for (let i = 1; i < triageData.length; i++) {
      const row = triageData[i];
      const encounterId = String(row[0]);
      const servicePoint = row[18]; // Check index

      // Filter: Service Point is FP AND Not Processed
      const servicePointStr = String(servicePoint).trim();
      if ((servicePointStr === 'Family Planning' || servicePointStr === 'Family Planning (FP)') && !processedEncounters.has(encounterId)) {
        // Need patient name. Triage has ID. Need to join with Registration?
        // Triage row[1] is patient_id.
        // Ideally we fetch names.
        waitingList.push({
          encounter_id: encounterId,
          patient_id: row[1],
          // Convert Date to string for safe serialization
          timestamp: (row[2] instanceof Date) ? row[2].toISOString() : String(row[2]),
          triage_priority: String(row[14] || ''),
          triage_note: String(row[3] || '') + ' (' + String(row[4] || '') + ')' // Chief complain + duration
        });
      }
    }

    // Fetch Names for waiting list
    if (waitingList.length > 0) {
      const regSheet = ss.getSheetByName('Registration_Records');
      if (regSheet) {
        const regData = regSheet.getDataRange().getValues();
        const patientMap = {};
        const phoneMap = {};

        // ID at 0
        // First Name at 4, Last Name at 6
        // Phone at 10
        for (let i = 1; i < regData.length; i++) {
          const id = regData[i][0];
          const fullName = regData[i][4] + ' ' + regData[i][6];
          patientMap[id] = fullName;
          phoneMap[id] = regData[i][10];
        }

        waitingList.forEach(item => {
          item.patient_name = patientMap[item.patient_id] || 'Unknown';
          item.phone_number = phoneMap[item.patient_id] || '';
        });
      }
    }

    Logger.log('FP Waiting list count: ' + waitingList.length);
    return waitingList ? waitingList.reverse() : [];

  } catch (e) {
    Logger.log('Error in getFPWaitingList: ' + e.toString());
    return [];
  }
}

function getFPTracking(limit, offset) {
  try {
    limit = limit || 10;
    offset = offset || 0;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('FP_Tracking_Records');

    if (!sheet || sheet.getLastRow() < 2) {
      return { records: [], hasMore: false };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var allRecords = [];

    var now = new Date();
    now.setHours(0, 0, 0, 0);

    for (var i = 1; i < data.length; i++) {
      if (data[i][10] === 'Active') {
        var record = {
          tracking_id: String(data[i][0]),
          client_number: String(data[i][1]),
          client_name: String(data[i][2]),
          phone_number: String(data[i][3]),
          fp_type: String(data[i][4]),
          dispense_date: data[i][5].toString(),
          dosage: String(data[i][6]),
          duration_days: Number(data[i][7]),
          return_date: data[i][8].toString(),
          notes: String(data[i][9]),
          status: String(data[i][10]),
          dispensed_by: String(data[i][11])
        };

        // Calculate days to return
        var returnDate = new Date(data[i][8]);
        returnDate.setHours(0, 0, 0, 0);
        var diffTime = returnDate.getTime() - now.getTime();
        var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        record.days_to_return = diffDays;

        allRecords.push(record);
      }
    }

    // Sort by return_date ascending
    allRecords.sort(function (a, b) {
      return new Date(a.return_date).getTime() - new Date(b.return_date).getTime();
    });

    // Apply pagination
    var paginatedRecords = allRecords.slice(offset, offset + limit);
    var hasMore = (offset + limit) < allRecords.length;

    return {
      records: paginatedRecords,
      hasMore: hasMore,
      total: allRecords.length
    };

  } catch (error) {
    Logger.log('ERROR in getFPTracking: ' + error.toString());
    return { records: [], hasMore: false, error: error.toString() };
  }
}

function getFPClients() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('FP_Client_Records');

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clients = [];

    for (let i = 1; i < data.length; i++) {
      const client = {};
      for (let j = 0; j < headers.length; j++) {
        client[headers[j]] = data[i][j];
      }
      clients.push(client);
    }

    // Sort by registration_date descending (newest first)
    clients.sort(function (a, b) {
      return new Date(b.registration_date) - new Date(a.registration_date);
    });

    return clients;
  } catch (error) {
    Logger.log('ERROR in getFPClients: ' + error.toString());
    return [];
  }
}

function getFPClientByNumber(clientNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('FP_Client_Records');

    if (!sheet) {
      return { success: false, message: 'FP Client Records not found' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === clientNumber) {
        const client = {};
        headers.forEach((header, index) => {
          client[header] = data[i][index];
        });
        return { success: true, client: client };
      }
    }

    return { success: false, message: 'Client not found' };
  } catch (error) {
    Logger.log('ERROR in getFPClientByNumber: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function archiveFPClient(trackingId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const trackingSheet = ss.getSheetByName('FP_Tracking_Records');
    let archiveSheet = ss.getSheetByName('FP_Archives');

    // Create archive sheet if it doesn't exist
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('FP_Archives');
      const headers = SHEET_CONFIG['FP_Archives'];
      archiveSheet.appendRow(headers);
      archiveSheet.setFrozenRows(1);
      archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#666').setFontColor('#ffffff');
    }

    if (!trackingSheet) {
      return { success: false, message: 'Tracking sheet not found' };
    }

    const data = trackingSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === trackingId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Tracking record not found' };
    }

    // Copy to archive with additional fields
    const recordData = data[rowIndex - 1];
    const archiveData = [...recordData, new Date(), Session.getActiveUser().getEmail()];
    archiveSheet.appendRow(archiveData);

    // Update status to Archived
    trackingSheet.getRange(rowIndex, 11).setValue('Archived');

    Logger.log('FP record archived: ' + trackingId);
    return { success: true, message: 'Client archived successfully' };
  } catch (error) {
    Logger.log('ERROR in archiveFPClient: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getFPStatistics() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('FP_Client_Records');

    if (!sheet || sheet.getLastRow() < 2) {
      return {
        clientType: { new: 0, revisit: 0 },
        firstEverUser: { yes: 0, no: 0 },
        ageDistribution: { under20: 0, age20_29: 0, age30_39: 0, age40plus: 0 },
        sex: { male: 0, female: 0, intersex: 0 },
        disability: { none: 0, visual: 0, hearing: 0, cognitive: 0, others: 0 }
      };
    }

    const data = sheet.getDataRange().getValues();
    const stats = {
      clientType: { new: 0, revisit: 0 },
      firstEverUser: { yes: 0, no: 0 },
      ageDistribution: { under20: 0, age20_29: 0, age30_39: 0, age40plus: 0 },
      sex: { male: 0, female: 0, intersex: 0 },
      disability: { none: 0, visual: 0, hearing: 0, cognitive: 0, others: 0 }
    };

    for (let i = 1; i < data.length; i++) {
      // Client Type
      if (data[i][3] === '1') stats.clientType.new++;
      else if (data[i][3] === '2') stats.clientType.revisit++;

      // First Ever User
      if (data[i][4] === 'Y') stats.firstEverUser.yes++;
      else if (data[i][4] === 'N') stats.firstEverUser.no++;

      // Age Distribution
      const age = parseInt(data[i][5]);
      if (age < 20) stats.ageDistribution.under20++;
      else if (age >= 20 && age < 30) stats.ageDistribution.age20_29++;
      else if (age >= 30 && age < 40) stats.ageDistribution.age30_39++;
      else stats.ageDistribution.age40plus++;

      // Sex
      if (data[i][6] === 'M') stats.sex.male++;
      else if (data[i][6] === 'F') stats.sex.female++;
      else if (data[i][6] === 'I') stats.sex.intersex++;

      // Disability
      const disability = data[i][7];
      if (disability === '0') stats.disability.none++;
      else if (disability === '1') stats.disability.visual++;
      else if (disability === '2') stats.disability.hearing++;
      else if (disability === '3') stats.disability.cognitive++;
      else if (disability === '4') stats.disability.others++;
    }

    return stats;
  } catch (error) {
    Logger.log('ERROR in getFPStatistics: ' + error.toString());
    return null;
  }
}
