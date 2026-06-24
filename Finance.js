/**
 * myDoc - Finance Module Functions
 */

function getServiceSummary(encounterId) {
  try {
    Logger.log('=== getServiceSummary for ' + encounterId + ' ===');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Get treatment record
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    if (!treatmentSheet) return { success: false, message: 'Treatment sheet not found' };

    const treatmentData = treatmentSheet.getDataRange().getValues();
    treatmentData.shift();

    const treatment = treatmentData.find(r => r[1] === encounterId);

    const services = [];

    // 1. Consultation & 2. Medications (from Treatment)
    if (treatment) {
      services.push({
        category: 'Consultation',
        items: ['Doctor Consultation']
      });

      // Medications
      const medicationsJson = treatment[10];
      const medications = [];
      if (medicationsJson) {
        try {
          const meds = JSON.parse(medicationsJson);
          if (Array.isArray(meds)) {
            meds.forEach(med => {
              medications.push((med.name || 'Unknown') + ' (' + (med.dose || 'as prescribed') + ')');
            });
          }
        } catch (e) {
          if (typeof medicationsJson === 'string') medications.push(medicationsJson);
        }
      }
      if (medications.length > 0) {
        services.push({ category: 'Drugs', items: medications });
      }
    }

    // 3. Lab Tests
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    const labTests = [];
    if (labSheet && labSheet.getLastRow() > 1) {
      const labData = labSheet.getDataRange().getValues();
      labData.shift();

      Logger.log('Checking lab orders for encounter: ' + encounterId);
      labData.forEach(lab => {
        if (lab[2] === encounterId) { // encounter_id is at index 2
          labTests.push(lab[5]); // test_name is at index 5
          Logger.log('Found lab test: ' + lab[5]);
        }
      });
    }

    if (labTests.length > 0) {
      services.push({
        category: 'Lab Tests',
        items: labTests
      });
    }

    // 4. Family Planning (FP)
    const fpSheet = ss.getSheetByName('FP_Tracking_Records');
    if (fpSheet && fpSheet.getLastRow() > 1) {
      const fpData = fpSheet.getDataRange().getValues();
      fpData.shift();
      // encounter_id at index 14, tracking_id at index 0
      const fpRecords = fpData.filter(r => String(r[14]) === String(encounterId) || String(r[0]) === String(encounterId));

      fpRecords.forEach(fp => {
        // fp_type(4), dosage(6), duration_days(7)
        const details = `Type: ${fp[4] || 'N/A'}, Dosage: ${fp[6] || 'N/A'}, Duration: ${fp[7] || '0'} days`;
        services.push({
          category: 'Family Planning',
          items: [details]
        });
      });
    }

    Logger.log('Total services found: ' + services.length);
    Logger.log('Services: ' + JSON.stringify(services));
    return {
      success: true,
      services: services
    };
  } catch (error) {
    Logger.log('ERROR in getServiceSummary: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function processPayment(data) {
  try {
    Logger.log('=== processPayment START ===');
    Logger.log('Received data: ' + JSON.stringify(data));

    // Validate required fields
    if (!data.encounter_id) {
      Logger.log('ERROR: Missing encounter_id');
      return {
        success: false,
        message: 'Missing encounter ID. Please try again.'
      };
    }

    if (!data.total_charge_amount || data.amount_paid_today === undefined || data.amount_paid_today === '') {
      return {
        success: false,
        message: 'Please fill in Total Charge and Amount Paid fields.'
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Billing_and_Payment_Records');

    // Generate Payment ID: PAY-YYYY-XXXXX
    const year = new Date().getFullYear();
    const lastRow = sheet.getLastRow();
    const idSuffix = (lastRow).toString().padStart(5, '0');
    const paymentId = `PAY-${year}-${idSuffix}`;

    const timestamp = new Date();
    const staffId = Session.getActiveUser().getEmail();

    const total = parseFloat(data.total_charge_amount) || 0;
    const paid = parseFloat(data.amount_paid_today) || 0;
    const balance = total - paid;

    // Parse charges breakdown to determine Service Category and detailed notes
    let categories = [];
    let breakdownDetails = '';

    if (data.charges_breakdown) {
      try {
        const bd = JSON.parse(data.charges_breakdown);
        if (bd.consultation && bd.consultation.length > 0) categories.push('Consultation');
        if (bd.lab_tests && bd.lab_tests.length > 0) categories.push('Lab');
        if (bd.drugs && bd.drugs.length > 0) categories.push('Pharmacy');

        // Construct a readable breakdown string for notes
        let parts = [];
        if (bd.consultation) bd.consultation.forEach(i => parts.push(`${i.description} (${i.amount})`));
        if (bd.lab_tests) bd.lab_tests.forEach(i => parts.push(`${i.description} (${i.amount})`));
        if (bd.drugs) bd.drugs.forEach(i => parts.push(`${i.description} (${i.amount})`));

        if (parts.length > 0) {
          breakdownDetails = 'Items: ' + parts.join(', ');
        }

      } catch (e) {
        Logger.log('Error parsing charges_breakdown: ' + e);
        breakdownDetails = 'Error parsing confirmation details.';
      }
    }

    // Determine finalized Service Category string
    // If categories were found, join them. Otherwise use the passed charge_details or default.
    const serviceCategory = categories.length > 0 ? categories.join(', ') : (data.charge_details || 'Consultation');

    // Combine user notes with system breakdown
    const finalNotes = (data.notes ? data.notes + ' | ' : '') + breakdownDetails;

    // Prepare itemized list for receipt
    let receiptItems = [];
    if (data.charges_breakdown) {
      try {
        const bd = JSON.parse(data.charges_breakdown);
        if (bd.consultation) bd.consultation.forEach(i => receiptItems.push({ category: 'Consultation', description: i.description, amount: i.amount }));
        if (bd.lab_tests) bd.lab_tests.forEach(i => receiptItems.push({ category: 'Lab Test', description: i.description, amount: i.amount }));
        if (bd.drugs) bd.drugs.forEach(i => receiptItems.push({ category: 'Drug', description: i.description, amount: i.amount }));
      } catch (e) {
        Logger.log('Error preparing receipt items: ' + e);
      }
    }

    const rowData = [
      paymentId,
      data.encounter_id,
      data.patient_id,
      timestamp,
      serviceCategory,
      total,
      paid,
      data.payment_method || 'Cash', // Combined string for split payments
      data.transaction_reference || '',
      balance,
      staffId,
      finalNotes
    ];

    Logger.log('Saving payment with encounter_id: ' + data.encounter_id);
    Logger.log('Row data: ' + JSON.stringify(rowData));

    sheet.appendRow(rowData);

    // If there's an outstanding balance, add/update debtor record
    if (balance > 0) {
      updateDebtorRecord({
        patient_id: data.patient_id,
        patient_name: data.patient_name,
        phone_number: data.phone_number,
        encounter_id: data.encounter_id,
        service_date: timestamp,
        service_description: serviceCategory,
        total_amount: total,
        amount_paid: paid,
        outstanding_balance: balance
      });
    } else {
      // If fully paid, mark debtor as cleared (if they were a debtor previously)
      clearDebtorRecord(data.encounter_id);

      // Also check if this payment clears a partial debt? 
      // The current logic seems to treat each payment as linked to an encounter.
      // Ideally, if it's a new full payment for an old encounter, we should handle that,
      // but 'processPayment' usually handles the initial bill context.
      // Partial payments for existing debts are handled by 'recordPartialPayment'.
    }

    Logger.log('Payment saved successfully');

    // Return complete receipt data
    const result = {
      success: true,
      message: 'Payment processed.',
      receipt: {
        receiptId: paymentId,
        patientId: data.patient_id,
        patientName: data.patient_name || 'N/A',
        encounterId: data.encounter_id,
        date: timestamp.toLocaleDateString('en-GB'),
        time: timestamp.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
        serviceCategory: serviceCategory,
        totalCharge: total,
        amountPaid: paid,
        balance: balance,
        paymentMethod: data.payment_method || 'Cash',
        reference: data.transaction_reference || 'N/A',
        items: receiptItems // New field for detailed breakdown
      }
    };

    Logger.log('=== processPayment END ===');
    return result;
  } catch (error) {
    Logger.log('ERROR in processPayment: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return {
      success: false,
      message: 'Payment processing failed: ' + error.toString()
    };
  }
}

function getPendingBills() {
  try {
    Logger.log('=== getPendingBills START ===');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const billingSheet = ss.getSheetByName('Billing_and_Payment_Records');
    const regSheet = ss.getSheetByName('Registration_Records');

    if (!treatmentSheet) {
      Logger.log('ERROR: Treatment_Records sheet not found');
      return [];
    }

    if (!regSheet) {
      Logger.log('ERROR: Registration_Records sheet not found');
      return [];
    }

    const treatmentData = treatmentSheet.getDataRange().getValues();
    treatmentData.shift();

    // Get billed encounters
    const billedEncounters = new Set();
    if (billingSheet && billingSheet.getLastRow() > 1) {
      const billingData = billingSheet.getDataRange().getValues();
      billingData.shift();
      billingData.forEach(r => {
        if (r[1]) billedEncounters.add(r[1]);
      });
    }

    // Get patient names
    const regData = regSheet.getDataRange().getValues();
    regData.shift();
    const patientMap = {};
    regData.forEach(r => {
      if (r[0]) {
        patientMap[r[0]] = (r[4] || '') + ' ' + (r[6] || '');
      }
    });

    // Filter unbilled treatments
    const unbilledTreatments = treatmentData.filter(r => {
      return r[1] && !billedEncounters.has(r[1]);
    });

    // Get recent 20
    const recentUnbilled = unbilledTreatments.slice(-20).reverse();

    const pendingBills = [...recentUnbilled];

    // Get unbilled FP Records
    const fpSheet = ss.getSheetByName('FP_Tracking_Records');
    if (fpSheet && fpSheet.getLastRow() > 1) {
      const fpData = fpSheet.getDataRange().getValues();
      fpData.shift();

      fpData.forEach(r => {
        const encId = String(r[14] || r[0]); // encounter_id or tracking_id
        // Add if not billed AND not already in list (from treatment)
        const alreadyListed = pendingBills.some(b => String(b[1]) === encId);

        if (!billedEncounters.has(encId) && !alreadyListed) {
          // Construct pseudo-treatment row for consistency
          // [?, encounter_id, patient_id, ?, date, ...]
          // FP indices: 2=Name, 1=ID, 5=Date
          // We need to match structure used below in map
          const pseudoRow = [
            '', // 0
            encId, // 1
            r[1], // 2 (patient_id)
            '', // 3
            r[5], // 4 (date)
            '', '', '', '', '', '', '', '', '',
            'FP Service' // 14 (disposition/type)
          ];
          pendingBills.push(pseudoRow);
        }
      });
    }

    // Sort combined list by date desc
    pendingBills.sort((a, b) => new Date(b[4]) - new Date(a[4]));
    const topBills = pendingBills.slice(0, 20);

    // Build result with services
    const result = topBills.map(r => {
      const encounterId = r[1];
      const serviceSummary = getServiceSummary(encounterId);

      return {
        encounter_id: encounterId,
        patient_id: r[2],
        patient_name: patientMap[r[2]] || 'Unknown (FP)', // Fallback if name not in map
        billing_date_time: r[4] instanceof Date ? r[4].toISOString() : r[4],
        disposition: r[14],
        services: serviceSummary.success ? serviceSummary.services : []
      };
    });

    Logger.log('Returning ' + result.length + ' bills with services');
    return result;

  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    return [];
  }
}

// Debtor management functions
function updateDebtorRecord(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let debtorSheet = ss.getSheetByName('Debtors_Records');

    // Create sheet if it doesn't exist
    if (!debtorSheet) {
      debtorSheet = ss.insertSheet('Debtors_Records');
      const headers = SHEET_CONFIG['Debtors_Records'];
      debtorSheet.appendRow(headers);
      debtorSheet.setFrozenRows(1);
      debtorSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');
    }

    // Check if debtor record already exists for this encounter
    const debtorData = debtorSheet.getDataRange().getValues();
    let existingRowIndex = -1;

    for (let i = 1; i < debtorData.length; i++) {
      if (debtorData[i][4] === data.encounter_id) { // encounter_id at index 4
        existingRowIndex = i + 1; // +1 for 1-indexed rows
        break;
      }
    }

    const year = new Date().getFullYear();

    if (existingRowIndex > 0) {
      // Update existing record
      debtorSheet.getRange(existingRowIndex, 9).setValue(data.amount_paid); // amount_paid
      debtorSheet.getRange(existingRowIndex, 10).setValue(data.outstanding_balance); // outstanding_balance
      debtorSheet.getRange(existingRowIndex, 11).setValue(data.service_date); // last_payment_date
      debtorSheet.getRange(existingRowIndex, 12).setValue(data.outstanding_balance > 0 ? 'PENDING' : 'CLEARED');
    } else {
      // Create new debtor record
      const debtorId = `DBT-${year}-${debtorSheet.getLastRow().toString().padStart(5, '0')}`;

      const rowData = [
        debtorId,
        data.patient_id,
        data.patient_name,
        data.phone_number || 'N/A',
        data.encounter_id,
        data.service_date,
        data.service_description,
        data.total_amount,
        data.amount_paid,
        data.outstanding_balance,
        data.service_date,
        'PENDING',
        ''
      ];

      debtorSheet.appendRow(rowData);
    }

    Logger.log('Debtor record updated for encounter: ' + data.encounter_id);
  } catch (error) {
    Logger.log('ERROR in updateDebtorRecord: ' + error.toString());
  }
}

function clearDebtorRecord(encounterId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const debtorSheet = ss.getSheetByName('Debtors_Records');

    if (!debtorSheet) return;

    const data = debtorSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === encounterId) { // encounter_id at index 4
        debtorSheet.getRange(i + 1, 12).setValue('CLEARED'); // status
        debtorSheet.getRange(i + 1, 10).setValue(0); // outstanding_balance
        Logger.log('Debtor record cleared for encounter: ' + encounterId);
        break;
      }
    }
  } catch (error) {
    Logger.log('ERROR in clearDebtorRecord: ' + error.toString());
  }
}

function getDebtors() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const debtorSheet = ss.getSheetByName('Debtors_Records');

    if (!debtorSheet || debtorSheet.getLastRow() < 2) {
      return [];
    }

    const data = debtorSheet.getDataRange().getValues();
    const headers = data.shift();

    // Filter for pending debtors only
    const debtors = data.filter(row => row[11] === 'PENDING' && row[9] > 0).map(row => ({
      debtor_id: row[0],
      patient_id: row[1],
      patient_name: row[2],
      phone_number: row[3],
      encounter_id: row[4],
      service_date: row[5] instanceof Date ? row[5].toISOString() : row[5],
      service_description: row[6],
      total_amount: row[7],
      amount_paid: row[8],
      outstanding_balance: row[9],
      last_payment_date: row[10] instanceof Date ? row[10].toISOString() : row[10],
      status: row[11]
    }));

    return debtors;
  } catch (error) {
    Logger.log('ERROR in getDebtors: ' + error.toString());
    return [];
  }
}

function recordPartialPayment(data) {
  try {
    Logger.log('=== recordPartialPayment START ===');
    Logger.log('Data received: ' + JSON.stringify(data));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const debtorSheet = ss.getSheetByName('Debtors_Records');
    const billingSheet = ss.getSheetByName('Billing_and_Payment_Records');

    if (!debtorSheet) {
      return { success: false, message: 'Debtors sheet not found' };
    }

    const newBalance = data.current_balance - data.payment_amount;
    const timestamp = new Date();

    // Update debtor record
    const debtorData = debtorSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < debtorData.length; i++) {
      if (debtorData[i][0] === data.debtor_id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      const currentPaid = parseFloat(debtorData[rowIndex - 1][8]) || 0;
      const newTotalPaid = currentPaid + data.payment_amount;

      debtorSheet.getRange(rowIndex, 9).setValue(newTotalPaid); // amount_paid
      debtorSheet.getRange(rowIndex, 10).setValue(newBalance); // outstanding_balance
      debtorSheet.getRange(rowIndex, 11).setValue(timestamp); // last_payment_date
      debtorSheet.getRange(rowIndex, 12).setValue(newBalance <= 0 ? 'CLEARED' : 'PENDING'); // status

      // Record in billing sheet
      const year = new Date().getFullYear();
      const paymentId = `PAY-${year}-${billingSheet.getLastRow().toString().padStart(5, '0')}`;
      const staffId = Session.getActiveUser().getEmail();

      const billingRow = [
        paymentId,
        data.encounter_id,
        data.patient_id,
        timestamp,
        'Partial Payment',
        0, // total_charge (already recorded)
        data.payment_amount,
        data.payment_method,
        data.transaction_reference,
        newBalance,
        staffId,
        'Partial payment towards debt'
      ];

      billingSheet.appendRow(billingRow);

      Logger.log('Partial payment recorded successfully');
      return { success: true, message: 'Partial payment recorded' };
    } else {
      return { success: false, message: 'Debtor record not found' };
    }
  } catch (error) {
    Logger.log('ERROR in recordPartialPayment: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}
