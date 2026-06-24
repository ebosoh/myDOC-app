/**
 * myDoc - Dashboard Functions
 * Provides real-time statistics and queue monitoring for clinic management
 */

function getTodayStatistics() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    // Get sheets
    const regSheet = ss.getSheetByName('Registration_Records');
    const triageSheet = ss.getSheetByName('Triage_Records');
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const financeSheet = ss.getSheetByName('Finance_Records');

    // Initialize counters
    let patientsRegistered = 0;
    let patientsTriaged = 0;
    let consultationsCompleted = 0;
    let billsPaid = 0;
    let drugRevenue = 0;
    let labRevenue = 0;
    let consultRevenue = 0;


    // Count patients registered today
    if (regSheet && regSheet.getLastRow() > 1) {
      const regData = regSheet.getDataRange().getValues();
      const regHeaders = regData[0];
      const regDateIdx = regHeaders.indexOf('registration_date');

      for (let i = 1; i < regData.length; i++) {
        const regDate = new Date(regData[i][regDateIdx]);
        if (regDate >= today && regDate < tomorrow) {
          patientsRegistered++;
        }
      }
    }
    // Count patients triaged today
    if (triageSheet && triageSheet.getLastRow() > 1) {
      const triageData = triageSheet.getDataRange().getValues();
      const triageHeaders = triageData[0];
      const triageDateIdx = triageHeaders.indexOf('triage_date_time');

      for (let i = 1; i < triageData.length; i++) {
        const triageDate = new Date(triageData[i][triageDateIdx]);
        if (triageDate >= today && triageDate < tomorrow) {
          patientsTriaged++;
        }
      }
    }
    // Count consultations completed today
    if (treatmentSheet && treatmentSheet.getLastRow() > 1) {
      const treatmentData = treatmentSheet.getDataRange().getValues();
      const treatmentHeaders = treatmentData[0];
      const consultDateIdx = treatmentHeaders.indexOf('initial_consultation_date');

      for (let i = 1; i < treatmentData.length; i++) {
        const consultDate = new Date(treatmentData[i][consultDateIdx]);
        if (consultDate >= today && consultDate < tomorrow) {
          consultationsCompleted++;
        }
      }
    }









    // Count bills paid and calculate revenue today
    const billingSheet = ss.getSheetByName('Billing_and_Payment_Records');
    if (billingSheet && billingSheet.getLastRow() > 1) {
      const billingData = billingSheet.getDataRange().getValues();
      const headers = billingData[0];

      // Find column indices
      const paymentDateIdx = headers.indexOf('billing_date_time');
      const notesIdx = headers.indexOf('notes');

      for (let i = 1; i < billingData.length; i++) {
        const row = billingData[i];
        const paymentDate = new Date(row[paymentDateIdx]);

        if (paymentDate >= today && paymentDate < tomorrow) {
          billsPaid++;

          // Parse notes to extract charges
          const notesStr = row[notesIdx];
          if (notesStr) {
            try {
              // Extract all amounts in parentheses
              const matches = notesStr.match(/\((\d+)\)/g);
              if (matches) {
                matches.forEach(match => {
                  const amount = parseFloat(match.replace(/[()]/g, ''));

                  // Get item name before the amount
                  const itemStart = notesStr.lastIndexOf(',', notesStr.indexOf(match));
                  const itemEnd = notesStr.indexOf(match);
                  const itemName = notesStr.substring(itemStart + 1, itemEnd).trim().toLowerCase();

                  // Simple categorization
                  if (itemName.includes('consultation')) {
                    consultRevenue += amount;
                  } else if (itemName.includes('test')) {
                    labRevenue += amount;
                  } else {
                    drugRevenue += amount;
                  }
                });
              }
            } catch (e) {
              Logger.log('Error parsing notes for row ' + i + ': ' + e.toString());
            }
          }
        }
      }
    }







    const totalRevenue = drugRevenue + labRevenue + consultRevenue;

    return {
      patients_registered: patientsRegistered,
      patients_triaged: patientsTriaged,
      consultations_completed: consultationsCompleted,
      bills_paid: billsPaid,
      collections: {
        drugs: Math.round(drugRevenue),
        lab_tests: Math.round(labRevenue),
        consultations: Math.round(consultRevenue),
        total: Math.round(totalRevenue)
      }
    };

  } catch (error) {
    Logger.log('Error in getTodayStatistics: ' + error.toString());
    return {
      patients_registered: 0,
      patients_triaged: 0,
      consultations_completed: 0,
      bills_paid: 0,
      collections: {
        drugs: 0,
        lab_tests: 0,
        consultations: 0,
        total: 0
      }
    };
  }
}

function getLiveQueueCounts() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Initialize counters
    let triageQueue = 0;
    let pendingLabResults = 0;
    let pendingBilling = 0;
    let pendingFPService = 0;
    let waitingConsultation = 0;
    let pendingLabOrders = 0;

    // 1. Triage Queue - Patients registered but not yet triaged
    const regSheet = ss.getSheetByName('Registration_Records');
    const triageSheet = ss.getSheetByName('Triage_Records');

    if (regSheet && regSheet.getLastRow() > 1) {
      const regData = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 1).getValues();
      const patientIds = regData.map(row => row[0]).filter(id => id);

      if (triageSheet && triageSheet.getLastRow() > 1) {
        const triageData = triageSheet.getRange(2, 2, triageSheet.getLastRow() - 1, 1).getValues();
        const triagedIds = triageData.map(row => row[0]).filter(id => id);

        // Count patients not yet triaged
        triageQueue = patientIds.filter(id => !triagedIds.includes(id)).length;
      } else {
        triageQueue = patientIds.length;
      }
    }

    // 2. Waiting Consultation - Patients triaged but not yet consulted (excluding FP)
    if (triageSheet && triageSheet.getLastRow() > 1) {
      const triageData = triageSheet.getDataRange().getValues();
      const headers = triageData.shift(); // Remove headers

      // Find service_point column index
      const servicePointIdx = headers.indexOf('service_point');

      // Get encounter IDs, excluding Family Planning patients
      const encounterIds = triageData
        .filter(row => {
          // Exclude Family Planning patients (they go to FP queue)
          if (servicePointIdx !== -1 && row[servicePointIdx] === 'Family Planning') {
            return false;
          }
          return row[0]; // Has encounter_id
        })
        .map(row => row[0]);

      const treatmentSheet = ss.getSheetByName('Treatment_Records');
      if (treatmentSheet && treatmentSheet.getLastRow() > 1) {
        const treatmentData = treatmentSheet.getRange(2, 2, treatmentSheet.getLastRow() - 1, 1).getValues();
        const consultedEncounters = treatmentData.map(row => row[0]).filter(id => id);

        waitingConsultation = encounterIds.filter(id => !consultedEncounters.includes(id)).length;
      } else {
        waitingConsultation = encounterIds.length;
      }
    }






    // 3. Pending Lab Results - Use existing function
    const labResultsQueue = getPendingLabResultsQueue();
    pendingLabResults = labResultsQueue.length;

    // 4. Pending Lab Orders - Lab orders not yet resulted
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    if (labSheet && labSheet.getLastRow() > 1) {
      const labData = labSheet.getRange(2, 1, labSheet.getLastRow() - 1, 8).getValues();
      labData.forEach(row => {
        const status = row[7]; // test_status is column H (index 7)
        if (status === 'Ordered') {
          pendingLabOrders++;
        }
      });
    }

    // 5. Pending Billing - Consultations completed but not yet billed
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    const financeSheet = ss.getSheetByName('Finance_Records');

    if (treatmentSheet && treatmentSheet.getLastRow() > 1) {
      const treatmentData = treatmentSheet.getRange(2, 1, treatmentSheet.getLastRow() - 1, 16).getValues();

      // Get all billed encounter IDs
      let billedEncounters = [];
      if (financeSheet && financeSheet.getLastRow() > 1) {
        const financeData = financeSheet.getRange(2, 2, financeSheet.getLastRow() - 1, 1).getValues();
        billedEncounters = financeData.map(row => row[0]).filter(id => id);
      }

      // Count finalized consultations not yet billed
      treatmentData.forEach(row => {
        const encounterId = row[1]; // encounter_id is column B (index 1)
        const consultationStage = row[15]; // consultation_stage is column P (index 15)

        if (consultationStage === 'final' && !billedEncounters.includes(encounterId)) {
          pendingBilling++;
        }
      });
    }


    // 6. Pending FP Service - Patients triaged for FP but not yet dispensed
    try {
      if (triageSheet && triageSheet.getLastRow() > 1) {
        const triageData = triageSheet.getDataRange().getValues();
        const headers = triageData[0];

        // Find column indices - safely check if they exist
        const encounterIdIdx = headers.indexOf('encounter_id');
        const servicePointIdx = headers.indexOf('service_point');

        // Only proceed if BOTH columns exist
        if (encounterIdIdx !== -1 && servicePointIdx !== -1) {
          // Get all FP dispense encounter IDs
          const fpSheet = ss.getSheetByName('FP_Dispense_Records');
          let dispensedEncounters = [];

          if (fpSheet && fpSheet.getLastRow() > 1) {
            const fpData = fpSheet.getRange(2, 2, fpSheet.getLastRow() - 1, 1).getValues();
            dispensedEncounters = fpData.map(row => row[0]).filter(id => id);
          }

          // Count FP patients not yet dispensed
          for (let i = 1; i < triageData.length; i++) {
            const row = triageData[i];
            const encounterId = row[encounterIdIdx];
            const servicePoint = row[servicePointIdx];

            if (servicePoint === 'Family Planning' && encounterId && !dispensedEncounters.includes(encounterId)) {
              pendingFPService++;
            }
          }
        }
        // If columns don't exist, pendingFPService stays 0 (no error)
      }
    } catch (error) {
      Logger.log('Error counting FP queue: ' + error.toString());
      // pendingFPService stays 0 on error
    }



    return {
      triage: triageQueue,
      pending_lab_results: pendingLabResults,
      pending_billing: pendingBilling,
      pending_fp_service: pendingFPService,
      waiting_consultation: waitingConsultation,
      pending_lab_orders: pendingLabOrders
    };

  } catch (error) {
    Logger.log('Error in getLiveQueueCounts: ' + error.toString());
    return {
      triage: 0,
      pending_lab_results: 0,
      pending_billing: 0,
      pending_fp_service: 0,
      waiting_consultation: 0,
      pending_lab_orders: 0
    };
  }
}
