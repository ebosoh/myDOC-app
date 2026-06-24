/**
 * myDoc - Reports & Statistics Module Functions
 */

function getClinicStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Total Patients
  const regSheet = ss.getSheetByName('Registration_Records');
  const totalPatients = regSheet ? regSheet.getLastRow() - 1 : 0;

  // 2. Revenue Today
  const billSheet = ss.getSheetByName('Billing_and_Payment_Records');
  let revenueToday = 0;
  let revenueTotal = 0;

  if (billSheet && billSheet.getLastRow() > 1) {
    const data = billSheet.getDataRange().getValues();
    data.shift();
    const today = new Date().toDateString();

    data.forEach(r => {
      const amount = Number(r[6]) || 0; // Amount Paid Today
      revenueTotal += amount;
      if (new Date(r[3]).toDateString() === today) {
        revenueToday += amount;
      }
    });
  }

  // 3. Triage Stats (BMI Categories)
  const triageSheet = ss.getSheetByName('Triage_Records');
  const bmiStats = { Underweight: 0, Normal: 0, Overweight: 0, Obese: 0 };

  if (triageSheet && triageSheet.getLastRow() > 1) {
    const data = triageSheet.getDataRange().getValues();
    data.shift();
    data.forEach(r => {
      const bmi = Number(r[13]); // BMI column
      if (bmi < 18.5) bmiStats.Underweight++;
      else if (bmi < 25) bmiStats.Normal++;
      else if (bmi < 30) bmiStats.Overweight++;
      else if (bmi >= 30) bmiStats.Obese++;
    });
  }

  return {
    totalPatients,
    revenueToday,
    revenueTotal,
    bmiStats
  };
}

function getStatistics() {
  try {
    Logger.log('=== getStatistics START ===');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const stats = {};

    // 1. Registration Demographics
    const regSheet = ss.getSheetByName('Registration_Records');
    if (regSheet && regSheet.getLastRow() > 1) {
      const regData = regSheet.getDataRange().getValues();
      regData.shift(); // Remove header

      stats.gender = countByField(regData, 8); // gender
      stats.maritalStatus = countByField(regData, 9); // marital_status
      stats.counties = countByField(regData, 12); // county
      stats.subCounties = countByField(regData, 13); // sub_county
      stats.wards = countByField(regData, 14); // ward
      stats.languages = countByField(regData, 15); // language
      stats.bloodGroups = countByField(regData, 19); // blood_group
      stats.allergies = countByField(regData, 20); // known_allergies
    }

    // 2. Triage Statistics
    const triageSheet = ss.getSheetByName('Triage_Records');
    if (triageSheet && triageSheet.getLastRow() > 1) {
      const triageData = triageSheet.getDataRange().getValues();
      triageData.shift();

      stats.triagePriority = countByField(triageData, 14); // triage_priority

      // Oxygen saturation ranges
      stats.oxygenSaturation = {
        'Low (<90%)': 0,
        'Moderate (90-95%)': 0,
        'Normal (95-100%)': 0
      };
      triageData.forEach(r => {
        const o2 = parseFloat(r[10]); // oxygen_saturation
        if (!isNaN(o2)) {
          if (o2 < 90) stats.oxygenSaturation['Low (<90%)']++;
          else if (o2 < 95) stats.oxygenSaturation['Moderate (90-95%)']++;
          else if (o2 <= 100) stats.oxygenSaturation['Normal (95-100%)']++;
        }
      });
    }

    // 3. Treatment Statistics
    const treatmentSheet = ss.getSheetByName('Treatment_Records');
    if (treatmentSheet && treatmentSheet.getLastRow() > 1) {
      const treatmentData = treatmentSheet.getDataRange().getValues();
      treatmentData.shift();

      stats.disposition = countByField(treatmentData, 14); // disposition

      // Prescriptions
      const prescriptions = {};
      treatmentData.forEach(r => {
        const medsJson = r[10]; // prescribed_medications
        if (medsJson) {
          try {
            const meds = JSON.parse(medsJson);
            meds.forEach(med => {
              prescriptions[med.name] = (prescriptions[med.name] || 0) + 1;
            });
          } catch (e) { }
        }
      });
      stats.prescriptions = prescriptions;
    }

    // 4. Lab Statistics
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    if (labSheet && labSheet.getLastRow() > 1) {
      const labData = labSheet.getDataRange().getValues();
      labData.shift();

      stats.labTests = countByField(labData, 5); // test_name
    }

    // 5. Payment Statistics
    const billingSheet = ss.getSheetByName('Billing_and_Payment_Records');
    if (billingSheet && billingSheet.getLastRow() > 1) {
      const billingData = billingSheet.getDataRange().getValues();
      billingData.shift();

      stats.paymentMethods = countByField(billingData, 7); // payment_method
    }

    Logger.log('=== getStatistics END ===');
    return { success: true, stats: stats };
  } catch (error) {
    Logger.log('ERROR in getStatistics: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function countByField(data, columnIndex) {
  const counts = {};
  data.forEach(row => {
    const value = row[columnIndex];
    if (value && value !== '' && value !== null) {
      const key = String(value).trim();
      counts[key] = (counts[key] || 0) + 1;
    }
  });
  return counts;
}

// Calendar Module
function getCalendarEvents() {
  const now = new Date();
  const sevenDaysFromNow = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));

  try {
    // Use configured CALENDAR_ID
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      return [{ title: 'Error', description: 'Calendar not found. Check CALENDAR_ID.', startTime: now.toLocaleString(), endTime: now.toLocaleString() }];
    }

    const events = calendar.getEvents(now, sevenDaysFromNow);

    return events.map(e => ({
      id: e.getId(),
      title: e.getTitle(),
      startTime: e.getStartTime().toLocaleString(),
      endTime: e.getEndTime().toLocaleString(),
      description: e.getDescription()
    }));
  } catch (e) {
    return [{ title: 'Error', description: e.toString(), startTime: now.toLocaleString(), endTime: now.toLocaleString() }];
  }
}

function createCalendarEvent(data) {
  try {
    // Use configured CALENDAR_ID
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      return { success: false, message: 'Calendar not found. Check CALENDAR_ID.' };
    }

    const start = new Date(data.startTime);
    const end = new Date(data.endTime);

    const event = calendar.createEvent(data.title, start, end, {
      description: data.description
    });

    return { success: true, message: 'Appointment booked.' };
  } catch (e) {
    return { success: false, message: 'Error booking appointment: ' + e.toString() };
  }
}
