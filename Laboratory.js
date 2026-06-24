/**
 * myDoc - Laboratory & Inventory Module Functions
 */

function getPendingLabOrders() {
  try {
    Logger.log('=== getPendingLabOrders START ===');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const labSheet = ss.getSheetByName('Laboratory_Orders');
    const regSheet = ss.getSheetByName('Registration_Records');

    if (!labSheet) {
      Logger.log('ERROR: Laboratory_Orders sheet not found');
      return [];
    }

    if (!regSheet) {
      Logger.log('ERROR: Registration_Records sheet not found');
      return [];
    }

    Logger.log('Both sheets found successfully');

    const labData = labSheet.getDataRange().getValues();
    Logger.log('Total rows in lab sheet (including header): ' + labData.length);

    const headers = labData.shift();
    Logger.log('Headers: ' + JSON.stringify(headers));
    Logger.log('Data rows after removing header: ' + labData.length);

    // Log first row to see structure
    if (labData.length > 0) {
      Logger.log('First data row: ' + JSON.stringify(labData[0]));
      Logger.log('Status in first row (index 7): "' + labData[0][7] + '"');
      Logger.log('Status type: ' + typeof labData[0][7]);
    }

    // Get Patient Map
    const regData = regSheet.getDataRange().getValues();
    regData.shift();
    const patientMap = {};
    regData.forEach(r => {
      if (r[0]) {
        patientMap[r[0]] = (r[4] || '') + ' ' + (r[6] || '');
      }
    });
    Logger.log('Patient map created with ' + Object.keys(patientMap).length + ' patients');

    // Filter for 'Ordered' or 'Pending' (case-insensitive, trimmed)
    const filtered = labData.filter(r => {
      if (!r || !r[0]) {
        Logger.log('Skipping empty row');
        return false;
      }
      const status = String(r[7] || '').trim().toLowerCase();
      Logger.log('Row ' + r[0] + ' status: "' + r[7] + '" -> normalized: "' + status + '"');
      const matches = status === 'ordered' || status === 'pending';
      Logger.log('  Matches filter: ' + matches);
      return matches;
    });

    Logger.log('Filtered rows: ' + filtered.length);

    const result = filtered.map(r => {
      const orderTime = r[4];
      return {
        lab_order_id: r[0],
        patient_name: patientMap[r[1]] || 'Unknown',
        test_name: r[5] || 'No Test Name',
        order_time: orderTime instanceof Date ? orderTime.toISOString() : orderTime,
        status: r[7]
      };
    });

    Logger.log('Final result count: ' + result.length);
    Logger.log('=== getPendingLabOrders END ===');

    return result;
  } catch (error) {
    Logger.log('ERROR in getPendingLabOrders: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return [];
  }
}

function saveLabResult(data) {
  try {
    Logger.log('=== saveLabResult START ===');
    Logger.log('Lab Order ID: ' + data.lab_order_id);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Laboratory_Orders');
    const rows = sheet.getDataRange().getValues();

    // Find row by ID
    const rowIndex = rows.findIndex(r => r[0] === data.lab_order_id);
    if (rowIndex === -1) {
      Logger.log('ERROR: Order not found');
      return { success: false, message: 'Order not found' };
    }

    // Get patient ID and test name for inventory deduction
    const patientId = rows[rowIndex][1];
    const testName = rows[rowIndex][5];

    Logger.log('Patient ID: ' + patientId);
    Logger.log('Test Name: ' + testName);

    // Update columns: Status(7), Value(8), Units(9), Range(10), Flag(11), Tech(12), Description(13), FileURL(14)
    const techId = Session.getActiveUser().getEmail();

    sheet.getRange(rowIndex + 1, 8).setValue('Resulted'); // Status
    sheet.getRange(rowIndex + 1, 9).setValue(data.result_value);
    sheet.getRange(rowIndex + 1, 10).setValue(data.result_units || '');
    sheet.getRange(rowIndex + 1, 11).setValue(data.reference_range || '');
    sheet.getRange(rowIndex + 1, 12).setValue(data.abnormal_flag || 'Normal');
    sheet.getRange(rowIndex + 1, 13).setValue(techId);
    sheet.getRange(rowIndex + 1, 14).setValue(data.detailed_description || '');
    sheet.getRange(rowIndex + 1, 15).setValue(data.file_url || '');

    Logger.log('Lab result saved to sheet');

    // Auto-deduct inventory items for this test
    try {
      const deductionResult = deductInventoryForTest(data.lab_order_id, testName, patientId);
      Logger.log('Inventory deduction result: ' + JSON.stringify(deductionResult));

      if (deductionResult.success) {
        Logger.log('=== saveLabResult END (Success with inventory deduction) ===');
        return {
          success: true,
          message: 'Result saved. ' + deductionResult.message,
          inventory_deducted: true,
          items_deducted: deductionResult.items_deducted
        };
      } else {
        // Lab result saved but inventory deduction failed
        Logger.log('WARNING: Inventory deduction failed: ' + deductionResult.message);
        Logger.log('=== saveLabResult END (Success but inventory warning) ===');
        return {
          success: true,
          message: 'Result saved. WARNING: ' + deductionResult.message,
          inventory_deducted: false,
          inventory_warning: deductionResult.message
        };
      }
    } catch (inventoryError) {
      // Lab result saved but inventory deduction threw error
      Logger.log('ERROR in inventory deduction: ' + inventoryError.toString());
      Logger.log('=== saveLabResult END (Success but inventory error) ===');
      return {
        success: true,
        message: 'Result saved. WARNING: Could not deduct inventory items.',
        inventory_deducted: false,
        inventory_error: inventoryError.message
      };
    }
  } catch (error) {
    Logger.log('ERROR in saveLabResult: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    Logger.log('=== saveLabResult END (Error) ===');
    return {
      success: false,
      message: 'Failed to save result: ' + error.message
    };
  }
}


function getInventory() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Laboratory_Inventory_Records');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  data.shift();

  // Aggregate by Item ID
  const inventory = {};

  data.forEach(r => {
    const itemId = r[1];
    if (!inventory[itemId]) {
      inventory[itemId] = {
        item_id: itemId,
        item_name: r[2],
        stock: 0
      };
    }
    // Add Qty In, Subtract Qty Out
    inventory[itemId].stock += (Number(r[4]) || 0) - (Number(r[6]) || 0);
  });

  return Object.values(inventory);
}

function addInventoryTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Laboratory_Inventory_Records');

  const timestamp = new Date();

  // Calculate new balance (simplified, ideally we fetch current balance first)
  // For this prototype, we just log the transaction. The getInventory aggregates it.

  const rowData = [
    timestamp,
    data.item_id,
    data.item_name,
    data.transaction_type, // Purchase or Usage
    data.transaction_type === 'Purchase' ? data.quantity : 0,
    data.unit_cost || '',
    data.transaction_type === 'Usage' ? data.quantity : 0,
    '' // Stock Balance (Running total not implemented in this simple append, aggregation used instead)
  ];

  sheet.appendRow(rowData);
  return { success: true, message: 'Transaction saved.' };
}

function uploadLabResultFile(fileData, fileName, mimeType, labOrderId) {
  try {
    // Get or create "Lab Results Files" folder
    const folders = DriveApp.getFoldersByName('Lab Results Files');
    let mainFolder;

    if (folders.hasNext()) {
      mainFolder = folders.next();
    } else {
      mainFolder = DriveApp.createFolder('Lab Results Files');
    }

    // Get patient ID and test name from lab order
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Laboratory_Orders');
    const rows = sheet.getDataRange().getValues();
    const rowIndex = rows.findIndex(r => r[0] === labOrderId);

    if (rowIndex === -1) {
      throw new Error('Lab order not found');
    }

    const patientId = rows[rowIndex][1] || 'Unknown';
    const testName = rows[rowIndex][5] || 'Test';

    // Create/get patient folder
    const patientFolders = mainFolder.getFoldersByName(patientId);
    let patientFolder;

    if (patientFolders.hasNext()) {
      patientFolder = patientFolders.next();
    } else {
      patientFolder = mainFolder.createFolder(patientId);
    }

    // Create/get test folder
    const testFolders = patientFolder.getFoldersByName(testName);
    let testFolder;

    if (testFolders.hasNext()) {
      testFolder = testFolders.next();
    } else {
      testFolder = patientFolder.createFolder(testName);
    }

    // Decode base64 and create file
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = testFolder.createFile(blob);

    // Set file sharing to anyone with link can view
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Return file URL
    return file.getUrl();

  } catch (error) {
    Logger.log('Error uploading file: ' + error.toString());
    throw new Error('Failed to upload file: ' + error.message);
  }
}
