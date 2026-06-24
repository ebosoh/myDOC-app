/**
 * myDoc - Lab Inventory Module Functions
 * Handles inventory management separate from lab operations
 */

// ==================== MASTER ITEM LIST ====================

/**
 * Get all inventory items with current stock levels
 */
function getInventoryItems() {
    try {
        Logger.log('=== getInventoryItems START ===');
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        Logger.log('Spreadsheet ID: ' + ss.getId());

        const itemsSheet = ss.getSheetByName('Lab_Inventory_Items');
        Logger.log('Sheet found: ' + (itemsSheet ? 'YES' : 'NO'));

        if (!itemsSheet || itemsSheet.getLastRow() < 2) {
            Logger.log('Returning empty array - Sheet: ' + (itemsSheet ? 'exists' : 'not found') + ', Last row: ' + (itemsSheet ? itemsSheet.getLastRow() : 'N/A'));
            return [];
        }

        const data = itemsSheet.getDataRange().getValues();
        Logger.log('Total rows (including header): ' + data.length);
        data.shift(); // Remove header
        Logger.log('Data rows after removing header: ' + data.length);

        const items = data.map(row => {
            const itemId = row[0];
            // TEMPORARY: Skip stock calculation to test serialization
            const currentStock = 0; // calculateStockLevel(itemId);
            const reorderLevel = row[3] || 0;
            const stockPercentage = 0; // reorderLevel > 0 ? (currentStock / reorderLevel) * 100 : 100;

            return {
                item_id: String(itemId),
                item_name: String(row[1]),
                unit_of_measure: String(row[2]),
                reorder_level: Number(reorderLevel),
                current_stock: Number(currentStock),
                stock_percentage: Number(stockPercentage),
                stock_status: 'green', // getStockStatus(stockPercentage),
                created_date: row[4] ? row[4].toString() : '',
                created_by: String(row[5] || '')
            };
        });

        Logger.log('Items created: ' + items.length);
        Logger.log('First item: ' + JSON.stringify(items[0]));
        Logger.log('=== getInventoryItems END ===');
        return items;
    } catch (error) {
        Logger.log('ERROR in getInventoryItems: ' + error.toString());
        Logger.log('Error stack: ' + error.stack);
        return [];
    }
}

/**
 * Add a new item to the master inventory list
 */
function addInventoryItem(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let itemsSheet = ss.getSheetByName('Lab_Inventory_Items');

        // Create sheet if it doesn't exist
        if (!itemsSheet) {
            itemsSheet = ss.insertSheet('Lab_Inventory_Items');
            itemsSheet.appendRow([
                'Item_ID', 'Item_Name', 'Unit_of_Measure', 'Reorder_Level',
                'Created_Date', 'Created_By'
            ]);
        }

        // Generate Item ID
        const lastRow = itemsSheet.getLastRow();
        const itemId = 'ITEM-' + String(lastRow).padStart(5, '0');

        const timestamp = new Date();
        const userId = Session.getActiveUser().getEmail();

        const rowData = [
            itemId,
            data.item_name,
            data.unit_of_measure,
            data.reorder_level || 100,
            timestamp,
            userId
        ];

        itemsSheet.appendRow(rowData);

        return {
            success: true,
            message: 'Item added successfully.',
            item_id: itemId
        };
    } catch (error) {
        Logger.log('ERROR in addInventoryItem: ' + error.toString());
        return {
            success: false,
            message: 'Failed to add item: ' + error.message
        };
    }
}

// ==================== PURCHASES (INCOMING STOCK) ====================

/**
 * Record a purchase transaction
 */
function recordPurchase(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let purchaseSheet = ss.getSheetByName('Lab_Inventory_Purchases');

        // Create sheet if it doesn't exist
        if (!purchaseSheet) {
            purchaseSheet = ss.insertSheet('Lab_Inventory_Purchases');
            purchaseSheet.appendRow([
                'Transaction_ID', 'Item_ID', 'Item_Name', 'Quantity', 'Unit_Cost',
                'Total_Cost', 'Supplier', 'Manufacture_Date', 'Expiry_Date',
                'Days_to_Expiry', 'Purchase_Date', 'Purchased_By', 'Batch_Number'
            ]);
        }

        // Generate Transaction ID
        const lastRow = purchaseSheet.getLastRow();
        const transactionId = 'PUR-' + String(lastRow).padStart(6, '0');

        const timestamp = new Date();
        const userId = Session.getActiveUser().getEmail();

        // Calculate days to expiry
        const expiryDate = new Date(data.expiry_date);
        const today = new Date();
        const daysToExpiry = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

        const totalCost = data.quantity * data.unit_cost;

        const rowData = [
            transactionId,
            data.item_id,
            data.item_name,
            data.quantity,
            data.unit_cost,
            totalCost,
            data.supplier || '',
            data.manufacture_date,
            data.expiry_date,
            daysToExpiry,
            timestamp,
            userId,
            data.batch_number || ''
        ];

        purchaseSheet.appendRow(rowData);

        return {
            success: true,
            message: 'Purchase recorded successfully.',
            transaction_id: transactionId
        };
    } catch (error) {
        Logger.log('ERROR in recordPurchase: ' + error.toString());
        return {
            success: false,
            message: 'Failed to record purchase: ' + error.message
        };
    }
}

/**
 * Get purchase history with optional filters
 */
function getPurchaseHistory(filters) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const purchaseSheet = ss.getSheetByName('Lab_Inventory_Purchases');

        if (!purchaseSheet || purchaseSheet.getLastRow() < 2) {
            return [];
        }

        const data = purchaseSheet.getDataRange().getValues();
        data.shift(); // Remove header

        let purchases = data.map(row => {
            const daysToExpiry = row[9];
            return {
                transaction_id: row[0],
                item_id: row[1],
                item_name: row[2],
                quantity: row[3],
                unit_cost: row[4],
                total_cost: row[5],
                supplier: row[6],
                manufacture_date: row[7] ? row[7].toString() : '',
                expiry_date: row[8] ? row[8].toString() : '',
                days_to_expiry: daysToExpiry,
                expiry_status: getExpiryStatus(daysToExpiry),
                purchase_date: row[10] ? row[10].toString() : '',
                purchased_by: row[11],
                batch_number: row[12]
            };
        });

        // Apply filters if provided
        if (filters && filters.expiry_status) {
            purchases = purchases.filter(p => p.expiry_status === filters.expiry_status);
        }

        return purchases;
    } catch (error) {
        Logger.log('ERROR in getPurchaseHistory: ' + error.toString());
        return [];
    }
}

/**
 * Get expiry statistics (percentage of items in each category)
 */
function getExpiryStatistics() {
    try {
        const purchases = getPurchaseHistory();

        if (purchases.length === 0) {
            return {
                expired: 0,
                orange: 0,
                yellow: 0,
                green: 0
            };
        }

        const counts = {
            expired: 0,
            orange: 0,
            yellow: 0,
            green: 0
        };

        purchases.forEach(p => {
            counts[p.expiry_status]++;
        });

        const total = purchases.length;

        return {
            expired: Math.round((counts.expired / total) * 100),
            orange: Math.round((counts.orange / total) * 100),
            yellow: Math.round((counts.yellow / total) * 100),
            green: Math.round((counts.green / total) * 100)
        };
    } catch (error) {
        Logger.log('ERROR in getExpiryStatistics: ' + error.toString());
        return { expired: 0, orange: 0, yellow: 0, green: 0 };
    }
}

// ==================== USAGE (OUTGOING STOCK) ====================

/**
 * Record usage transaction (called automatically when test is performed)
 */
function recordUsage(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let usageSheet = ss.getSheetByName('Lab_Inventory_Usage');

        // Create sheet if it doesn't exist
        if (!usageSheet) {
            usageSheet = ss.insertSheet('Lab_Inventory_Usage');
            usageSheet.appendRow([
                'Transaction_ID', 'Item_ID', 'Item_Name', 'Quantity_Used',
                'Lab_Order_ID', 'Test_Name', 'Patient_ID', 'Usage_Date', 'Recorded_By'
            ]);
        }

        // Generate Transaction ID
        const lastRow = usageSheet.getLastRow();
        const transactionId = 'USE-' + String(lastRow).padStart(6, '0');

        const timestamp = new Date();
        const userId = Session.getActiveUser().getEmail();

        const rowData = [
            transactionId,
            data.item_id,
            data.item_name,
            data.quantity_used,
            data.lab_order_id,
            data.test_name,
            data.patient_id,
            timestamp,
            userId
        ];

        usageSheet.appendRow(rowData);

        return {
            success: true,
            message: 'Usage recorded successfully.',
            transaction_id: transactionId
        };
    } catch (error) {
        Logger.log('ERROR in recordUsage: ' + error.toString());
        return {
            success: false,
            message: 'Failed to record usage: ' + error.message
        };
    }
}

/**
 * Get usage history
 */
function getUsageHistory() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const usageSheet = ss.getSheetByName('Lab_Inventory_Usage');

        if (!usageSheet || usageSheet.getLastRow() < 2) {
            return [];
        }

        const data = usageSheet.getDataRange().getValues();
        data.shift(); // Remove header

        const usage = data.map(row => ({
            transaction_id: row[0],
            item_id: row[1],
            item_name: row[2],
            quantity_used: row[3],
            lab_order_id: row[4],
            test_name: row[5],
            patient_id: row[6],
            usage_date: row[7] ? row[7].toString() : '',
            recorded_by: row[8]
        }));

        return usage;
    } catch (error) {
        Logger.log('ERROR in getUsageHistory: ' + error.toString());
        return [];
    }
}

/**
 * Get stock level statistics
 */
function getStockStatistics() {
    try {
        const items = getInventoryItems();

        if (items.length === 0) {
            return {
                red: 0,
                yellow: 0,
                orange: 0,
                green: 0
            };
        }

        const counts = {
            red: 0,
            yellow: 0,
            orange: 0,
            green: 0
        };

        items.forEach(item => {
            counts[item.stock_status]++;
        });

        const total = items.length;

        return {
            red: Math.round((counts.red / total) * 100),
            yellow: Math.round((counts.yellow / total) * 100),
            orange: Math.round((counts.orange / total) * 100),
            green: Math.round((counts.green / total) * 100)
        };
    } catch (error) {
        Logger.log('ERROR in getStockStatistics: ' + error.toString());
        return { red: 0, yellow: 0, orange: 0, green: 0 };
    }
}

// ==================== TEST-ITEM MAPPING ====================

/**
 * Save test-item mapping
 */
function saveTestItemMapping(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let mappingSheet = ss.getSheetByName('Lab_Test_Items_Mapping');

        // Create sheet if it doesn't exist
        if (!mappingSheet) {
            mappingSheet = ss.insertSheet('Lab_Test_Items_Mapping');
            mappingSheet.appendRow([
                'Mapping_ID', 'Test_Name', 'Item_ID', 'Item_Name',
                'Quantity_Required', 'Created_Date', 'Created_By'
            ]);
        }

        // Generate Mapping ID
        const lastRow = mappingSheet.getLastRow();
        const mappingId = 'MAP-' + String(lastRow).padStart(5, '0');

        const timestamp = new Date();
        const userId = Session.getActiveUser().getEmail();

        const rowData = [
            mappingId,
            data.test_name,
            data.item_id,
            data.item_name,
            data.quantity_required,
            timestamp,
            userId
        ];

        mappingSheet.appendRow(rowData);

        return {
            success: true,
            message: 'Test-item mapping saved successfully.',
            mapping_id: mappingId
        };
    } catch (error) {
        Logger.log('ERROR in saveTestItemMapping: ' + error.toString());
        return {
            success: false,
            message: 'Failed to save mapping: ' + error.message
        };
    }
}

/**
 * Get items required for a specific test
 */
function getTestItemMapping(testName) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const mappingSheet = ss.getSheetByName('Lab_Test_Items_Mapping');

        if (!mappingSheet || mappingSheet.getLastRow() < 2) {
            return [];
        }

        const data = mappingSheet.getDataRange().getValues();
        data.shift(); // Remove header

        const mappings = data
            .filter(row => row[1] === testName)
            .map(row => ({
                mapping_id: row[0],
                test_name: row[1],
                item_id: row[2],
                item_name: row[3],
                quantity_required: row[4],
                created_date: row[5],
                created_by: row[6]
            }));

        return mappings;
    } catch (error) {
        Logger.log('ERROR in getTestItemMapping: ' + error.toString());
        return [];
    }
}

/**
 * Get all test mappings
 */
function getAllTestMappings() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const mappingSheet = ss.getSheetByName('Lab_Test_Items_Mapping');

        if (!mappingSheet || mappingSheet.getLastRow() < 2) {
            return [];
        }

        const data = mappingSheet.getDataRange().getValues();
        data.shift(); // Remove header

        const mappings = data.map(row => ({
            mapping_id: row[0],
            test_name: row[1],
            item_id: row[2],
            item_name: row[3],
            quantity_required: row[4],
            created_date: row[5] ? row[5].toString() : '',
            created_by: row[6]
        }));

        return mappings;
    } catch (error) {
        Logger.log('ERROR in getAllTestMappings: ' + error.toString());
        return [];
    }
}

/**
 * Delete a test-item mapping
 */
function deleteTestItemMapping(mappingId) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const mappingSheet = ss.getSheetByName('Lab_Test_Items_Mapping');

        if (!mappingSheet) {
            return { success: false, message: 'Mapping sheet not found' };
        }

        const data = mappingSheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[0] === mappingId);

        if (rowIndex === -1) {
            return { success: false, message: 'Mapping not found' };
        }

        mappingSheet.deleteRow(rowIndex + 1);

        return { success: true, message: 'Mapping deleted successfully' };
    } catch (error) {
        Logger.log('ERROR in deleteTestItemMapping: ' + error.toString());
        return {
            success: false,
            message: 'Failed to delete mapping: ' + error.message
        };
    }
}

// ==================== AUTO-DEDUCTION ====================

/**
 * Automatically deduct inventory when a test is performed
 */
function deductInventoryForTest(labOrderId, testName, patientId) {
    try {
        Logger.log('=== deductInventoryForTest START ===');
        Logger.log('Lab Order ID: ' + labOrderId);
        Logger.log('Test Name: ' + testName);
        Logger.log('Patient ID: ' + patientId);

        // Get items required for this test
        const requiredItems = getTestItemMapping(testName);

        if (requiredItems.length === 0) {
            Logger.log('No items mapped for test: ' + testName);
            return {
                success: true,
                message: 'No inventory items configured for this test.',
                items_deducted: 0
            };
        }

        const deductedItems = [];
        const errors = [];

        // Deduct each required item
        requiredItems.forEach(item => {
            const currentStock = calculateStockLevel(item.item_id);

            if (currentStock < item.quantity_required) {
                errors.push(`Insufficient stock for ${item.item_name}. Required: ${item.quantity_required}, Available: ${currentStock}`);
                Logger.log('ERROR: Insufficient stock for ' + item.item_name);
            } else {
                // Record usage
                const usageResult = recordUsage({
                    item_id: item.item_id,
                    item_name: item.item_name,
                    quantity_used: item.quantity_required,
                    lab_order_id: labOrderId,
                    test_name: testName,
                    patient_id: patientId
                });

                if (usageResult.success) {
                    deductedItems.push(item.item_name);
                    Logger.log('Deducted: ' + item.quantity_required + ' x ' + item.item_name);
                } else {
                    errors.push(`Failed to deduct ${item.item_name}: ${usageResult.message}`);
                }
            }
        });

        Logger.log('=== deductInventoryForTest END ===');

        if (errors.length > 0) {
            return {
                success: false,
                message: 'Some items could not be deducted: ' + errors.join('; '),
                items_deducted: deductedItems.length,
                errors: errors
            };
        }

        return {
            success: true,
            message: `Successfully deducted ${deductedItems.length} item(s) from inventory.`,
            items_deducted: deductedItems.length,
            items: deductedItems
        };
    } catch (error) {
        Logger.log('ERROR in deductInventoryForTest: ' + error.toString());
        return {
            success: false,
            message: 'Failed to deduct inventory: ' + error.message
        };
    }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Calculate current stock level for an item
 */
function calculateStockLevel(itemId) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

        // Get total purchases
        const purchaseSheet = ss.getSheetByName('Lab_Inventory_Purchases');
        let totalPurchased = 0;

        if (purchaseSheet && purchaseSheet.getLastRow() > 1) {
            const purchaseData = purchaseSheet.getDataRange().getValues();
            purchaseData.shift(); // Remove header

            purchaseData.forEach(row => {
                if (row[1] === itemId) { // Item_ID column
                    totalPurchased += Number(row[3]) || 0; // Quantity column
                }
            });
        }

        // Get total usage
        const usageSheet = ss.getSheetByName('Lab_Inventory_Usage');
        let totalUsed = 0;

        if (usageSheet && usageSheet.getLastRow() > 1) {
            const usageData = usageSheet.getDataRange().getValues();
            usageData.shift(); // Remove header

            usageData.forEach(row => {
                if (row[1] === itemId) { // Item_ID column
                    totalUsed += Number(row[3]) || 0; // Quantity_Used column
                }
            });
        }

        return totalPurchased - totalUsed;
    } catch (error) {
        Logger.log('ERROR in calculateStockLevel: ' + error.toString());
        return 0;
    }
}

/**
 * Get expiry status based on days to expiry
 */
function getExpiryStatus(daysToExpiry) {
    if (daysToExpiry < 0) return 'expired';
    if (daysToExpiry < 30) return 'orange';
    if (daysToExpiry <= 60) return 'yellow';
    return 'green';
}

/**
 * Get stock status based on stock percentage
 */
function getStockStatus(stockPercentage) {
    if (stockPercentage < 10) return 'red';
    if (stockPercentage < 30) return 'yellow';
    if (stockPercentage < 40) return 'orange';
    return 'green';
}

/**
 * Delete a purchase record
 */
function deletePurchaseRecord(transactionId) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const purchaseSheet = ss.getSheetByName('Lab_Inventory_Purchases');

        if (!purchaseSheet) {
            return { success: false, message: 'Purchase sheet not found' };
        }

        const data = purchaseSheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[0] === transactionId);

        if (rowIndex === -1) {
            return { success: false, message: 'Purchase record not found' };
        }

        purchaseSheet.deleteRow(rowIndex + 1);

        return { success: true, message: 'Purchase record deleted successfully' };
    } catch (error) {
        Logger.log('ERROR in deletePurchaseRecord: ' + error.toString());
        return {
            success: false,
            message: 'Failed to delete purchase record: ' + error.message
        };
    }
}

