/**
 * TEMPORARY TEST FUNCTION - Simple version without stock calculation
 * Use this to test if deployment is working
 */
function getInventoryItemsSimple() {
    try {
        Logger.log('=== getInventoryItemsSimple START ===');
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const itemsSheet = ss.getSheetByName('Lab_Inventory_Items');

        if (!itemsSheet || itemsSheet.getLastRow() < 2) {
            Logger.log('No items found');
            return [];
        }

        const data = itemsSheet.getDataRange().getValues();
        data.shift(); // Remove header

        const items = data.map(row => {
            return {
                item_id: String(row[0]),
                item_name: String(row[1]),
                unit_of_measure: String(row[2]),
                reorder_level: Number(row[3]) || 0
            };
        });

        Logger.log('Items created: ' + items.length);
        return items;
    } catch (error) {
        Logger.log('ERROR: ' + error.toString());
        return [];
    }
}
