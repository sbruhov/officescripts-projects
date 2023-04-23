"use strict";
function main(workbook) {
    var _a;
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    (_a = workbook.getWorksheet('Combined')) === null || _a === void 0 ? void 0 : _a.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
        const tables = workbook.getWorksheet(sheet).getTables();
        if (!targetTableCreated) {
            const headerValues = tables[0].getHeaderRowRange().getTexts();
            const targetRange = updateRange(newSheet, headerValues);
            combinedTable = newSheet.addTable(targetRange.getAddress(), true);
            targetTableCreated = true;
        }
        for (let table of tables) {
            let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
            let rowCount = table.getRowCount();
            if (rowCount > 0) {
                combinedTable.addRows(-1, dataValues);
            }
        }
    });
}
function updateRange(sheet, data) {
    const targetRange = sheet.getRange('A1').getResizedRange(data.length - 1, data[0].length - 1);
    targetRange.setValues(data);
    return targetRange;
}
