"use strict";
function main(workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}
function addRow(sheet, data) {
    const usedRange = sheet.getUsedRange();
    let startCell;
    // IF the sheet is empty, then use A1 as starting cell for update
    if (usedRange) {
        startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    }
    else {
        startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
