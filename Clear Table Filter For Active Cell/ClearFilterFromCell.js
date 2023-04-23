"use strict";
function main(workbook) {
    // Get active cell
    const cell = workbook.getActiveCell();
    // Get all tables associated with that cell 
    const tables = cell.getTables();
    // If there is no table on the selection, return
    if (tables.length !== 1) {
        console.log("The selection is not in a table.");
        return;
    }
    // Get table (since it is already determined that there is only a single table part of the selection )
    const currentTable = tables[0];
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());
    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());
    const headerCellValue = intersect.getCell(0, 0).getValue();
    console.log(headerCellValue);
    // Get column
    const col = currentTable.getColumnByName(headerCellValue);
    // Clear filter
    col.getFilter().clear();
}
