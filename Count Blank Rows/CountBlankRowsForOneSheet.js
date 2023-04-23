"use strict";
function main(workbook) {
    const sheet = workbook.getWorksheet('Sheet1');
    // const sheet = workbook.getActiveWorksheet(); // For active worksheet - not suitable for Power Automate related script.  
    const range = sheet.getUsedRange(true); // get value only 
    if (!range) {
        console.log(`No data on this sheet. `);
        return;
    }
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
    const values = range.getValues();
    let emptyRows = 0;
    for (let row of values) {
        let len = 0;
        for (let cell of row) {
            len = len + cell.toString().length;
        }
        if (len === 0) {
            emptyRows++;
        }
    }
    console.log(`Total empty row: ` + emptyRows);
    return emptyRows;
}
