"use strict";
function main(workbook) {
    const sheet = workbook.getWorksheet('Interviews');
    const table = sheet.getTables()[0];
    const dataRows = table.getRange().getTexts();
    // or
    // let dataRows = sheet.getUsedRange().getValues();
    const recordDetails = returnObjectFromValues(dataRows);
    console.log(recordDetails);
    return recordDetails;
}
/**
 * This helper funciton converts table values into an object array.
 */
function returnObjectFromValues(values) {
    let objArray = [];
    let objKeys = [];
    for (let i = 0; i < values.length; i++) {
        if (i === 0) {
            objKeys = values[i];
            continue;
        }
        let obj = {};
        for (let j = 0; j < values[i].length; j++) {
            obj[objKeys[j]] = values[i][j];
        }
        objArray.push(obj);
    }
    console.log(JSON.stringify(objArray));
    return objArray;
}
