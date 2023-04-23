"use strict";
function main(workbook) {
    const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
    const range = table.getRange();
    // If you know the table name, you can also do this...
    // const table = workbook.getTable('Table13436');
    const texts = table.getRange().getTexts();
    let returnObjects = [];
    if (table.getRowCount() > 0) {
        returnObjects = returnObjectFromValues(texts, range);
    }
    console.log(JSON.stringify(returnObjects));
    return returnObjects;
}
function returnObjectFromValues(values, range) {
    let objArray = [];
    let objKeys = [];
    for (let i = 0; i < values.length; i++) {
        if (i === 0) {
            objKeys = values[i];
            continue;
        }
        let obj = {};
        for (let j = 0; j < values[i].length; j++) {
            // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
            if (j === 4) {
                obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
            }
            else {
                obj[objKeys[j]] = values[i][j];
            }
        }
        objArray.push(obj);
    }
    return objArray;
}
