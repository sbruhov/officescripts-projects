"use strict";
function main(workbook) {
    const table = workbook.getWorksheet('PlainTable').getTables()[0];
    // If you know the table name, you can also do this...
    // const table = workbook.getTable('Table13436');
    const texts = table.getRange().getTexts();
    let returnObjects = [];
    if (table.getRowCount() > 0) {
        returnObjects = returnObjectFromValues(texts);
    }
    console.log(JSON.stringify(returnObjects));
    return returnObjects;
}
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
    return objArray;
}
