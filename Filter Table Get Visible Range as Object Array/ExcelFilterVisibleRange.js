"use strict";
function main(workbook) {
    const table1 = workbook.getTable("Table1");
    const keyColumnValues = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0]);
    // const uniqueKeys= [...Array.from(new Set(keyColumnValues))];
    const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);
    console.log(uniqueKeys);
    const returnObj = {};
    uniqueKeys.forEach((key) => {
        table1.getColumnByName('Station').getFilter()
            .applyValuesFilter([key]);
        const rangeView = table1.getRange().getVisibleView();
        returnObj[key] = returnObjectFromValues(rangeView.getValues());
    });
    table1.getColumnByName('Station').getFilter().clear();
    console.log(JSON.stringify(returnObj));
    return returnObj;
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
