"use strict";
function main(workbook) {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    let records = [];
    for (let row of rows) {
        let [event, date, location, capacity] = row;
        records.push({
            event: event,
            date: date,
            location: location,
            capacity: capacity
        });
    }
    console.log(JSON.stringify(records));
    return records;
}
