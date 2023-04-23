"use strict";
function main(workbook) {
    var _a;
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    let sheet1 = workbook.getWorksheet("Sheet1");
    const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
    const rows = table.getRange().getTexts();
    const selectColumns = rows.map((row) => {
        return [row[2], row[5]];
    });
    table.setShowTotals(true);
    selectColumns.splice(selectColumns.length - 1, 1);
    console.log(selectColumns);
    (_a = workbook.getWorksheet('ChartSheet')) === null || _a === void 0 ? void 0 : _a.delete();
    const chartSheet = workbook.addWorksheet('ChartSheet');
    const targetRange = updateRange(chartSheet, selectColumns);
    // Insert chart on sheet 'Sheet1'
    let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
    chart_2.setPosition('D1');
    const chartImage = chart_2.getImage();
    const tableImage = table.getRange().getImage();
    return {
        chartImage,
        tableImage
    };
}
function updateRange(sheet, data) {
    const targetRange = sheet.getRange('A1').getResizedRange(data.length - 1, data[0].length - 1);
    targetRange.setValues(data);
    return targetRange;
}
