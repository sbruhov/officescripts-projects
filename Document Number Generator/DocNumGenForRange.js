"use strict";
function main(workbook, inputString) {
    // Temporary placeholder for testing 
    const incoming = {
        docType: 'form',
        documentName: 'cilx.png'
    };
    inputString = JSON.stringify(incoming);
    // End temp testing area
    // Object to hold key prefixes for each document type
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    };
    // Length of the numeric part of the key
    const KEYLENGTH = 6;
    // Parse the incoming string as object 
    const input = JSON.parse(inputString);
    // Reject invalid request 
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`;
    }
    // Get existing data in the sheet 
    const sheet = workbook.getWorksheet('PlainSheet'); /* plain range sheet */
    const range = sheet.getUsedRange();
    const data = range.getValues();
    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();
    // Get the max document Id for the type
    const maxId = selectIds[selectIds.length - 1];
    // Extract numeric part 
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;
    // If we ever reach the max key numbrer, throw an error
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`;
    }
    // Get the correct prefix value
    const prefixVal = PREFIX[input.docType.toLowerCase()];
    // Compute next key value
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);
    // Get last row and compute next row address
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);
    // Add a row with incoming data plus the computed key value
    target.setValues([
        [
            nextKey,
            /* Capitalize the document type */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]
    ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`);
    // Return the key value recorded in Excel
    return nextKey;
}
