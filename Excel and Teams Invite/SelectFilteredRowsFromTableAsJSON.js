"use strict";
function main(workbook) {
    console.log("Current date time: " + new Date().toUTCString());
    const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue();
    const sheet = workbook.getWorksheet('Interviews');
    const table = sheet.getTables()[0];
    const dataRows = table.getRange().getTexts();
    // or use this if there's no table
    // let dataRows = sheet.getUsedRange().getValues();
    const selectedRows = dataRows.filter((row, i) => {
        // Select header row and any data row with the status column equal to approach value
        return (row[1] === 'FALSE' || i === 0);
    });
    const recordDetails = returnObjectFromValues(selectedRows);
    const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
    console.log(JSON.stringify(inviteRecords));
    return inviteRecords;
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
    return objArray;
}
/**
 * Generate interview records by selecting required columns
 * @param records Input records
 * @param mins Number of minutes to add to the start date-time
 */
function generateInterviewRecords(records, mins) {
    const interviewinvites = [];
    records.forEach((record) => {
        // Interviewer 1    
        // If the start date-time is greather than current date-time, add to output records
        if ((new Date(record['Start time1'])) > new Date()) {
            console.log("selected " + new Date(record['Start time1']).toUTCString());
            let startTime = new Date(record['Start time1']).toISOString();
            // compute the finish time of the meeting
            let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
            interviewinvites.push({
                ID: record.ID,
                Candidate: record.Candidate,
                CandidateEmail: record['Candidate email'],
                CandidateContact: record['Candidate contact'],
                Interviewer: record.Interviewer1,
                InterviewerEmail: record['Interviewer1 email'],
                StartTime: startTime,
                FinishTime: finishTime
            });
        }
        else {
            console.log("Rejected " + (new Date(record['Start time1']).toUTCString()));
        }
        // Interviewer 2 
        // If the start date-time is greather than current date-time, add to output records
        if ((new Date(record['Start time2'])) > new Date()) {
            console.log("selected " + new Date(record['Start time2']).toUTCString());
            let startTime = new Date(record['Start time2']).toISOString();
            // compute the finish time of the meeting
            let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
            interviewinvites.push({
                ID: record.ID,
                Candidate: record.Candidate,
                CandidateEmail: record['Candidate email'],
                CandidateContact: record['Candidate contact'],
                Interviewer: record.Interviewer2,
                InterviewerEmail: record['Interviewer2 email'],
                StartTime: startTime,
                FinishTime: finishTime
            });
        }
        else {
            console.log("Rejected " + (new Date(record['Start time2']).toUTCString()));
        }
    });
    return interviewinvites;
}
/**
 * Add minutes to start date-time
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date time
 */
function addMins(startDateTime, mins) {
    return new Date(startDateTime.getTime() + mins * 60 * 1000);
}
