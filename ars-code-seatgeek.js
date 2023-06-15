//This is the ARS Engine [month] link
var sourceSheetLink = 'https://docs.google.com/spreadsheets/d/12QFs0DONcuxz1b6DA-IFZgIzU-th-I8bFFUIJyc9Jzg/edit#gid=149758626';

function getDataForARSFAHW() {
    importData();
    getSchedule('Schedules', 38);
    getDistribution('Schedules', 116, 41, 33);
    getDistribution('Schedules', 149, 74, 33);
    getSchedule('Projection Training', 38);
    getTrainingAttendance();
}

function importData() {
    removeDataAndSchedules('RD', 18);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('Talkdesk Hours');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('H6').getValue();
    var data = sourceTab.getRange(1, 1, rows, 18).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet()
    var desSheet = thiSheet.getSheetByName('RD')
    var rango = desSheet.getRange(1, 1, rows, 18);
    rango.setValues(data)
}

function getSchedule(tabName, columnRange) {
    removeDataAndSchedules(tabName, columnRange);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName(tabName);
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('H7').getValue();
    var data = sourceTab.getRange(1, 1, rows, columnRange).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desSheet = thiSheet.getSheetByName(tabName);
    var rango = desSheet.getRange(1, 1, rows, columnRange);
    rango.setValues(data)
}


function getDistribution(tabName, targetColumnPosition, columnPosition, numberColumns) {
    removeDistribution(tabName, columnPosition, numberColumns)
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName(tabName);
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('H7').getValue();
    var data = sourceTab.getRange(1, targetColumnPosition, rows, numberColumns).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desSheet = thiSheet.getSheetByName(tabName);
    var rango = desSheet.getRange(1, columnPosition, rows, numberColumns);
    rango.setValues(data)
}

function getTrainingAttendance() {
    removeDataAndSchedules('Training Attendance', 14);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('Attendance Consolidated');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('H8').getValue();
    var data = sourceTab.getRange(1, 1, rows, 14).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet()
    var desSheet = thiSheet.getSheetByName('Training Attendance')
    var rango = desSheet.getRange(1, 1, rows, 14);
    rango.setValues(data)
}

function removeDataAndSchedules(sheetName, numberColumns) {
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desTab = thiSheet.getSheetByName(sheetName);
    desTab.getRange(1, 1, desTab.getMaxRows(), numberColumns).clearContent();
}

function removeDistribution(sheetName, columnPosition, numberColumns) {
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desTab = thiSheet.getSheetByName(sheetName);
    desTab.getRange(1, columnPosition, desTab.getMaxRows(), numberColumns).clearContent();
}