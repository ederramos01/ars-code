//This is the ARS Engine [month] link
var sourceSheetLink = 'https://docs.google.com/spreadsheets/d/1otYwseHbr6DYcC544-S0AlZjKVo76jsXMOFJXs5LbSA/edit#gid=647400929';

function getDataForARSFAHW() {
    importData();
    getSchedule('Schedules SV', 39);
    getDistribution('Schedules SV', 116, 41, 33);
    getDistribution('Schedules SV', 149, 74, 33);
    getSchedule('Schedules GT', 39);
    getDistribution('Schedules GT', 116, 41, 33);
    getDistribution('Schedules GT', 149, 74, 33);
}

function importData() {
    removeDataAndSchedules('RD', 16);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('CMR Hours');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('C13').getValue();
    var data = sourceTab.getRange(1, 1, rows, 16).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet()
    var desSheet = thiSheet.getSheetByName('RD')
    var rango = desSheet.getRange(1, 1, rows, 16);
    rango.setValues(data)
}

function getSchedule(tabName, columnRange) {
    removeDataAndSchedules(tabName, columnRange);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName(tabName);
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('C14').getValue();
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
    var rows = commandTab.getRange('C14').getValue();
    var data = sourceTab.getRange(1, targetColumnPosition, rows, numberColumns).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desSheet = thiSheet.getSheetByName(tabName);
    var rango = desSheet.getRange(1, columnPosition, rows, numberColumns);
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