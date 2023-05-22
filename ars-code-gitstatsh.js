//This is the ARS Engine [month] link
var sourceSheetLink = 'https://docs.google.com/spreadsheets/d/1otYwseHbr6DYcC544-S0AlZjKVo76jsXMOFJXs5LbSA/edit#gid=647400929';

function getDataForARSFAHW() {
    importData();
    getScheduleSV();
    getDistributionSV();
}

function importData() {
    removeDataAndSchedules('RD', 13);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('CMR Hours');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('C13').getValue();
    var data = sourceTab.getRange(1, 1, rows, 13).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet()
    var desSheet = thiSheet.getSheetByName('RD')
    var rango = desSheet.getRange(1, 1, rows, 13);
    rango.setValues(data)
}

function getScheduleSV() {
    removeDataAndSchedules('Schedules SV', 39);
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('Schedules SV');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('C14').getValue();
    var data = sourceTab.getRange(1, 1, rows, 39).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desSheet = thiSheet.getSheetByName('Schedules SV');
    var rango = desSheet.getRange(1, 1, rows, 39);
    rango.setValues(data)
}


function getDistributionSV() {
    removeDistribution('Schedules SV', 42, 33)
    var sourceSheet = SpreadsheetApp.openByUrl(sourceSheetLink);
    var sourceTab = sourceSheet.getSheetByName('Schedules SV');
    var commandTab = sourceSheet.getSheetByName('Command tab');
    var rows = commandTab.getRange('C14').getValue();
    var data = sourceTab.getRange(1, 117, rows, 33).getDisplayValues();
    var thiSheet = SpreadsheetApp.getActiveSpreadsheet();
    var desSheet = thiSheet.getSheetByName('Schedules SV');
    var rango = desSheet.getRange(1, 42, rows, 33);
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