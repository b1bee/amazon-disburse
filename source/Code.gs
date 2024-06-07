const defaultColorMapSpreadsheetId = '15cp2wTfdq1hgqrtmJxKllqK-3fRBqfXxEttfOkzlXJM';
const colorMapSheetName = 'ColorMap';
const pivotSheetName = 'Summary';
const yellow = '#ffff00'; // sales
const blue = '#a4c2f4'; // shipping
const purple = '#d5a6bd'; // commission
const orange = '#ff9900'; // storage
const green = '#93c47d'; // subscription
const red = '#ff0000'; // reservers

var formulas = [];
var colorMapSpreadsheetId = '';

function onOpen() {
    SpreadsheetApp.getUi().createMenu('Automation')
        .addItem('Create Disbursement Summary', 'generateReport')
        .addToUi();
}

function generateReport(spreadsheetId) {
    if(!spreadsheetId){
        console.log('Using the default DisbursementColorMap spreadsheetId');
        colorMapSpreadsheetId = defaultColorMapSpreadsheetId;
    }else{
        console.log('Using spreadsheetId: ' + spreadsheetId);
        colorMapSpreadsheetId = spreadsheetId;
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if(!spreadsheet){
        console.error('This addon requires an Active Spreadsheet');
    }else{
        createPivotTable();
        setColorByCategory();
        addSummary();
    }
}

function createPivotTable(){
    console.log('createPivotTable...');

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // Graping the first Sheet
    const range = sheet.getDataRange();

    // Find the column index where "amount-description" occurs in the first row
    const firstRowValues = range.getValues()[0];
    const amountDescriptionIndex = firstRowValues.indexOf('amount-description') + 1; // Adding 1 to convert from 0-based to 1-based index

    // Find the column index where "amount" occurs in the first row
    const amountIndex = firstRowValues.indexOf('amount') + 1; // Adding 1 to convert from 0-based to 1-based index

    // Create a new sheet for the pivot table
    let pivotSheet = spreadsheet.getSheetByName(pivotSheetName); // Change to the name you want for your pivot table sheet
    if (!pivotSheet) {
        pivotSheet = spreadsheet.insertSheet(pivotSheetName);
    }

    // Create the pivot table
    var pivotTableRange = pivotSheet.getRange('A1'); // Change to the cell where you want your pivot table to start
    var pivotTable = pivotTableRange.createPivotTable(range);

    // Configure the pivot table
    pivotTable.addRowGroup(amountDescriptionIndex); // Group based on the found column index
    pivotTable.addPivotValue(amountIndex, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Summarize the next column assuming it's for "amount"
}

function setColorByCategory(){
    console.log('setColorsByCategory...');
    const colorMapSpreadsheet = SpreadsheetApp.openById(colorMapSpreadsheetId);
    const colorMapSourceSheet = colorMapSpreadsheet.getSheetByName(colorMapSheetName);

    const optionsSales = getColorMapValues(colorMapSourceSheet, 'A'); // sales
    const optionsShipping = getColorMapValues(colorMapSourceSheet, 'B'); // shipping
    const optionsCommission = getColorMapValues(colorMapSourceSheet, 'C'); // commission
    const optionsStorage = getColorMapValues(colorMapSourceSheet, 'D'); // storage
    const optionsSubscription = getColorMapValues(colorMapSourceSheet, 'E'); // subscription
    const optionsReservers = getColorMapValues(colorMapSourceSheet, 'F'); // reservers

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var pivotSheet = spreadsheet.getSheetByName(pivotSheetName);
    const lr = pivotSheet.getLastRow();

    const column = 1; // column A
    for (var row = 3; row <= lr; row++) {
        // Get the value from the specified cell
        var cell = pivotSheet.getRange(row, column);
        var value = pivotSheet.getRange(row, column).getValue();
        setColorIfMatch(row, value, cell, optionsSales, 0, yellow); // sales
        setColorIfMatch(row, value, cell, optionsShipping, 1, blue); // shipping
        setColorIfMatch(row, value, cell, optionsCommission, 2, purple); // commission
        setColorIfMatch(row, value, cell, optionsStorage, 3, orange); // storage
        setColorIfMatch(row, value, cell, optionsSubscription, 4, green); // subscription
        setColorIfMatch(row, value, cell, optionsReservers, 5, red); // reservers
    }

    // console.log(formulas);
}

function addSummary(){
    console.log('addSummary...');

    var pivotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pivotSheetName);
    var cell = null;
    // Set the value and background color for the cell

    cell = pivotSheet.getRange('D3');
    cell.setValue('Sales');
    cell.setBackground(yellow);
    cell = pivotSheet.getRange('E3');
    cell.setValue(formulas[0]);

    cell = pivotSheet.getRange('D4');
    cell.setValue('Shipping Service');
    cell.setBackground(blue);
    cell = pivotSheet.getRange('E4');
    cell.setValue(formulas[1]);

    cell = pivotSheet.getRange('D5');
    cell.setValue('Commission');
    cell.setBackground(purple);
    cell = pivotSheet.getRange('E5');
    cell.setValue(formulas[2]);

    cell = pivotSheet.getRange('D6');
    cell.setValue('Storage');
    cell.setBackground(orange);
    cell = pivotSheet.getRange('E6');
    cell.setValue(formulas[3]);

    cell = pivotSheet.getRange('D7');
    cell.setValue('Subscription');
    cell.setBackground(green);
    cell = pivotSheet.getRange('E7');
    cell.setValue(formulas[4]);

    cell = pivotSheet.getRange('D9');
    cell.setValue('Total');
    cell = pivotSheet.getRange('E9');
    cell.setValue("=sum(E3:E7)");

    cell = pivotSheet.getRange('D10');
    cell.setValue('Reservers');
    cell.setBackground(red);
    cell = pivotSheet.getRange('E10');
    cell.setValue(formulas[5]);

    pivotSheet.setColumnWidth(1, 300); //A
    pivotSheet.setColumnWidth(2, 100); //B
    pivotSheet.setColumnWidth(4, 200); //D
    pivotSheet.setColumnWidth(5, 100); //E
}

function setColorIfMatch(row, value, cell, options, fIndex, color){
    if (options.includes(value)) {
        cell.setBackground(color);

        var formula = formulas[fIndex];
        formulas[fIndex] = !formula? "=B"+row : formula + "+B"+row;
    }
}

function getColorMapValues(colorMapSourceSheet, column){
    const lr = colorMapSourceSheet.getLastRow();
    const range = colorMapSourceSheet.getRange(column +'2:' + column + lr); // position 1 is the header

    // Get the values in the column
    var values = range.getValues();
    values = getNonEmptyValues(values);
    return values;
}

function getNonEmptyValues(values){
    // Create an array to hold non-empty values
    let nonEmptyValues = [];

    // Loop through the values and add non-empty values to the array
    for (var i = 0; i < values.length; i++) {
        if (values[i][0] !== '') {
            nonEmptyValues.push(values[i][0]);
        }
    }
    return nonEmptyValues;
}
