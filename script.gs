const COST_CENTER_SPREADSHEET_ID = '1pDfhyWhQaBtoMG1S0hlCUOBwrDKP9SHBedo2OB7YaaI';
const PEOPLE_DATA_SPREADSHEET_ID = '1SUyWn6yxf0of94wA_TWCb5zevG5yQo5fyEPb8SSA6ik';
const RESULT_ID = '1JTSi4mMq4l6tORFKkCxyVmOxyu7j69xiGSQNXY5l4ag'; // spreadsheet for the result of query

function doGet() {
  return HtmlService.createHtmlOutputFromFile('interface.html')
    .setTitle('Data Integration App');
}

function getGroups() {
  const spreadsheet = SpreadsheetApp.openById(COST_CENTER_SPREADSHEET_ID);
  const costCenterSheet = spreadsheet.getSheetByName('Cost_Centers_Data');
  const data = costCenterSheet.getDataRange().getValues();
  const headers = data[0];
  const groupIndex = headers.indexOf('Group');

  if (groupIndex === -1) {
    throw new Error('Column "Group" not found in headers: ' + headers);
  }

  const uniqueGroups = Array.from(new Set(data.slice(1).map(row => row[groupIndex]))).filter(Boolean);

  uniqueGroups.unshift("No specific group");

  return uniqueGroups;
}


function getCostCenters(group) {
  const spreadsheet = SpreadsheetApp.openById(COST_CENTER_SPREADSHEET_ID);
  const costCenterSheet = spreadsheet.getSheetByName('Cost_Centers_Data');
  const data = costCenterSheet.getDataRange().getValues();
  const headers = data[0];
  const groupIndex = headers.indexOf('Group');
  const costCenterIndex = headers.indexOf('Cost center');

  if (groupIndex === -1) {
    throw new Error('Column "Group" not found in headers.');
  }
  if (costCenterIndex === -1) {
    throw new Error('Column "Cost center" not found in headers.');
  }

  const costCenters = data
    .slice(1)
    .filter(row => group === "No specific group" || !group || row[groupIndex] === group)
    .map(row => row[costCenterIndex]);

  return Array.from(new Set(costCenters)).filter(Boolean);
}


function getColumns() {
  const costCenterSheet = SpreadsheetApp.openById(COST_CENTER_SPREADSHEET_ID)
    .getSheetByName('Cost_Centers_Data');
  const peopleSheet = SpreadsheetApp.openById(PEOPLE_DATA_SPREADSHEET_ID)
    .getSheetByName('Reporting_Finance_Data');

  const costCenterHeaders = costCenterSheet.getDataRange().getValues()[0];
  const peopleHeaders = peopleSheet.getDataRange().getValues()[0];

  return {
    costCenters: costCenterHeaders,
    people: peopleHeaders
  };
}

function getPeopleByCostCenter(selectedCostCenters) {
  const peopleSheet = SpreadsheetApp.openById(PEOPLE_DATA_SPREADSHEET_ID)
    .getSheetByName('Reporting_Finance_Data');
  const data = peopleSheet.getDataRange().getValues();
  const headers = data[0];
  const ccIndex = headers.indexOf('CC');

  if (ccIndex === -1) {
    throw new Error('Column "CC" not found in the headers.');
  }

  const people = data.slice(1).filter(row => selectedCostCenters.includes(row[ccIndex]));
  return { headers, people };
}

// creation of google spreadsheet
function generateFilteredReport(filters) {
  const { costCenters, selectedColumns } = filters;

  const peopleSheet = SpreadsheetApp.openById(PEOPLE_DATA_SPREADSHEET_ID)
    .getSheetByName('Reporting_Finance_Data');
  const data = peopleSheet.getDataRange().getValues();
  const headers = data[0];
  const ccIndex = headers.indexOf('CC');
  const fteIndex = headers.indexOf('FTE');
  const n1Index = headers.indexOf('N+1');

  if (ccIndex === -1) {
    throw new Error('Column "CC" not found in headers.');
  }
  if (fteIndex === -1) {
    throw new Error('Column "FTE" not found in headers.');
  }
  if (n1Index === -1) {
    throw new Error('Column "N+1" not found in headers.');
  }

  const filteredRows = data.slice(1).filter(row => costCenters.includes(row[ccIndex]));

  const columnIndices = selectedColumns.map(col => headers.indexOf(col)).filter(idx => idx !== -1);

  if (columnIndices.length === 0) {
    throw new Error('No valid columns selected for the report.');
  }

  const filteredData = filteredRows.map(row => columnIndices.map(idx => row[idx]));

  const totalFTE = filteredRows.reduce((sum, row) => {
    const fteValue = parseFloat(row[fteIndex].toString().replace(',', '.'));
    return sum + (isNaN(fteValue) ? 0 : fteValue);
  }, 0);

  const totalEmployees = filteredRows.length;

  const uniqueN1s = new Set(filteredRows.map(row => row[n1Index]).filter(Boolean)).size;

  // Create or access the Report sheet
  const reportSheet = SpreadsheetApp.openById(RESULT_ID)
    .getSheetByName('Filtered_Report') || SpreadsheetApp.openById(RESULT_ID).insertSheet('Filtered_Report');
  reportSheet.clear();

  const selectedHeaders = columnIndices.map(idx => headers[idx]);
  reportSheet.appendRow(selectedHeaders);
  filteredData.forEach(row => reportSheet.appendRow(row));

  reportSheet.appendRow(['', '']); // Spacer
  reportSheet.appendRow(['Summary']);
  reportSheet.appendRow(['Total FTE', totalFTE.toFixed(2)]);
  reportSheet.appendRow(['Total Employees', totalEmployees]);
  reportSheet.appendRow(['Total Unique N+1s', uniqueN1s]);

  return `Report generated successfully with ${totalEmployees} rows, ${selectedColumns.length} columns, and Total FTE: ${totalFTE.toFixed(2)}.`;
}

// download to excel format
function downloadAsXLSX() {
  try {
    const resultSpreadsheet = SpreadsheetApp.openById(RESULT_ID);
    const sheet = resultSpreadsheet.getSheetByName('Filtered_Report');

    if (!sheet) {
      throw new Error('Filtered_Report sheet not found.');
    }

    const exportUrl = `https://docs.google.com/spreadsheets/d/${RESULT_ID}/export?format=xlsx&sheet=${encodeURIComponent(sheet.getName())}`;
    
    return exportUrl;

  } catch (error) {
    console.error('Error in downloadAsXLSX:', error.toString());
    throw error;
  }
}


// email ready report
function createDesignatedSpreadsheet(managerEmail) {
  try {
    if (!managerEmail || !managerEmail.includes('@')) {
      throw new Error('Invalid manager email provided.');
    }

    const newSpreadsheet = SpreadsheetApp.create(`Report for ${managerEmail}`);

    const resultSpreadsheet = SpreadsheetApp.openById(RESULT_ID);
    const resultSheet = resultSpreadsheet.getSheetByName('Filtered_Report');

    if (!resultSheet) {
      throw new Error('Filtered_Report sheet does not exist in the source spreadsheet.');
    }

    const data = resultSheet.getDataRange().getValues();

    const newSheet = newSpreadsheet.getSheets()[0];
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    newSpreadsheet.addEditor(managerEmail);

    const fileUrl = newSpreadsheet.getUrl();
    const subject = `Your Designated Report is Ready`;
    const body = `Dear Manager,\n\nYour report has been created. You can access it using the link below:\n${fileUrl}\n\nBest regards,\nData Integration App`;
    GmailApp.sendEmail(managerEmail, subject, body);

    return `Spreadsheet created and emailed to ${managerEmail} successfully.`;
  } catch (error) {
    console.error('Error in createDesignatedSpreadsheet:', error.toString());
    throw error;
  }
}

// email ready report v2
function createDesignatedSpreadsheetAndExport(managerEmail) {
  try {
    if (!managerEmail || !managerEmail.includes('@')) {
      throw new Error('Invalid manager email provided.');
    }

    const newSpreadsheet = SpreadsheetApp.create(`Report for ${managerEmail}`);

    const resultSpreadsheet = SpreadsheetApp.openById(RESULT_ID);
    const resultSheet = resultSpreadsheet.getSheetByName('Filtered_Report');

    if (!resultSheet) {
      throw new Error('Filtered_Report sheet does not exist in the source spreadsheet.');
    }

    const data = resultSheet.getDataRange().getValues();

    const newSheet = newSpreadsheet.getSheets()[0];
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    newSpreadsheet.addEditor(managerEmail);

    const exportUrl = `https://docs.google.com/spreadsheets/d/${newSpreadsheet.getId()}/export?format=xlsx`;

    const fileUrl = newSpreadsheet.getUrl();
    const subject = `Your Designated Report is Ready`;
    const body = `Dear Manager,\n\nYour report has been created. You can access it using the link below:\n${fileUrl}\n\nTo download the report as an Excel file (.xlsx), click the link below:\n${exportUrl}\n\nBest regards,\nData Integration App`;

    GmailApp.sendEmail(managerEmail, subject, body);

    return `Spreadsheet created and emailed to ${managerEmail} successfully with XLSX download link.`;
  } catch (error) {
    console.error('Error in createDesignatedSpreadsheetAndExport:', error.toString());
    throw error;
  }
}

