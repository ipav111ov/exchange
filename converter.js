// Constants for headers, tax rate, API URL, and event details
const HEADERS = [
  ["Business ID", "", "Phone", "", "", "", "", ""],
  ["Date", "EUR", "Exchange", "GEL", "Cumulative Total GEL", "Tax(1%)", "Text for bank", "Text for bank GEO"]
];

const TAX_RATE = 0.01; // Tax rate (1%)
function my() {
  const URL = "https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/ka/json/?currencies=EUR&date=";
  getJson(URL)
}
const EVENT = {
  eventType: "ON_EDIT",
  handlerFunction: "onEditTrigger"
};

const SHEETNAME = "Exchange Sheet";

// Dictionary for sheet settings
const DICT_SHEET = {
  header: {
    firstRowNum: 1,
    numberColumns: 6,
    indexRow: 0,
    indexColumnHeaderId: 0,
    indexColumnId: 1,
    textColumnHeaderId: `Enter Your Business ID`,
  },
  rowStartData: 3,
  columnNumber: 8,
  columnIndexDate: 0,
  columnIndexAmount: 1,
  columnTotal: 5,
  columnId: 2,
  dateColumn: 1,
  amountColumn: 2,
  endValidationRow: 1000,
  validationNumberstart: 100000000, // count of digits not amount - it is 9 digits
  validationNumberend: 999999999,

};

// Function to create a custom menu in Google Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Work with Converter')
    .addItem('3.Create Exchange Sheet', 'createExchangeSheet')
    .addItem('2.Installation', 'checkIfTriggerExistsAndCreate')
    .addItem('1.Authorization', 'checkIfTriggerExistsAndCreate')
    .addToUi()
};

// Function to create a new exchange sheet
function createExchangeSheet() {
  const sheet = addSheet(SHEETNAME)
  sheet.getRange(1, 1, HEADERS.length, HEADERS[0].length).setValues(HEADERS);
  setWidthAndValidations(sheet)
}

// Function to add a new sheet with optional parameters
function addSheet(shName) {
  const ss = SpreadsheetApp.getActive();
  let n = 0;
  let nShName = shName;
  while (ss.getSheetByName(nShName) != null) {
    n += 1;
    nShName = shName + ' (' + n + ')';
  }
  shName = nShName;
  return ss.insertSheet(shName);
}

// Function to set column widths and data validations
function setWidthAndValidations(sheet) {
  sheet.getRange(DICT_SHEET.rowStartData, DICT_SHEET.dateColumn, DICT_SHEET.endValidationRow, DICT_SHEET.dateColumn)
    .setNumberFormat('dd"."mm"."yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .setHelpText('Enter a valid date')
      .requireDate()
      .build());
  sheet.getRange(DICT_SHEET.rowStartData, DICT_SHEET.amountColumn, DICT_SHEET.endValidationRow,)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .requireNumberGreaterThan(0)
      .setHelpText('Enter amount')
      .build());
  sheet.getRange(DICT_SHEET.header.firstRowNum, DICT_SHEET.columnId)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .setHelpText("Enter 9-digit number")
      .requireNumberBetween(DICT_SHEET.validationNumberstart, DICT_SHEET.validationNumberend)
      .build());

}

// Function to check if the trigger exists and create it if not
function checkIfTriggerExistsAndCreate() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExists = triggers.reduce((flag, trigger) => {
    if (trigger.getEventType() == EVENT.eventType && trigger.getHandlerFunction() == EVENT.handlerFunction) {
      flag = true;
    }
    return flag;
  }, false);
  if (triggerExists) {
    Browser.msgBox('Already done');
  } else {
    ScriptApp.newTrigger(EVENT.handlerFunction)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
    Browser.msgBox('Сompleted');
  }

}

// Function triggered on edit of the spreadsheet
function onEditTrigger(e) {
  const changeRange = e.range;
  const column = changeRange.getColumn();
  if (column === DICT_SHEET.dateColumn || column === DICT_SHEET.amountColumn) {
    const rowIndex = changeRange.getRow();
    const sheet = changeRange.getSheet();
    const allSheetRange = sheet.getDataRange();
    const values = allSheetRange.getValues();
    const valuesRow = values[rowIndex - 1];
    const firstRow = values[0];
    const date = valuesRow[DICT_SHEET.columnIndexDate];
    const amount = valuesRow[DICT_SHEET.columnIndexAmount];
    if (date !== "" && amount !== "") {
      const monthNumber = date.getMonth() + 1;
      const dateNumber = date.getDate() + 1;
      const yearNumber = date.getFullYear();
      const dateStr = `${yearNumber}-${monthNumber}-${dateNumber}`;
      const monthStr = date.toLocaleString('default', { month: 'long' });
      const fullUrl = URL + dateStr;
      const id = getId(sheet, firstRow);
      const textForBank = getTextForBank(id, monthStr, yearNumber);
      const json = getJson(fullUrl);
      const rate = json[0].currencies[0].rate;
      const lari = Number((amount * rate).toFixed(2));
      const tax = Number((lari * TAX_RATE).toFixed(2));
      let cumulativeTotal = lari;
      if (rowIndex === DICT_SHEET.rowStartData) {
        cumulativeTotal = lari;
      } else {
        const cumTotalYearRange = Number(values[rowIndex - 2][DICT_SHEET.columnTotal - 1]); // values[rowIndex - 2] = минус 2 потому что нужна предыдущая строка
        cumulativeTotal = cumTotalYearRange + lari;
      }
      updateSheet(sheet, rowIndex, values, dateStr, amount, rate, lari, cumulativeTotal, tax, textForBank[0], textForBank[1]);
    }
  }
}

// Function to get the business ID, prompting the user if not available
function getId(sheet, row) {
  let id = row[DICT_SHEET.header.indexColumnId];
  if (id === '') {
    const inputData = Browser.inputBox(`Your 9-digit Business ID`, Browser.Buttons.OK_CANCEL);
    if (inputData != 'cancel' || inputData != '') {
      id = inputData;
      row[DICT_SHEET.header.indexColumnId] = inputData;
      sheet.getRange(DICT_SHEET.header.firstRowNum, DICT_SHEET.columnId).setValue(id);
    }
  }
  return id
}

// Function to generate text for the bank based on business ID, month, and year
function getTextForBank(id, month, year) {
  const text = `${id}, ${month} ${year}, Small Business Tax`;
  const textGEO = LanguageApp.translate(text, `en`, `ka`);
  const texts = [text, textGEO];
  return texts
}

// Function to fetch JSON data from the given URL
function getJson(url) {
  const response = UrlFetchApp.fetch(url);
  const content = response.getContentText();
  let json;
  if (response.getResponseCode() == 200) {
    json = JSON.parse(content);
  } else {
    Browser.msgBox("Request failed. Response code: " + response.getResponseCode());
    Browser.msgBox("Response content: " + content);
  }
  return json;
}

// Function to put array in the sheet with calculated values
function updateSheet(sheet, rowIndex, values, dateStr, amount, rate, lari, cumulativeTotal, tax, textForBank, textForBankGeo) {
  values[rowIndex - 1] = [dateStr, amount, rate, lari, cumulativeTotal, tax, textForBank, textForBankGeo];
  sheet.getRange(rowIndex, 1, 1, DICT_SHEET.columnNumber).setValues([values[rowIndex - 1]]);
}
