const CELL_INFO = {
  'NAME': 1
  , 'BASICSALAY': 2
  , 'HOURLYWAGE': 3
  , 'OVERTIME': 4
  , 'OVERTIMEPAY': 5
  , 'TOTALSARALY': 6
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('追加メニュー');
  menu.addItem('支給額算出', 'enterSalary');
  menu.addToUi();

  setMonthList();
}

function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function setMonthList() {
  const sheet = getSheet('設定_算出用');
  const values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  const cell = sheet.getRange('C2');
  cell.setDataValidation(rule);
}

function getMemberLists() {
  const sheet = getSheet('支給額算出'); //const sheet = getSheet('設定_従業員');
  const lastRow = sheet.getLastRow();
  const lists = [];
  for (let i = 0; i < lastRow; i++) {
    lists[i] = [];
    lists[i].name = sheet.getRange(i + 2, 1).getValue();
    lists[i].basePay = sheet.getRange(i + 2, 2).getValue();
    lists[i].hourlyPay = sheet.getRange(i + 2, 3).getValue();
    lists[i].sheet = sheet.getRange(i + 2, 8).getValue();
  }
  return lists;
}

function enterSalary() {
  const sheet1 = getSheet("支給額算出")
  let lastRow = sheet1.getRange(sheet1.getMaxRows(), CELL_INFO.NAME).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  for (let row = 2; row <= lastRow; row++) {
    executioy(row, sheet1)
  }
}

function executioy(row, sheet1) {
  overTimeEnter(row, sheet1)
  enterOverTimePay(row, sheet1)
  enterTotalSalary(row, sheet1)
}

function getScheduledWorkingHours(number) {
  let lists = getMemberLists()
  let URL = lists[number].sheet
  let sheetName = getSheetName()
  if (isSheetExists(sheetName, URL)) {
    let time =`=IMPORTRANGE("${URL}","${sheetName}!H2")`
    return time;
  }
  return null;
}

function isSheetExists(sheetName, url) {
  let sheet = SpreadsheetApp.openByUrl(url)
  let allSheets = sheet.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (sheetName === allSheets[i].getName()) return true;
  }
  return false;
}

function getSheetName() {
  const sheet = getSheet("設定_算出用")
  let year = sheet.getRange(1, 3).getValue()
  let month = sheet.getRange(2, 3).getValue()
  let sheetName = `${year}/${month}`
  return sheetName;
}

function overTimeEnter(row, sheet) {
  let time = getScheduledWorkingHours(row - 2)
  if (time) {
    let overTime = time
    sheet.getRange(row, CELL_INFO.OVERTIME).setValue(overTime);
  } else {
    sheet.getRange(row, CELL_INFO.OVERTIME).setValue("0")
  }
}

function getHourlyWage(row, sheet) {
  hourlyWage = sheet.getRange(row, CELL_INFO.HOURLYWAGE).getValue()
  return hourlyWage;
}

function enterOverTimePay(row, sheet) {
  let overTime = sheet.getRange(row, CELL_INFO.OVERTIME).getValue()
  if (overTime) {
    let hour = overTime.getHours();
    let minute = overTime.getMinutes();
    let hourlyWage = getHourlyWage(row, sheet)
    let timePay = (hour * hourlyWage + minute * hourlyWage / 60) * 1.25
    timePay = Math.round(timePay / 1) * 1
    sheet.getRange(row, CELL_INFO.OVERTIMEPAY).setValue(timePay);
  } else {
    sheet.getRange(row, CELL_INFO.OVERTIMEPAY).setValue("0")
  }
}

function enterTotalSalary(row, sheet) {
  let basicSalary = sheet.getRange(row, CELL_INFO.BASICSALAY).getValue()
  let overTimePay = sheet.getRange(row, CELL_INFO.OVERTIMEPAY).getValue()
  if (!overTimePay) {
    overTimePay = "0"
  }
  totalSalary = `=SUM(${basicSalary}+${overTimePay})`;
  sheet.getRange(row, CELL_INFO.TOTALSARALY).setValue(totalSalary)
}
