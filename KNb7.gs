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
  const sheet = getSheet('設定_従業員');
  const lastRow = sheet.getLastRow();
  const lists = [];
  for (let i = 0; i < lastRow; i++) {
    lists[i] = [];
    lists[i].name = sheet.getRange(i + 2, 1).getValue();
    lists[i].basePay = sheet.getRange(i + 2, 2).getValue();
    lists[i].hourlyPay = sheet.getRange(i + 2, 3).getValue();
    lists[i].sheet = sheet.getRange(i + 2, 4).getValue();
  }
  return lists;
}

//ーーーーーーーーーーーーーーーーーーーーーーーーーーーーここから冨田が追加ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

function enterSalary() {
  const sheet1 = getSheet("支給額算出")
  const sheet3 = getSheet("設定_従業員")
  let lastRow = sheet1.getRange(sheet1.getMaxRows(), CELL_INFO.NAME).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();//最終行を取得
  for (let row = 2; row <= lastRow; row++) {
    executioy(row, sheet1, sheet3)
  }
}

function executioy(row, sheet1, sheet3) {
  overTimeEnter(row, sheet1)
  enterHourlyPay(row, sheet1, sheet3)
  enterOverTimePay(row, sheet1)
  enterBasicSalary(row, sheet1, sheet3)
  enterTotalSalary(row, sheet1)
}

function getScheduledWorkingHours(number) { //所定労働時間と残業時間を取得 引数
  let lists = getMemberLists()
  let URL = lists[number].sheet //スプレッドシートのURL
  let sheetName = getSheetName()
  if (isSheetExists(sheetName, URL)) {
    let times = []
    times.push(`=IMPORTRANGE("${URL}","${sheetName}!G3")`)
    times.push(`=IMPORTRANGE("${URL}","${sheetName}!H3")`)
    return times;
  }
  return null;
}

function isSheetExists(sheetName, url) { //シートが存在するかを確認する。
  let sheet = SpreadsheetApp.openByUrl(url)
  let allSheets = sheet.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (sheetName === allSheets[i].getName()) return true;
  }
  return false;
}

function getSheetName() { //userSheetのシートの名前を取得
  const sheet = getSheet("設定_算出用")
  let year = sheet.getRange(1, 3).getValue()
  let month = sheet.getRange(2, 3).getValue()
  let sheetName = `${year}/${month}`
  return sheetName;
}

function overTimeEnter(row, sheet) { //支出額算出に残業時間を記入する
  let times = getScheduledWorkingHours(row - 2)
  if (times) {
    let overTime = times[1]
    sheet.getRange(row, CELL_INFO.OVERTIME).setValue(overTime);
  } else {
    sheet.getRange(row, CELL_INFO.OVERTIME).setValue("0")
  }
}

function getHourlyWage(row, sheet) { //時給の値を取得する
  hourlyWage = sheet.getRange(row, CELL_INFO.HOURLYWAGE).getValue()
  return hourlyWage;
}

function enterOverTimePay(row, sheet) { //残業代を記入する
  let overTime = sheet.getRange(row, CELL_INFO.OVERTIME).getValue()
  if (overTime) {
    let hour = overTime.getHours();
    let minute = overTime.getMinutes();
    let hourlyWage = getHourlyWage(row, sheet)
    let timePay = (hour * hourlyWage + minute * hourlyWage / 60) * 1.25
    timePay = Math.round(timePay / 1) * 1 //四捨五入
    sheet.getRange(row, CELL_INFO.OVERTIMEPAY).setValue(timePay);
  } else {
    sheet.getRange(row, CELL_INFO.OVERTIMEPAY).setValue("0")
  }
}

function enterBasicSalary(row, sheet1, sheet3) { //基本給を記入する
  basicSalary = sheet3.getRange(row, CELL_INFO.BASICSALAY).getValue()
  sheet1.getRange(row, CELL_INFO.BASICSALAY).setValue(basicSalary)
}

function enterHourlyPay(row, sheet1, sheet3) { //時給を記入する
  basicSalary = sheet3.getRange(row, CELL_INFO.HOURLYWAGE).getValue()
  sheet1.getRange(row, CELL_INFO.HOURLYWAGE).setValue(basicSalary)
}

function enterTotalSalary(row, sheet) { //総支給額を記入する
  let basicSalary = sheet.getRange(row, CELL_INFO.BASICSALAY).getValue()
  let overTimePay = sheet.getRange(row, CELL_INFO.OVERTIMEPAY).getValue()
  if (!overTimePay) {
    overTimePay = "0"
  }
  totalSalary = `=SUM(${basicSalary}+${overTimePay})`;
  sheet.getRange(row, CELL_INFO.TOTALSARALY).setValue(totalSalary)
}