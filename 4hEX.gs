const CELL_INFO = {
  'DATE': 1
  , 'DAY': 2
  , 'START': 3
  , 'END': 4
  , 'WORKTIME': 5
  , 'REST': 6
  , 'REGULARTIME': 7
  , 'OVERTIME': 8
  , 'REMARK': 9
}
const USER_INFO = {
  'NAMECELL': 1
  , 'CALNEDERID': 2
}
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];

function onOpen() {
  let ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu("追加メニュー");
  menu.addItem("開始", "setStartTime");
  menu.addSeparator();
  menu.addItem("終了", "setEndTime");
  menu.addToUi();
}

function getMasterSheet() {
  let userSheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("勤怠管理");
  return [userSheet, sheet]
}

function getDate() {
  let date = new Date();
  return date
}

function getActiveSheet(sheetName) {
  let [userSheet, sheet1] = getMasterSheet()
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  if (!sheet)
    sheet = createSheet(sheetName, userSheet, sheet1);
  return sheet;
}

function createSheet(sheetName, userSheet, sheet1) {
  sheet = sheet1.copyTo(userSheet)
  sheet.setName(sheetName);
  return sheet;
}

function getDateInfo() {
  let date = getDate();
  let day = date.getDay();
  let dateInfo = {
    'date': date
    , 'day': day
    , 'workMonth': Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM')
    , 'workDate': Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd')
    , 'currnetTime': Utilities.formatDate(date, 'Asia/Tokyo', 'H:mm')
    , 'workDay': WEEK[day]
    , 'today': Number(Utilities.formatDate(date, 'Asia/Tokyo', 'dd'))
  };
  return dateInfo;
}

function setStartTime() {
  let dateInfo = getDateInfo();
  let sheet = getActiveSheet(dateInfo.workMonth);
  let lastRow = sheet.getLastRow();
  let preDate = null;
  for (let col = lastRow; col >= 3; col--) {
    if (sheet.getRange(col, CELL_INFO.DATE).isBlank() === false) {
      preDate = sheet.getRange(col, CELL_INFO.DATE).getValue().getDate()
      break;
    }
  }
  if (lastRow != 2 && sheet.getRange(lastRow, CELL_INFO.END).isBlank() === true) {
    let result = Browser.msgBox("終了ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    if (lastRow < 3) {
      sheet.getRange(lastRow + 1, CELL_INFO.DATE).setValue(dateInfo.workDate);
      sheet.getRange(lastRow + 1, CELL_INFO.DAY).setValue(dateInfo.workDay);
      sheet.getRange(lastRow + 2, CELL_INFO.START).setValue(dateInfo.currnetTime);
    } else if (preDate != dateInfo.today) {
      changeCellRed(sheet, lastRow)
      sheet.getRange(lastRow + 1, CELL_INFO.DATE).setValue(dateInfo.workDate);
      sheet.getRange(lastRow + 1, CELL_INFO.DAY).setValue(dateInfo.workDay);
      sheet.getRange(lastRow + 2, CELL_INFO.START).setValue(dateInfo.currnetTime);
    }
    else {
      sheet.getRange(lastRow + 1, CELL_INFO.START).setValue(dateInfo.currnetTime)
      calcDiffTime("D", "C", CELL_INFO.REST, lastRow, sheet)
    }
  }
}

function setEndTime() {
  let dateInfo = getDateInfo();
  let sheet = getActiveSheet(dateInfo.workMonth);
  let lastRow = sheet.getLastRow();
  if (sheet.getRange(lastRow, CELL_INFO.END).isBlank() === false && sheet.getRange(lastRow + 1, CELL_INFO.START).isBlank() === true) {
    let result = Browser.msgBox("開始ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    sheet.getRange(lastRow, CELL_INFO.END).setValue(dateInfo.currnetTime);
    calcDiffTime("C", "D", CELL_INFO.WORKTIME, lastRow, sheet)
    let cellNumber = checkDateRow(sheet, lastRow)
    calcWorkingTime(cellNumber, sheet)
    calcOverTime(cellNumber, sheet)
    restAlert(sheet, cellNumber)
    register(sheet, lastRow, dateInfo.date)
  }
}

function checkDateRow(sheet, lastRow) {
  let col = null;
  let cellNumber = null;
  for (col = lastRow; col >= 2; col--) {
    if (sheet.getRange(col, CELL_INFO.DATE).isBlank() == false) {
      cellNumber = col
      break;
    }
  }
  return cellNumber;
}

function calcDiffTime(before, after, low, lastRow, sheet) {
  let calcTime = "";
  let diffTime = "";
  let cellNumber = null;
  cellNumber = checkDateRow(sheet, lastRow)
  for (let cell = cellNumber + 1; cell <= lastRow; cell++) {
    let beforeTime = cell
    let afterTime = cell
    if (before == "D") {
      afterTime = cell + 1
    }
    if (cell != cellNumber + 1) {
      diffTime += ',';
    }
    diffTime += `MINUS(${after}${afterTime},${before}${beforeTime})`;
  }
  calcTime = `=SUM(${diffTime})`;
  sheet.getRange(cellNumber, low).setValue(calcTime)
}

function checkWorkTime(hour) {
  if (hour >= 8) {
    return 0
  } else if (hour >= 6) {
    return 1
  }
}

function checkRestTime(hour, minute) {
  if (hour <= 0 && minute < 45) {
    return 0
  } else if (hour <= 0 && minute <= 59) {
    return 1
  }
}

function restAlert(activeSheet, lastRow) {
  let sheet = activeSheet;
  let checkCol = checkDateRow(sheet, lastRow)
  let hour = sheet.getRange(checkCol, CELL_INFO.WORKTIME).getValue().getHours();
  if (checkWorkTime(hour) == 0) {
    Browser.msgBox("8時間以上連続勤務のため、1時間以上の休憩が必要です。");
  } else if (checkWorkTime(hour) == 1) {
    Browser.msgBox("6時間以上連続勤務のため、45分以上の休憩が必要です。");
  }
}

function changeCellRed(sheet, lastRow) {
  let workHour = null;
  let restHour = null;
  let restMinute = null;
  let col = checkDateRow(sheet, lastRow)
  if (!sheet.getRange(col, CELL_INFO.WORKTIME).isBlank()) {
    workHour = sheet.getRange(col, CELL_INFO.WORKTIME).getValue().getHours();
  }
  if (!sheet.getRange(col, CELL_INFO.REST).isBlank()) {
    restHour = sheet.getRange(col, CELL_INFO.REST).getValue().getHours();
    restMinute = sheet.getRange(col, CELL_INFO.REST).getValue().getMinutes();
  }
  if (checkWorkTime(workHour) == 0 && checkRestTime(restHour, restMinute) == 0
    || checkWorkTime(workHour) == 0 && checkRestTime(restHour, restMinute) == 1
    || checkWorkTime(workHour) == 1 && checkRestTime(restHour, restMinute) == 0) {
    sheet.getRange(col, CELL_INFO.REST).setBackground('red');
    sheet.getRange(col, CELL_INFO.REMARK).setValue('休憩が足りません');
  }
}

function getScheduledWorkingHours() {
  let URL = "管理側用スプレッドシートのID/edit#gid=0"
  let time = `IMPORTRANGE("${URL}","支給額算出!G2")`
  return time
}

function calcOverTime(col, sheet) {
  let time = getScheduledWorkingHours()
  let hour = sheet.getRange(col, CELL_INFO.WORKTIME).getValue().getHours();
  if (hour >= 8) {
    overTime = `=minus(E${col},${time})`
    sheet.getRange(col, CELL_INFO.OVERTIME).setValue(overTime);
  } else {
    sheet.getRange(col, CELL_INFO.OVERTIME).setValue(0);
  }
}

function calcWorkingTime(col, sheet) {
  let hour = sheet.getRange(col, CELL_INFO.WORKTIME).getValue().getHours();
  let time = getScheduledWorkingHours()
  if (hour >= 8) {
    sheet.getRange(col, CELL_INFO.REGULARTIME).setValue(`=${time}`);
  } else {
    sheet.getRange(col, CELL_INFO.REGULARTIME).setValue(sheet.getRange(col, CELL_INFO.WORKTIME).getValue());
  }
}

function register(sheet, lastRow, date) {
  let userSttingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  let account = userSttingSheet.getRange(2, USER_INFO.CALNEDERID).getValue();
  let calender = CalendarApp.getCalendarById(account);
  let startTime = sheet.getRange(lastRow, CELL_INFO.START).getValue();
  let endTime = sheet.getRange(lastRow, CELL_INFO.END).getValue();
  let title = userSttingSheet.getRange(2, USER_INFO.NAMECELL).getValue();
  let startDate = new Date(date);
  startDate.setHours(startTime.getHours());
  startDate.setMinutes(startTime.getMinutes());
  let endDate = new Date(date);
  endDate.setHours(endTime.getHours());
  endDate.setMinutes(endTime.getMinutes());
  calender.createEvent(title, startDate, endDate);
  Browser.msgBox("カレンダーに登録しました");
}
