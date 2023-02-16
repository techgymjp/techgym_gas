const CELL_INFO = {
  'DATE': 1
  , 'DAY': 2
  , 'START': 3
  , 'END': 4
  , 'WORKTIME': 5
  , 'REST': 6
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
    if (!sheet.getRange(col, CELL_INFO.DATE).isBlank()) {
      preDate = sheet.getRange(col, CELL_INFO.DATE).getValue().getDate()
      break;
    }
  }
  if (lastRow != 2 && sheet.getRange(lastRow, CELL_INFO.END).isBlank()) {
    let result = Browser.msgBox("終了ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    if (lastRow < 3 || preDate != dateInfo.today) {
      sheet.getRange(lastRow + 1, CELL_INFO.DATE).setValue(dateInfo.workDate);
      sheet.getRange(lastRow + 1, CELL_INFO.DAY).setValue(dateInfo.workDay);
      sheet.getRange(lastRow + 2, CELL_INFO.START).setValue(dateInfo.currnetTime);
    } else {
      sheet.getRange(lastRow + 1, CELL_INFO.START).setValue(dateInfo.currnetTime)
      calcDiffTime("D", "C", CELL_INFO.REST, lastRow, sheet)
    }
  }
}

function setEndTime() {
  let dateInfo = getDateInfo();
  let sheet = getActiveSheet(dateInfo.workMonth);
  let lastRow = sheet.getLastRow();
  if (!sheet.getRange(lastRow, CELL_INFO.END).isBlank() && sheet.getRange(lastRow + 1, CELL_INFO.START).isBlank()) {
    let result = Browser.msgBox("開始ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    sheet.getRange(lastRow, CELL_INFO.END).setValue(dateInfo.currnetTime);
    calcDiffTime("C", "D", CELL_INFO.WORKTIME, lastRow, sheet)
    restAlert(sheet, lastRow)
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
