let ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("勤怠管理");
let lastRow = sheet1.getLastRow();
let date = new Date();
let day = date.getDay();
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];
const DATECELL = 1
const DAYCELL = 2
const STARTCELL = 3
const ENDCELL = 4

function onOpen() {
  let ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu("追加メニュー");
  menu.addItem("開始", "setStartTime");
  menu.addSeparator();
  menu.addItem("終了", "setEndTime");
  menu.addToUi();
}

function set_sheet(sheet_name) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name)
  if (sheet)
    return sheet
  sheet = sheet1.copyTo(ss)
  sheet.setName(sheet_name);
  return sheet;
}

function setStartTime() {
  let workMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM');
  let sheet = set_sheet(workMonth);
  let lastRow = sheet.getLastRow();
  let workDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  let workStart = Utilities.formatDate(date, 'Asia/Tokyo', 'H:mm');
  let workDay = WEEK[day]
  let today = Number(Utilities.formatDate(date, 'Asia/Tokyo', 'dd'));
  let preDate = null;
  for (let i = lastRow; i >= 3; i--) {
    if (!sheet.getRange(i, DATECELL).isBlank()) {
      preDate = sheet.getRange(i, DATECELL).getValue().getDate()
      break;
    }
  }
  if (lastRow != 2 && sheet.getRange(lastRow, ENDCELL).isBlank()) {
    let result = Browser.msgBox("終了ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    if (lastRow < 3 || preDate != today) {
      sheet.getRange(lastRow + 1, DATECELL).setValue(workDate);
      sheet.getRange(lastRow + 1, DAYCELL).setValue(workDay);
      sheet.getRange(lastRow + 2, STARTCELL).setValue(workStart);
    } else {
      sheet.getRange(lastRow + 1, STARTCELL).setValue(workStart)
    }
  }
}

function setEndTime() {
  let workMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM');
  let sheet = set_sheet(workMonth);
  let lastRow = sheet.getLastRow();
  let workEnd = Utilities.formatDate(date, 'Asia/Tokyo', 'H:mm');
  if (!sheet.getRange(lastRow, ENDCELL).isBlank() && sheet.getRange(lastRow + 1, STARTCELL).isBlank()) {
    let result = Browser.msgBox("開始ボタンを押し忘れています", Browser.Buttons.OK);
    if (result == "ok") {
      return false
    }
  } else {
    sheet.getRange(lastRow, ENDCELL).setValue(workEnd);
  }
}