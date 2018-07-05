function TEST_init() {
  Log_ = BBLog.getLog({
    level:                BBLog.Level.ALL, 
    displayFunctionNames: BBLog.DisplayFunctionNames.YES,
    sheetId:              STAFF_DATA_SHEET_ID_,
  })  
}

function TEST_maintainPromotionCalendar() {
  TEST_init()
  maintainPromotionCalendar_()
}

function TEST_transpose() {
  TEST_init()
  transpose_(STAFF_DATA_SHEET_ID_)
}
