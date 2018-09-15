function TEST_init() {
  Log_ = BBLog.getLog({
    level:                BBLog.Level.ALL, 
    displayFunctionNames: BBLog.DisplayFunctionNames.NO,
    sheetId:              STAFF_DATA_SHEET_ID_,
  })  
}

function TEST_misc() {
  var parentFolder = DriveApp.getFolderById('1jGF-Md5vsJHP41FuUZ8XWtcSHKvRn87i');
  var childFolders = parentFolder.getFoldersByName('Archive');
  var archiveFolder
  if (childFolders.hasNext()) {
    archiveFolder = childFolders.next()
    if (childFolders.hasNext()) {
      throw 'has two'
    }
  } else {
    throw 'no folders'
  }
}

function TEST_maintainPromotionCalendar() {
  TEST_init()
  maintainPromotionCalendar_()
}

function TEST_transpose() {
  TEST_init()
  transpose_(STAFF_DATA_SHEET_ID_)
}
