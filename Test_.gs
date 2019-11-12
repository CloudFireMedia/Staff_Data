function TEST_init() {
  Log_ = BBLog.getLog({
    level:                BBLog.Level.ALL, 
    displayFunctionNames: BBLog.DisplayFunctionNames.NO,
    sheetId:              TEST_STAFF_DATA_SHEET_ID_,
  })  
}

function TEST_getSheets() {

  var promotionsSpreadsheet = SpreadsheetApp.openById('1jKUuqzmz_DM4Xsd2eXHjPpb7Yo0XLFWK9-7kTgbhcDQ')
  var a = getSheets()

  function getSheets() {  
  
    var ignoreSheets = Config.get('STAFF_DATA_IGNORE_LIST').split(',')  
   
    var sheets = promotionsSpreadsheet.getSheets().filter(function(sheet) {
      return !ignoreSheets.some(function(sheetName) {        
        if (sheet.getName().trim() === sheetName.trim()) {
          return true
        }
      })       
    })

    return sheets
    
  } // maintainPromotionCalendar_.getSheets()
  
  return 
}

function TEST_misc() {
  Logger.log(PropertiesService.getDocumentProperties().getProperties())
  return 
}

function TEST_maintainPromotionCalendar() {
  TEST_init()
  maintainPromotionCalendar_()
}

function TEST_staffFolders() {
  TEST_init()
  staffFolders_()
}