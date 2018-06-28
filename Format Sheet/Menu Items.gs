function onOpen(e){
  SpreadsheetApp.getUi().createMenu("[ Push Updates ]")
  .addItem("Update Staff Folders in Google Drive", "StaffFolders")
  .addItem("Update Event Sponsorship Pages for Teams", "MaintainPromotionCalendar")
  .addToUi();
}
