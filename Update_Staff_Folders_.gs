/**
 * Look through the staff data sheet:
 * 
 * - Create a new employee folder
 * - Archive any employees "No longer employed" or unarchive others
 */

function staffFolders_() {

  var staffFolderId = Config.get('STAFF_FOLDER_ID');
  
  var staffFolder = DriveApp.getFolderById(staffFolderId);
  
  if (staffFolder === null) {
    throw new Error('Could not find Staff folder')
  }
    
  var archiveFolder = staffFoldersFind(staffFolder, 'Archive');
  
  if (archiveFolder === null) {
    throw new Error('Could not find Archive folder in ' + staffFolderId)
  }
  
  var staffDataSheetId = Config.get('STAFF_DATA_SHEET_ID');
  var sheet = SpreadsheetApp.openById(staffDataSheetId).getSheetByName(STAFF_DATA_SHEET_NAME_);
  var lastRow = sheet.getLastRow();
  var va = sheet.getSheetValues(3, 1, lastRow-2, 2); // First and last name 
  var va2 = sheet.getSheetValues(3, 6, lastRow-2, 1); // Status
  
  for (var row = 0; row < va.length; row++) {
  
    var folderName = va[row][1]+', '+va[row][0];
    var inStaffFolder;
    
    // Look in the main staff folder first
    var employeeFolder = staffFoldersFind(staffFolder, folderName);
    
    if (employeeFolder === null) { 
    
      // Next check the archive folder
      var employeeFolder = staffFoldersFind(archiveFolder, folderName);
      
      if (employeeFolder === null) {
      
        // In neither Archive or Staff so create in Staff
        employeeFolder = staffFolder.createFolder(folderName);
        inStaffFolder = true;
        Log_.info('Created folder "' + folderName + '"');        
        
      } else {
        inStaffFolder = false;
      }
      
    } else {
      inStaffFolder = true;
    }
    
    var status = va2[row][0];
    
    if (status === 'No longer employed') {
      if (inStaffFolder) {
        staffFolder.removeFolder(employeeFolder);
        archiveFolder.addFolder(employeeFolder);
        Log_.info('Archived folder "' + folderName + '"');
      }
    } else {
      if (!inStaffFolder) {
        staffFolder.addFolder(employeeFolder);
        archiveFolder.removeFolder(employeeFolder);
         Log_.info('Move "' + folderName + '" out of Archive folder ');       
      }
    }
  }
  
  return 
  
  // Private Functions
  // -----------------
  
  // Get the folder name or null
  
  function staffFoldersFind(parentFolder, folderName) {
  
    Log_.fine('parentFolder: ' + parentFolder);
    Log_.fine('folderName: ' + folderName);
  
    var childFolder = '';
    var foundFolder = '';
    
    // Look in the top level of the parent
    
    var childFolders = parentFolder.getFoldersByName(folderName);
    
    if (childFolders.hasNext()) {    
    
      childFolder = childFolders.next();
      
      if (childFolders.hasNext()) {
        throw new Error('There are multiple folders called "' + folderName + '"');
      }
      
      Log_.fine('Found "' + folderName + '" by its name')
      return childFolder;
    } 
     
    // Look for the "Archive" folder and look in there
     
    childFolders = parentFolder.getFolders();
    
    while (childFolders.hasNext()) {
    
      childFolder = childFolders.next();
      
      if (childFolder.getName() === 'Archive') {
      
        foundFolder = staffFoldersFind(childFolder, folderName);
        
        if (foundFolder) {
          Log_.fine('Found "' + folderName + '" in Archive folder');
          return foundFolder;
        }
      }
    }
    
    Log_.fine('Failed to find folder')
    return null;
     
  } // staffFolders_.staffFoldersFind()
  
} // staffFolders_()