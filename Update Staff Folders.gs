// Redevelopment note: the trigger for 'on edit' for this project does not work reliably 

/**
 * Look through the staff data sheet:
 * 
 * - Create a new employee folder
 * - Archive any employees "No longer employed" or unarchive others
 */

function staffFolders_()
{
  var staffFolder = DriveApp.getFolderById(STAFF_FOLDER_ID);
  var archiveFolder = StaffFolders_find(staffFolder, 'Archive');
  
  var sheet = SpreadsheetApp.openById(STAFF_DATA_SHEET_ID).getSheetByName(STAFF_DATA_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var va = sheet.getSheetValues(3, 1, lastRow-2, 2); // First and last name 
  var va2 = sheet.getSheetValues(3, 6, lastRow-2, 1); // Status
  
  for (var row = 0; row < va.length; row++)
  {
    var folderName = va[row][1]+', '+va[row][0];
    var inStaffFolder;
    
    // Look in the main staff folder first
    var employeeFolder = StaffFolders_find(staffFolder, folderName);
    
    if (employeeFolder === null)
    {    
      // Next check the archive folder
      var employeeFolder = StaffFolders_find(archiveFolder, folderName);
      
      if (employeeFolder === null) 
      {  
        // In neither Archive or Staff so create in Staff
        employeeFolder = staffFolder.createFolder(folderName);
        inStaffFolder = true;
        Log_.info('Created folder "' + folderName + '"');        
      } 
      else 
      {
        inStaffFolder = false;
      }
      
    } 
    else 
    {
      inStaffFolder = true;
    }
    
    var status = va2[row][0];
    
    if (status === 'No longer employed')
    {
      if (inStaffFolder) {
        staffFolder.removeFolder(employeeFolder);
        archiveFolder.addFolder(employeeFolder);
        Log_.info('Archived folder "' + folderName + '"');
      }
    }
    else
    {
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
  
  function StaffFolders_find(parentFolder, folderName)
  {
    var childFolder = '';
    var foundFolder = '';
    
    var childFolders=parentFolder.getFoldersByName(folderName);
    while (childFolders.hasNext()) 
    {
      childFolder = childFolders.next();
      return childFolder;
    }
    
    childFolders = parentFolder.getFolders();
    while (childFolders.hasNext()) 
    {
      childFolder = childFolders.next();
      if (childFolder.getName()==' Archive') 
      {
        foundFolder=StaffFolders_find(childFolder,folderName);
        if (foundFolder!=false)
        {
          return foundFolder;
        }
      }
    }
    return null;
  }
  
} // staffFolders_()