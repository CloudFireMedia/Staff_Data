// Redevelopment note: the trigger for 'on edit' for this project does not work reliably 

var folder_Staff='0BzM8_MjdRURAVWNPTm9IZWVoajA';
var doc_CCN_Staff_Data_Sheet='1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI';
var sheet1_CCN_Staff_Data_Sheet='Staff';

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
  return false;
}

function StaffFolders()
{
  var parentFolder=DriveApp.getFolderById(folder_Staff);
  var archiveFolder=StaffFolders_find(parentFolder,' Archive');
  
  var sheet=SpreadsheetApp.openById(doc_CCN_Staff_Data_Sheet).getSheetByName(sheet1_CCN_Staff_Data_Sheet);
  var lastRow=sheet.getLastRow();
  var va=sheet.getSheetValues(3, 1, lastRow-2, 2);
  var va2=sheet.getSheetValues(3, 6, lastRow-2, 1);
  
  for (var row=0;row<va.length;row++)
  {
    var folderName=va[row][1]+', '+va[row][0];
    var found=StaffFolders_find(parentFolder,folderName);
    if (found==false)
    {
      found=parentFolder.createFolder(folderName);
    }
    
    var fvalue=va2[row][0];
    if (va2[row][0]=='No longer employed')
    {
      parentFolder.removeFolder(found);
      archiveFolder.addFolder(found);
    }
    else
    {
      parentFolder.addFolder(found);
      archiveFolder.removeFolder(found);
    }
  }
}