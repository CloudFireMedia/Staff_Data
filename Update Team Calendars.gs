/* Redevelopment notes: 
 *   
 *   - trigger for 'on edit' does not work reliably (09.015.2017)
 *   - should aphabetize only the sheets named after a team
 *   - should delete all the sheets in the target spreadsheet that are named after a team first BEFORE adding the sheets again
 */

/**
 * Maintain Promotion Calendar
 */

function maintainPromotionCalendar_()
{
  var staffDataSheet = SpreadsheetApp
    .openById(STAFF_DATA_SHEET_ID_)
    .getSheetByName(STAFF_DATA_SHEET_NAME_)
    
  var lastRow = staffDataSheet.getLastRow();
  
  var promotionsSpreadsheet = SpreadsheetApp.openById(EVENTS_PROMOTION_CALENDAR_ID_);
  
  var statuss = staffDataSheet.getSheetValues(3, 6, lastRow-2, 1); // Status
  var staffTeamNames = staffDataSheet.getSheetValues(3, 12, lastRow-2, 1); // Team
  var teamLeaders = staffDataSheet.getSheetValues(3, 13, lastRow-2, 1); // Team Leader
  var teams = {};
  
  // Create an ES tab
  // ----------------
  //
  // For each team create a new Event Sponsorship tab if one doesn't exist and
  // it has a team leader
 
  for (var rowIndex = 0; rowIndex < staffTeamNames.length; rowIndex++)
  {
    var nextTeamName = staffTeamNames[rowIndex][0];
    
    if (nextTeamName === 'N/A') {
      continue;
    }

    var nextTeamLeader = teamLeaders[rowIndex][0];
    var nextEmployed = employed(statuss[rowIndex][0]);

    if (!teams.hasOwnProperty(nextTeamName)) {
      teams[nextTeamName] = {};
    }
    
    var nextTeam = teams[nextTeamName];
    
    if (nextTeamLeader === 'Yes' && nextEmployed)
    {    
      if (nextTeam.hasOwnProperty('teamLeader') && nextTeam.teamLeader === 'Yes') 
      {
        throw new Error('There are two team leaders assigned to ' + nextTeamName)
      }
    
      nextTeam.teamLeader = true;
      
      var ds = promotionsSpreadsheet.getSheetByName(nextTeamName);
      
      if (ds === null)
      {     
        ds = SpreadsheetApp
          .openById(EVENT_TEMPLATE_ID_)
          .getSheetByName(TEMPLATE_SHEET_NAME_)
          .copyTo(promotionsSpreadsheet)
          .setName(nextTeamName);
        
        var a1 = ds.getRange(1,1); // ES Tab title
        a1.setValue(a1.getValue().toString().replace(/TEMPLATE/g,nextTeamName));
        
        var a3 = ds.getRange(3,1); // Query formula
        a3.setFormula(a3.getFormula().toString().replace(/TEMPLATE/g,nextTeamName));
        
        teams[nextTeamName].tabId = ds.getSheetId();
        
        Log_.info('Created "' + nextTeamName + '" tab in promotions calendar');      
      }      
      
    } else {
    
      if (!nextTeam.hasOwnProperty('teamLeader') && !nextTeam.hasOwnProperty('nonTeamLeader')) 
      {
        nextTeam.teamLeader = false;
        nextTeam.nonTeamLeader = true;
      }
    }
    
  } // for each row
  
  // Delete redundant ES tabs
  // ------------------------
  //
  // Do a second pass through the list of staff teams and if that team does not have a 
  // leader or any members, then delete the tab
  
  var sheets = getSheets();
  Log_.fine('Number of sheets: ' + sheets.length);
  
  sheets.forEach(function(sheet) {
  
    var sheetName = sheet.getName();
    Log_.fine('sheetName: ' + sheetName);
    var deleteTab = false;
  
    if (!teams.hasOwnProperty(sheetName)) 
    {
      // There are no staff member in this team
      deleteTab = true;
    }
    else 
    {       
      var nextTeam = teams[sheetName];
      
      if (nextTeam.hasOwnProperty('teamLeader')) {
      
        if (!nextTeam.teamLeader && !nextTeam.nonTeamLeader) {
        
          // This tab has no team leader and no members
          deleteTab = true;
        }
      }    
    }
    
    if (deleteTab) 
    {
      promotionsSpreadsheet.deleteSheet(sheet);
      Log_.info('Deleted "' + sheetName + '" tab from promotions calendar');                
    }
  })
  
  // Sort the sheets
  // ---------------
  
  sortSheets();
  
  return;
  
  // Private Functions
  // -----------------

  function employed(status) {
  
    return (status === 'Full-time' || status === 'Part-time' || status === 'Volunteer') ? true : false;
    
  } // maintainPromotionCalendar_.employed()

  function getSheets() {
  
    var allSheets = promotionsSpreadsheet.getSheets();
    var sheets = [];
  
    allSheets.forEach(function(sheet) 
    {
      if (ignore(sheet)) {
        return;
      }
    
      sheets.push(sheet);
      
      return;
      
      // Private Functions
      // -----------------
      
      function ignore(sheet) 
      {      
        var ignoreSheet = false;
    
        SPECIAL_CALENDAR_SHEET_NAMES.forEach(function(specialSheetName) 
        {    
          if (sheet.getName() === specialSheetName) 
          {
            ignoreSheet = true;
          }
        })
        
        return ignoreSheet;
      }
    })
    
    return sheets;
    
  } // maintainPromotionCalendar_.getSheets()

  function sortSheets()
  {
    var spreadsheet=SpreadsheetApp.openById(EVENTS_PROMOTION_CALENDAR_ID_);
    var sheeta=spreadsheet.getSheets();
    var sic=0;
    
    for (var si=2;si<sheeta.length;si++)
    {
      sic=si;
      
      var ss1=sheeta[si].getName();
      
      for (var si2=si+1;si2<sheeta.length;si2++)
      {
        if (sheeta[sic].getName().localeCompare(sheeta[si2].getName())>0)
        {
          var s1s=sheeta[sic].getName();
          var s2s=sheeta[si2].getName();
          sic=si2;
        }
      }
      
      if (sic!=si)
      {
        var sin1=sheeta[si].getIndex();
        var sin2=sheeta[sic].getIndex();
        var sis1=sheeta[si].getName();
        var sis2=sheeta[sic].getName();
        sheeta[sic].activate();
        spreadsheet.moveActiveSheet(si+1);
        sheeta=spreadsheet.getSheets();
        var stest1=sheeta[0].getName();
        var stest2=sheeta[1].getName();
        var stest3=sheeta[2].getName();
        var stest4=sheeta[3].getName();
      }
    }
    
  } // maintainPromotionCalendar_.sortSheets()
    
} // maintainPromotionCalendar_()