/**
 * Maintain Promotion Calendar
 */

function maintainPromotionCalendar_() {

  var staffDataSpreadsheet = SpreadsheetApp.openById(STAFF_DATA_SHEET_ID_);

  var staffDataSheet = staffDataSpreadsheet
    .getSheetByName(STAFF_DATA_SHEET_NAME_);
    
  var lastRow = staffDataSheet.getLastRow();
  
  var promotionsSpreadsheet = SpreadsheetApp.openById(EVENTS_PROMOTION_CALENDAR_ID_);
  
  var jobTitles      = staffDataSheet.getSheetValues(3, 5,  lastRow-2, 1); 
  var statuss        = staffDataSheet.getSheetValues(3, 6,  lastRow-2, 1); 
  var emails         = staffDataSheet.getSheetValues(3, 9,  lastRow-2, 1);
  var staffTeamNames = staffDataSheet.getSheetValues(3, 12, lastRow-2, 1); 
  var teamLeaders    = staffDataSheet.getSheetValues(3, 13, lastRow-2, 1); 
  
  var teams = {};
  var communicationsDirectorEmail = null;
  
  // Create an ES tab
  // ----------------
  //
  // For each team create a new Event Sponsorship tab if one doesn't exist and
  // it has a team leader
 
  for (var rowIndex = 0; rowIndex < staffTeamNames.length; rowIndex++) {
  
    var nextJobTitle = jobTitles[rowIndex][0];
    Log_.fine('nextJobTitle: ' + nextJobTitle);
  
    if (nextJobTitle === 'Communications Director') {
    
      if (communicationsDirectorEmail !== null) {
        throw new Error('There are two communications directors in the Staff Data sheet');
      }
    
      communicationsDirectorEmail = emails[rowIndex][0];
    }
  
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
    
    if (nextTeamLeader === 'Yes' && nextEmployed) {
    
      if (nextTeam.hasOwnProperty('teamLeader') && nextTeam.teamLeader === 'Yes') {
        throw new Error('There are two team leaders assigned to ' + nextTeamName)
      }
    
      nextTeam.teamLeader = true;
      
      var ds = promotionsSpreadsheet.getSheetByName(nextTeamName);
      
      if (ds === null) {
      
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
    
      if (!nextTeam.hasOwnProperty('teamLeader') && !nextTeam.hasOwnProperty('nonTeamLeader')) {
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
  
    if (!teams.hasOwnProperty(sheetName)) {
    
      // There are no staff member in this team
      deleteTab = true;
      
    } else {       
    
      var nextTeam = teams[sheetName];
      
      if (nextTeam.hasOwnProperty('teamLeader')) {
      
        if (!nextTeam.teamLeader) {
        
          // This tab has no team leader, any members are ignored
          deleteTab = true;
        }
      }    
    }
    
    Log_.fine('deleteTab: ' + deleteTab);
    
    if (deleteTab) {

      sendDeleteNotification();

      if (TEST_DELETE_ES_TAB) {
        promotionsSpreadsheet.deleteSheet(sheet);
      } else {
        Log_.warning('Delete ES tab disabled');
      }
      
      Log_.info('Deleted "' + sheetName + '" tab from promotions calendar');                
    }
    
    // Private Functions
    // -----------------
    
    /**
     * Send a notification about the deleted ES tab
     */
     
    function sendDeleteNotification() {
    
      Log_.fine('communicationsDirectorEmail: ' + communicationsDirectorEmail);
    
      if (communicationsDirectorEmail === null) {
        throw new Error('There is no Communications Director in the Staff Data sheet');
      }
  
      if (communicationsDirectorEmail === '') {
        throw new Error('There is no email for the Communications Director in the Staff Data sheet');
      }
  
      var subject = Utilities.formatString(NOTIFICATION_SUBJECT, sheetName);
      var textBody = '';
      
      var template = HtmlService.createTemplateFromFile('Notification_')
      template.leaderlessTeamName = sheetName;
      
      var options = {
        htmlBody: template.evaluate().getContent(),
      }
    
      MailApp.sendEmail(
        communicationsDirectorEmail, 
        subject, 
        textBody, 
        options);
        
      Log_.info('Sent "ES Tab deleted" notification to ' + communicationsDirectorEmail);
   
    } // maintainPromotionCalendar_.sendDeleteNotification()
      
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
  
    allSheets.forEach(function(sheet) {
    
      if (ignore(sheet)) {
        return;
      }
    
      sheets.push(sheet);
      
      return;
      
      // Private Functions
      // -----------------
      
      function ignore(sheet) {
      
        var ignoreSheet = false;
    
        SPECIAL_CALENDAR_SHEET_NAMES.forEach(function(specialSheetName) {
        
          if (sheet.getName() === specialSheetName) {
            ignoreSheet = true;
          }
        })
        
        return ignoreSheet;
      }
    })
    
    return sheets;
    
  } // maintainPromotionCalendar_.getSheets()

  function sortSheets() {
  
    var spreadsheet=SpreadsheetApp.openById(EVENTS_PROMOTION_CALENDAR_ID_);
    var sheeta=spreadsheet.getSheets();
    var sic=0;
    
    for (var si=2;si<sheeta.length;si++) {
      sic=si;
      
      var ss1=sheeta[si].getName();
      
      for (var si2=si+1;si2<sheeta.length;si2++) {
      
        if (sheeta[sic].getName().localeCompare(sheeta[si2].getName())>0) {
          var s1s=sheeta[sic].getName();
          var s2s=sheeta[si2].getName();
          sic=si2;
        }
      }
      
      if (sic!=si) {
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