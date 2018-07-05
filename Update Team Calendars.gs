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
  var tsheet=SpreadsheetApp.openById(EVENT_TEMPLATE_ID_).getSheetByName(TEMPLATE_SHEET_NAME_);
  var staffDataSheet=SpreadsheetApp.openById(STAFF_DATA_SHEET_ID_).getSheetByName(STAFF_DATA_SHEET_NAME_);
  var lastRow=staffDataSheet.getLastRow();
  
  var dest=SpreadsheetApp.openById(CCN_EVENTS_PROMOTION_CALENDAR_ID_);
  
  var staffTeams=staffDataSheet.getSheetValues(3, 12, lastRow-2, 1); // Team
  var teamLeaders=staffDataSheet.getSheetValues(3, 13, lastRow-2, 1); // Team Leader
  var teamsWithLeader=[];
  var teamsWithLeaderIndex=0;
  
  // For each team create a new Event Sponsorship tab if one doesn't exist, 
  // and store the name of this team
 
  for (var row=0;row<staffTeams.length;row++)
  {
    var name=staffTeams[row][0];
    
    if (teamLeaders[row][0]=='Yes')
    {         
      var ds=dest.getSheetByName(name);
      
      if (ds==null)
      {
        ds=tsheet.copyTo(dest);
        ds.setName(name);
        var a1=ds.getRange(1,1);
        a1.setValue(a1.getValue().toString().replace(/TEMPLATE/g,name));
        var a3=ds.getRange(3,1);
        a3.setFormula(a3.getFormula().toString().replace(/TEMPLATE/g,name));
        Log_.info('Created "' + name + '" tab in promotions calendar')        
      }
      
      teamsWithLeader[teamsWithLeaderIndex]=name;
      teamsWithLeaderIndex++;
    }
  }
  
  // Do a second pass through the list of staff teams and if that team does not have a 
  // leaders, then delete the tab
  
  for (var row=0;row<staffTeams.length;row++)
  {
    var name=staffTeams[row][0];
    
    if (name === 'N/A') {
      continue;
    }
    
    var teamHasLeader=0;
    
    for (teamsWithLeaderIndex=0; teamsWithLeaderIndex<teamsWithLeader.length; teamsWithLeaderIndex++)
    {
      if (teamsWithLeader[teamsWithLeaderIndex]==name) {
        teamHasLeader++;
      }
    }
    
    if (teamHasLeader==0)
    {
      var ds=dest.getSheetByName(name);
      if (ds!=null)
      {
        dest.deleteSheet(ds);
        Log_.info('Deleted "' + name + '" tab from promotions calendar')        
      }
    }
  }
  
  sortSheets();
  
  return
  
  // Private Functions
  // -----------------
  
  function sortSheets()
  {
    var spreadsheet=SpreadsheetApp.openById(CCN_EVENTS_PROMOTION_CALENDAR_ID_);
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
        //spreadsheet.setActiveSheet(sheeta[sic]);
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