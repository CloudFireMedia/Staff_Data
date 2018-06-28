/* Redevelopment notes: 
*   
*   - trigger for 'on edit' does not work reliably (09.015.2017)
*   - should aphabetize only the sheets named after a team
*   - should delete all the sheets in the target spreadsheet that are named after a team first BEFORE adding the sheets again
*/

var doc_CCN_Staff_Data_Sheet='1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI';
var doc_CCN_Events_Promotion_Calendar='1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc';
var doc_TEMPLATE_CCN_Events_Promotion_Calendar='1Idf44phe12hFr6TRk-h8qDFTnu_BanOP0H0swnLAARw';
var sheet1_CCN_Staff_Data_Sheet='Staff';
var sheet1_TEMPLATE_CCN_Events_Promotion_Calendar='TEMPLATE';

function MaintainPromotionCalendar_sortSheets()
{
  var spreadsheet=SpreadsheetApp.openById(doc_CCN_Events_Promotion_Calendar);
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
}

function MaintainPromotionCalendar()
{
  var tsheet=SpreadsheetApp.openById(doc_TEMPLATE_CCN_Events_Promotion_Calendar).getSheetByName(sheet1_TEMPLATE_CCN_Events_Promotion_Calendar);
  var sheet=SpreadsheetApp.openById(doc_CCN_Staff_Data_Sheet).getSheetByName(sheet1_CCN_Staff_Data_Sheet);
  var lastRow=sheet.getLastRow();
  
  var dest=SpreadsheetApp.openById(doc_CCN_Events_Promotion_Calendar);
  
  var va=sheet.getSheetValues(3, 12, lastRow-2, 1);
  var va2=sheet.getSheetValues(3, 13, lastRow-2, 1);
  var vta=[];
  var vti=0;
  
  for (var row=0;row<va.length;row++)
  {
    var name=va[row][0];
    if (va2[row][0]=='Yes')
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
      }
      vta[vti]=name;
      vti++;
    }
  }
  for (var row=0;row<va.length;row++)
  {
    var name=va[row][0];
    var f=0;
    for (var vti=0; vti<vta.length; vti++)
    {
      if (vta[vti]==name) f++;
    }
    if (f==0)
    {
      var ds=dest.getSheetByName(name);
      if (ds!=null)
      {
        dest.deleteSheet(ds);
      }
    }
  }
  MaintainPromotionCalendar_sortSheets();
}