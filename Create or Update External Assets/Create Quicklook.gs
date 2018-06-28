var objSetTransp = {
	strSheetNameSRC: "Staff",
	strSheetNameDST: "Quick Look",
	strSheetNameTMPL: "Template",
}

function transpose() {

	var sp = SpreadsheetApp.getActiveSpreadsheet();
	var sh = sp.getSheetByName(objSetTransp.strSheetNameSRC)
		var shDst = sp.getSheetByName(objSetTransp.strSheetNameDST)
		var shTmpl = sp.getSheetByName(objSetTransp.strSheetNameTMPL)
		var data = sh.getDataRange().getValues();

	var arrDest = []
	var counter = 1;
	for (var i = 2; i < data.length; i++) {

		if (data[i][0] != "" && (('' + data[i][5]).trim().toUpperCase() == "FULL-TIME" || ('' + data[i][5]).trim().toUpperCase() == "PART-TIME" || ('' + data[i][5]).trim().toUpperCase() == "VOLUNTEER")) {
			var row = [counter, data[i][6], data[i][1] + ', ' + data[i][0], data[i][4], data[i][7], data[i][9], data[i][5]];
			arrDest.push(row);
			counter++;

		}

	}

	shDst.clear();
	var lr = shDst.getMaxRows();
    if(lr>1){
	shDst.deleteRows(2, lr - 1);
    }
	var rng = shTmpl.getRange(1, 1, 19, 9);
	rng.copyTo(shDst.getRange("A1"));
	var dates = shDst.getRange("A19").getValue()
		dates = ('' + dates).replace('{date_updated_in_MM_/_DD_/_YY}', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy"))
		shDst.getRange("A19").setValue(dates)

		//Logger.log(arrDest)

		if (arrDest.length > 1) {
			shDst.insertRows(5, (arrDest.length - 3));
		}

		shDst.getRange(4, 2, arrDest.length, arrDest[0].length).setValues(arrDest);

}