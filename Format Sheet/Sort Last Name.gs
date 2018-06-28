
var objS = {

	strSheetName: 'Staff',
	numStartRow: 3,

}

function sortByLastName() {

	var sp = SpreadsheetApp.getActiveSpreadsheet();
	var sh = sp.getSheetByName(objS.strSheetName);
	var lastRow = sh.getLastRow();
	var lastCol = sh.getLastColumn();

	var dupSH = sh.copyTo(sp);
	dupSH.getDataRange().clearFormat();

	dupSH.insertColumnBefore(1);
	var range = dupSH.getRange(1, 1, lastRow, 1);
	var vals = range.getValues();
	for (var i = 0; i < vals.length; i++) {

		for (var j = 0; j < vals[i].length; j++) {

			vals[i][j] = 1;

		}
	}
	range.setValues(vals);
	range.setNumberFormat("0.000");

	dupSH.insertColumnBefore(1);
	var range1 = dupSH.getRange(1, 1, lastRow, 1);
	var vals1 = range1.getFormulas();
	for (var i = 0; i < vals1.length; i++) {

		for (var j = 0; j < vals1[i].length; j++) {

			vals1[i][j] = "=(SUBTOTAL(102,offset(indirect(ADDRESS(ROW(),COLUMN())),-1,1,2)))";

		}
	}
	range1.setFormulas(vals1);
	range1.setNumberFormat("0.000");

	var valt = range1.getDisplayValues();

	var hr = [];
	var flag = false;
	for (var i = 0; i < valt.length; i++) {

		if (valt[i][0] == 1 && flag == false) {
			hr.push(i);
			flag = true;
		} else if (valt[i][0] == 0) {
			hr.push(i);
		} else if (valt[i][0] == 1) {
			flag = false;
		}

	}

	if (hr.length > 0) {
		for (var i = 0; i < hr.length; i++) {
			vals[hr[i]][0] = 0;
			var rg = dupSH.getRange(1 + hr[i], 1, 1, 1);
			dupSH.unhideRow(rg);

			var rg1 = sh.getRange(1 + hr[i], 1, 1, 1);
			sh.unhideRow(rg1);
		}
	}
	range.setValues(vals);

	var rangeS = dupSH.getRange(objS.numStartRow, 1, lastRow + 1 - objS.numStartRow, lastCol);
	rangeS.sort(4);

	var rangeS1 = sh.getRange(objS.numStartRow, 1, lastRow + 1 - objS.numStartRow, lastCol);
	rangeS1.sort(2);

	//Logger.log(valt);
	//Logger.log(hr);

	var vals = range.getValues();
	//Logger.log(vals)

	for (var i = 0; i < vals.length; i++) {
		if (vals[i][0] == 0) {
			var rg = sh.getRange(1 + i, 1, 1, 1);
			sh.hideRow(rg);

		}
	}
	sp.deleteSheet(dupSH);

}
