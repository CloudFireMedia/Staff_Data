////Hide Unemployed.gs
//
//var objSettings = {
//	numColumn : 6,
//	strString : "No longer employed",
//	strSheetName : "Staff",
//}
//
//function onEdit(e) {
//
//	var range = e.range;
//	var strValue = e.value;
//	var oldStrValue = e.oldValue;
//	var sh = range.getSheet();
//
//	if (strValue != oldStrValue) {
//		if (sh.getName() == objSettings.strSheetName) {
//
//			if (strValue == objSettings.strString) {
//				if (range.getColumn() == objSettings.numColumn) {
//					//Logger.log(oldStrValue)
//					//Logger.log(strValue)
//
//					sh.hideRows(range.getRow(), 1);
//
//				}
//			}
//            
//          transpose();
//            
//		}
//
//	}
//}