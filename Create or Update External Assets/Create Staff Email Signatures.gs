var objSet = {
	spreadsheetID: '1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI',
	sheetName: 'Staff',
	staffFolderID: '0BzM8_MjdRURAVWNPTm9IZWVoajA',
	destFolderID: '0BzM8_MjdRURANkYxLThqa3JsZGc',  
}

function gSES() {

	var filename = 'BubbleHead';
	var staffFolder = DriveApp.getFolderById(objSet.staffFolderID);
	var destFolder = DriveApp.getFolderById(objSet.destFolderID);
	var sp = SpreadsheetApp.openById(objSet.spreadsheetID);
	var sh = sp.getSheetByName(objSet.sheetName);
	var data = sh.getDataRange().getValues();

	for (var i = 2; i < data.length; i++) {

		if (data[i][0] != "" && (('' + data[i][5]).trim().toUpperCase() == "FULL-TIME")) {
			var first = data[i][0];
			var last = data[i][1];
			var jTitle = data[i][4];

			//Logger.log(last)

			var ownFolderS = staffFolder.getFoldersByName(last + ", " + first);

			if (ownFolderS.hasNext()) {
				var ownFolder = ownFolderS.next();

				var foldersI = ownFolder.getFolders();

				
                while (foldersI.hasNext()) {
					var folderI = foldersI.next();

					var files = folderI.searchFiles('title contains "' + filename + '"');
					if (files.hasNext()) {
						var file = files.next();
						var img = "https://drive.google.com/uc?export=view&id=" + file.getId();

						var fileBody = HtmlService.createHtmlOutputFromFile('Staff Email Template').getContent();

						fileBody = fileBody.replace('{firstName}', first);
						fileBody = fileBody.replace('{src}', img);
						fileBody = fileBody.replace('{fullName}', '' + first + ' ' + last);
						fileBody = fileBody.replace('{jobF}', jTitle);
                        
                        var filesD=destFolder.searchFiles('title contains "' + last + ", " + first + ".html" + '"');
                        
                        while (filesD.hasNext()) {
                         
                                   var fileD = filesD.next();
                                   fileD.setTrashed(true);
                        
                        }

						destFolder.createFile(last + ", " + first + '.html', fileBody)

						break;
					}

				}

			}

		}

	}

}

function doGet() {
	return HtmlService
	.createTemplateFromFile('template')
	.evaluate();
}