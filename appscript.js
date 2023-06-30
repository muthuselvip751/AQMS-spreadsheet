var server = 'mysql1006.mochahost.com';
var port = 3306;

var user = 'lanandan_idealab_database_google_js';
var userPwd = '5ylW6y5=Iogq';
var db = 'lanandan_idealab_database_google_js';

// var user = 'lanandan_idealab';
// var userPwd = 'Fq*vU;~jxZL7';
// var db = 'lanandan_idealab';



var dbUrl = 'jdbc:mysql://' + server + ':' + port + '/' + db;

var ui = SpreadsheetApp.getUi();

function onOpen() {
	ui.createMenu('Synchronization')
		.addItem('Upload Current Row', 'menuItem4')
		.addItem('Upload Entire Table', 'menuItem5')
		.addToUi();
}

function menuItem4() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheetNames = ["Customers", "Locations", "Branches", "Facilities", "Buildings", "Floors", "Lab_departments", "Category", "Devices", "sensorCategory", "Sensors"];
	var activeSheet = ss.getActiveSheet();
	var sheetIndex = sheetNames.indexOf(activeSheet.getName());

	if (sheetIndex !== -1) {
		var sheet = ss.getSheetByName(sheetNames[sheetIndex]);
		var sheetName = sheetNames[sheetIndex];
		var activeCell = sheet.getActiveCell();

		if (activeCell) {
			var row = activeCell.getRow();
			exportToSQL(row, false, sheet);
		} else {

		}
	}
}

function menuItem5() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var currentSheetName = ss.getActiveSheet();
	var lastRow = currentSheetName.getLastRow();
	console.log(lastRow);
	var colCount = currentSheetName.getLastColumn();
	console.log(colCount);
	for (var row = 2; row <= lastRow; row++) {
		exportToSQL(row, true, currentSheetName);
	}
}

function exportToSQL(row, isEntireTable, sheet) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	console.log({
		row: row,
		isEntireTable: isEntireTable,
		sheet: sheet
	})
	var tableMap = {
		"Customers": "customers",
		"Locations": "locations",
		"Branches": "branches",
		"Facilities": "facilities",
		"Buildings": "buildings",
		"Floors": "floors",
		"Lab_departments": "lab_departments",
		"Category": "categories",
		"Devices": "devices",
		"sensorCategory": "sensor_categories",
		"Sensors": "sensors"
	};
	var sheetName = sheet.getName();
	var actualSheet = ss.getSheetByName(sheetName);
	console.log('sheetName: ' + sheetName);
	console.log('actualSheet: ' + actualSheet);
	if (isEntireTable == true) {
		var rowCount = actualSheet.getLastRow();
		var colCount = actualSheet.getLastColumn();
		var sheetData = actualSheet.getRange(row - 1, 1, rowCount, colCount).getValues();
	} else {
		var rowCount = row;
		var colCount = actualSheet.getLastColumn();
		var sheetData = actualSheet.getRange(row, 1, 1, colCount).getValues();
	}
	for (var i = 1; i < sheetData[0].length; i++) {}

	var sqlColumns = '';
	var insertString = '';
	var updateString = '';
	var table = '';

	if (colCount > 0) {
		var sheetLabels = actualSheet.getRange(1, 1, 1, colCount).getValues()[0];
		var sheetData = actualSheet.getRange(row, 1, 1, colCount).getValues();

		console.log("heet Data " + sheetData)
	}

	var lastRow = actualSheet.getLastRow();

	var excludedColumns = ["id", "locationDropDown", "branchDropDown", "facilityDropDown", "buildingDropDown", "floorDropDown", "labDropDown", "categoryDropDown"];
	var RemoveBrackets = ["deviceCategory", "sensorCategoryName", "deviceName"];
	sqlColumns += ' companyCode';
	var sheetId = sheetLabels[0];
	var devicesCategoryIndex = 0;
	var deviceNameIndex = 0;
	var sensorCategoryNameIndex = 0;

	for (var col = 0; col < colCount; col++) {
		var columnName = sheetLabels[col];

		var keys = sheetData[col];

		if (!excludedColumns.includes(columnName)) {
			console.log("Column name " + columnName)
			if (columnName == "deviceCategory" || columnName == "deviceName" || columnName == "sensorCategoryName") {
				devicesCategoryIndex = col;
        sheetData[0][col] = sheetData[0][col].replace(/\(\d+\)/g, "");
			}
			// if (columnName == "deviceName") {
			// 	deviceNameIndex = col;
      //   sheetData[0][col] = sheetData[0][col].replace(/\(\d+\)/g, "");
			// }
			// if (columnName == "sensorCategoryName") {
			// 	sensorCategoryNameIndex = col;
      //   sheetData[0][col] = sheetData[0][col].replace(/\(\d+\)/g, "");
			// }

			sqlColumns += ',' + sheetLabels[col];
      

			updateString += sheetLabels[col] + '=' + "'" + sheetData[0][col] + "'" + ',';
		}
	}

	var timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss");

	sqlColumns += ', created_at, updated_at';
	updateString += ' updated_at=' + "'" + timestamp + "'" + '';

	var check = [];
	var inserId = [];
	var insertData = [];
	var datas = [];
	for (var i = 0; i < sheetData.length; i++) {

		check.push(sheetData[i]);
		var rowData = sheetData[i];
		var dataId = sheetData[i];


		var dataRow = [];
		var dataRow_sample = [];

		for (var col = 0; col < colCount; col++) {
			var columnName = sheetLabels[col];
			console.log("Column name " + columnName)
			if (columnName == "deviceCategory") {
				devicesCategoryIndex = col;
			}
			if (columnName == "deviceName") {
				deviceNameIndex = col;
			}
			if (columnName == "sensorCategoryName") {
				sensorCategoryNameIndex = col;
			}

		}
		console.log("############")
		console.log(devicesCategoryIndex)

		for (var col = 0; col < colCount; col++) {
			var columnName = sheetLabels[col];
			console.log(columnName)


			var values = dataId[col];
			datas.push(values);
			if (!excludedColumns.includes(columnName)) {
				var value = rowData[col];
				console.log("value " + value)
				console.log("index " + col)

				if (col == devicesCategoryIndex || col == deviceNameIndex || col == sensorCategoryNameIndex) {
					value = value.replace(/\(\d+\)/g, "");
					console.log("******************")
					console.log(value)
					console.log("******************")
				}


				if (typeof value === 'number') {
					dataRow.push(value);
					dataRow_sample.push(values);
				} else if (typeof value === 'string' && value.trim() !== '' && value !== "DELETE") {
					var checkErrorSqlSpcVal = value.toString();
					if (checkErrorSqlSpcVal.indexOf("'") !== -1) {
						checkErrorSqlSpcVal = checkErrorSqlSpcVal.replace("'", "''");
					}

					dataRow.push(checkErrorSqlSpcVal);
					dataRow_sample.push(checkErrorSqlSpcVal);
				} else {
					dataRow.push(null);
					dataRow_sample.push(null);
				}
			}
}
		var timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss");
		dataRow_sample.push(dataRow[0]);


		var customerSheet = ss.getSheetByName("Customers");
		var columnF = customerSheet.getRange("F2:F");
		var customervalues = columnF.getValues();
		console.log(customervalues[0][0]);

		dataRow.unshift(customervalues[0][0]);
		dataRow.push(timestamp, timestamp);
		insertData.push(dataRow);
	}
	console.log("insertData.length " + insertData.length)
	console.log(insertData)
	if (insertData.length > 0) {
		var insertStringValues = insertData.map(row => '(' + row.map(value => typeof value === 'string' ? "'" + value + "'" : value) + ')');
		insertString = sqlColumns + ') VALUES ' + insertStringValues;
		console.log("Insert string values..........")
		console.log(insertStringValues)
	}

	if (tableMap.hasOwnProperty(sheetName)) {
		table = tableMap[sheetName];
	}

	console.log("sheetData[0][0] " + sheetData[0][0])
	if (sheetData[0][0]) {
		console.log('update');
		var conditionValue = sheetData[0][0];
		var updateQuery = 'UPDATE ' + table + ' SET ' + updateString + ' WHERE id =' + conditionValue + ';';
		executeQuery(updateQuery);
		console.log("update query : " + updateQuery);
	} else {
		console.log('insert')
		console.log("insertString " + insertString)
		if (insertString.length > 0) {
			var conn = Jdbc.getConnection(dbUrl, user, userPwd);
			var stmt = conn.createStatement();
			var insertQuery = 'INSERT INTO ' + table + ' (' + insertString + ';';

			Logger.log("insert Query : " + insertQuery);

			executeQuery(insertQuery);

			var currentSheetName = SpreadsheetApp.getActiveSheet();
			var current = currentSheetName.getName();
			console.log("active" + current);
			var conn = Jdbc.getConnection(dbUrl, user, userPwd);
			var stmt = conn.createStatement();
			console.log("table " + table);
			var results = stmt.executeQuery('SELECT MAX(id) FROM ' + table);
			console.log("results " + results);

			if (results.next()) {
				console.log("I am running");
				var lastId = results.getString(1);
				lastId = parseInt(lastId);
				lastId += 1
				console.log("lastId " + lastId);


				var lastRow = currentSheetName.getLastRow();
				var values = currentSheetName.getRange(1, 1, lastRow).getValues();
				var emptyRowIndex = values.findIndex(function(row) {
					return row[0] === "";
				});

				var targetRow = emptyRowIndex !== -1 ? emptyRowIndex + 1 : lastRow + 1;

				currentSheetName.getRange(targetRow, 1).setValue(lastId - 1);
			}
		}
	}
}

function executeQuery(sqlQuery) {
	var start = new Date();
	var conn = Jdbc.getConnection(dbUrl, user, userPwd);
	var stmt = conn.createStatement();
	stmt.setMaxRows(1000);
	stmt.execute(sqlQuery, 1);
	stmt.close();
	conn.close();
	var end = new Date();
}

function getCurrentColumnHeaderName(activeCol) {
	var sh = SpreadsheetApp.getActiveSheet();
	var headerName = sh.getRange(1, activeCol).getValue();
	return headerName;
}

function updateData() {
	var activeCell = SpreadsheetApp.getActiveRange();
	var activeRow = activeCell.getRow();
	var activeCol = activeCell.getColumn();
	var activeValue = activeCell.getValue();
	var activeSheet = activeCell.getSheet();
	var activeSheetName = activeSheet.getName();
	var activeColumnHeader = getCurrentColumnHeaderName(activeCol);

	if (activeSheetName != "Locations") {
     var isUpdateAnotherRange = true;
		if (activeSheetName == "Branches") {
			if (activeColumnHeader === "locationDropDown") {
				isUpdateAnotherRange = false;
			}
		}
		if (activeSheetName == "Facilities") {
			if (activeColumnHeader === "branchDropDown") {
				isUpdateAnotherRange = false;
			}
		}
		if (activeSheetName == "Buildings") {
			if (activeColumnHeader === "facilityDropDown") {
				isUpdateAnotherRange = false;
			}
		}
		if (activeSheetName == "Floors") {
			if (activeColumnHeader === "buildingDropDown") {
				isUpdateAnotherRange = false;
			}
		}
		if (activeSheetName == "Lab_departments") {
			if (activeColumnHeader === "floorDropDown") {
				isUpdateAnotherRange = false;
			}                                                 
		}
		if (activeSheetName == "Devices") {
			if (activeColumnHeader === "deviceCategory") {
				isUpdateAnotherRange = false;
			}
		}
		updateRange(activeColumnHeader, activeValue, activeSheet, activeRow, activeCol, activeCell, isUpdateAnotherRange)
	}
}

function updateRange(activeColumnHeader, activeValue, activeSheet, activeRow, activeCol, activeCell, isUpdateAnotherRange = false) {
	var workSheet = SpreadsheetApp.getActiveSpreadsheet();



	switch (activeColumnHeader) {
		case "locationDropDown":
			Logger.log("***updateRange - locationDropDown ***");
			var locationsSheet = workSheet.getSheetByName('Locations');
			var data = locationsSheet.getDataRange().getValues();
			var stateNames = data.slice(1).map((row) => {
				var stateName = row[1];
				var locationId = row[0];
				if (stateName.trim() !== '') {
					return stateName + '(' + locationId + ')';
				}
				return '';
			}).filter(Boolean);
			var locationIds = data.slice(1).map((row) => {
				return row[0];
			});
			Logger.log(stateNames);

			var columnRange = activeSheet.getRange("C2:C");
			var validation = SpreadsheetApp.newDataValidation().requireValueInList(stateNames).setAllowInvalid(false).build();
			columnRange.setDataValidation(validation);

			var selectedIndex = stateNames.indexOf(activeValue);
			if (selectedIndex !== -1) {
				var locationId = locationIds[selectedIndex];
				activeSheet.getRange(activeRow, activeCol - 1).setValue(locationId);
			} else {
				activeSheet.getRange(activeRow, activeCol - 1);
			}
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var spreadSheet = workSheet.getSheetByName("Branches");
				var branchData = spreadSheet.getDataRange().getValues();
				var list = branchData.filter((row) => {
					return row[1] == currentRowValues[1] && row[2] == currentRowValues[2]
				}).map((row) => {
					return row[3] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;
		case "branchDropDown":
			Logger.log("***updateRange - branchDropDown ***");
			var spreadSheet = workSheet.getSheetByName('Branches');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Facilities');
				var branchData = currentSpreadSheet.getDataRange().getValues();
				var list = branchData.filter((row) => {
					return row[1] == currentRowValues[1] && row[2] == currentRowValues[2]
				}).map((row) => {
					return row[5] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;
		case "facilityDropDown":
			Logger.log("***updateRange - facilityDropDown ***");
			var spreadSheet = workSheet.getSheetByName('Facilities');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Buildings');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.filter((row) => {
					return row[1] == currentRowValues[1] && row[2] == currentRowValues[2]
				}).map((row) => {
					return row[7] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;
		case "buildingDropDown":
			Logger.log("***updateRange - buildingDropDown ***");
			var spreadSheet = workSheet.getSheetByName('Buildings');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Floors');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.filter((row) => {
					return row[1] == currentRowValues[1] && row[2] == currentRowValues[2]
				}).map((row) => {
					return row[10] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;
		case "floorDropDown":
			Logger.log("***updateRange - floorDropDown ***");
			var spreadSheet = workSheet.getSheetByName('Floors');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Lab_departments');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.filter((row) => {
					return row[1] == currentRowValues[1] && row[2] == currentRowValues[2]
				}).map((row) => {
					return row[11] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
				var activeSheetName = activeSheet.getName();
				if (activeSheetName === "Devices") {
					activeCell.offset(0, 3).setDataValidation(validation);
				}
				if (activeSheetName === "Sensors") {
					activeCell.offset(0, 2).setDataValidation(validation);
				}
			}
			break;
		case "labDropDown":
			Logger.log("***updateRange - labDropDown ***");
			var spreadSheet = workSheet.getSheetByName('Lab_departments');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Category');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.slice(1).map((row) => {
					return row[1] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
				var activeSheetName = activeSheet.getName();
				if (activeSheetName === "Devices") {
					activeCell.offset(0, 3).setDataValidation(validation);
				}
				if (activeSheetName === "Sensors") {
					activeCell.offset(0, 2).setDataValidation(validation);
				}
			}
			break;

		case "deviceCategory":
			Logger.log("***updateRange - deviceCategory ***");
			var spreadSheet = workSheet.getSheetByName('Category');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('sensorCategory');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.slice(1).map((row) => {
					return row[1] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
				var activeSheetName = activeSheet.getName();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;

		case "sensorCategoryName":
			Logger.log("***updateRange - sensorCategoryName ***");
			var spreadSheet = workSheet.getSheetByName('sensorCategory');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			if (isUpdateAnotherRange) {
				var rowIndex = activeSheet.getCurrentCell().getRow();
				var currentRowValues = activeSheet.getRange(rowIndex, 1, 1, activeSheet.getLastColumn()).getValues()[0];
				var currentSpreadSheet = workSheet.getSheetByName('Devices');
				var sheetData = currentSpreadSheet.getDataRange().getValues();
				var list = sheetData.slice(1).map((row) => {
					return row[14] + "(" + row[0] + ")";
				});
				Logger.log(list);
				var validation = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build();
				var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
				var activeSheetName = activeSheet.getName();
				activeCell.offset(0, 2).setDataValidation(validation);
			}
			break;

		case "deviceName":
			Logger.log("***updateRange - deviceName ***");
			var spreadSheet = workSheet.getSheetByName('Devices');
			var data = spreadSheet.getDataRange().getValues();
			var originalLocation = activeValue.split("(")[1].replace(")", "");
			activeSheet.getRange(activeRow, activeCol - 1).setValue(originalLocation);
			break;

	}

}


function onEdit() {
	updateData()
}