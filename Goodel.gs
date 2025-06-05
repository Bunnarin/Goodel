function Goodel (sheetName) {
  return Goodel.Modeler(sheetName);
}
/*
 * Modeler takes the name of the sheet where the table is kept.
 * It expects the first row to be a header with column names
 * Ie. it won't search that row.
 */
Goodel.Modeler = function (sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = spreadsheet.getSheetByName(sheetName),
      table = new Goodel.Table(sheet),
      columns = table.getRow(1).getValues()[0];

  if (sheet == null) {
    var missingSheetMsg = 'Could not find sheet "<sheetName>" in active spreadsheet.'
                            .replace("<sheetName>", sheetName);
    throw new Error (missingSheetMsg);
  }

  function Model (instanceAttrHash) {
    this.model = Model;
    for (var attr in instanceAttrHash) this[attr] = instanceAttrHash[attr];
  }

  Model.sheet = sheet;
  Model.name = sheet.getName();
  Model.table = table;
  Model.columns = columns;

  for (var classMethod in Goodel._modelClassMethods) {
    Model[classMethod] = Goodel._modelClassMethods[classMethod];
  }

  Model.prototype = Object.create(Goodel._ModelInstance.prototype);
  
  return Model;
}


Goodel._modelClassMethods = function () {}

Goodel._modelClassMethods.all = function () {
  var allRecords = [];
  var sheet = this.sheet;
  var columns = this.columns; // Header columns

  // Get all values from the sheet, excluding the header row
  // Start from row 2 (index 1 in a 0-indexed array) up to the last row
  var dataRange = sheet.getRange(2, 1, this.table.numRows - 1, this.table.numColumns); //- 1 to exclude the header
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var rowData = values[i];
    var recordHash = {};
    for (let j = 0; j < columns.length; j++) recordHash[columns[j]] = rowData[j];
    allRecords.push(new this(recordHash)); // Create a new Model instance
  }
  return allRecords;
};

Goodel._modelClassMethods.findWhere = function (searchHash) {
  return this.table.findWhere(searchHash);
}

Goodel._modelClassMethods.findBy = function (searchHash) {
  var attrs = this.table.findBy(searchHash);

  return new this(attrs);
}

Goodel._modelClassMethods.findRowBy = function (searchHash) {
  return this.table.customManFindBy(searchHash);
}

Goodel._modelClassMethods.findRowWhere = function (searchHash) {
  return this.table.customManFindWhere(searchHash);
}

Goodel._modelClassMethods.getAllByColumn = function (columnName) {
  var columnIdx = this.table.columnMap[columnName];

  if (columnIdx === undefined) this.table._throwBadAttrMsg(columnName);

  var values = [];
  for (var rowIdx = 2; rowIdx <= this.table.numRows; rowIdx++) {
    values.push(this.table.getCell(rowIdx, columnIdx).getValue());
  }
  return values;
};

// Append this to your Goodel._modelClassMethods
Goodel._modelClassMethods.filterWhere = function (callback) {
  var searchResults = [];
  var allRecords = this.all(); // Use the 'all' method we just added

  for (var i = 0; i < allRecords.length; i++) {
    var record = allRecords[i];
    if (callback(record)) searchResults.push(record);
  }
  return searchResults;
};

// Append this to your Goodel._modelClassMethods
Goodel._modelClassMethods.filterBy = function (callback) {
  var allRecords = this.all();
  for (var i = 0; i < allRecords.length; i++) {
    var record = allRecords[i];
    if (callback(record)) return record; // Return the first match
  }
  return null; // No match found
};

Goodel._modelClassMethods.setColumnValues = function (columnName, value) {
  var columnIdx = this.table.columnMap[columnName];

  if (columnIdx === undefined) this.table._throwBadAttrMsg(columnName);

  // Get the range for the entire column, starting from the second row (after headers)
  // The number of rows will be the total number of rows in the table minus the header row.
  var columnRange = this.sheet.getRange(2, columnIdx, this.table.numRows - 1, 1);

  // Create an array of arrays with the desired value for each cell in the column
  var valuesToSet = [];
  for (var i = 0; i < this.table.numRows - 1; i++) valuesToSet.push([value]);

  // Set the values for the entire column
  columnRange.setValues(valuesToSet);
};

Goodel._modelClassMethods.updateWhere = function (searchHash, updateHash) {
  // Get the row indices that match the search criteria.
  // Assuming customManFindWhere now returns an array of row indices.
  var rowsIdxToUpdate = this.table.customManFindWhere(searchHash);

  if (!rowsIdxToUpdate || rowsIdxToUpdate.length === 0) {
    Logger.log("No records found matching the search criteria for update.");
    return;
  }

  var sheet = this.sheet;
  var columnMap = this.table.columnMap;

  // Iterate over each column in the updateHash
  for (var columnName in updateHash) {
    if (updateHash.hasOwnProperty(columnName)) {
      var newValue = updateHash[columnName];
      var columnIdxToSet = columnMap[columnName];

      if (columnIdxToSet === undefined) {
        // If a column in updateHash doesn't exist, throw an error
        this.table._throwBadAttrMsg(columnName);
      }

      // For each matching row, set the new value for the current column
      for (var i = 0; i < rowsIdxToUpdate.length; i++) {
        var rowIdx = rowsIdxToUpdate[i];
        sheet.getRange(rowIdx, columnIdxToSet).setValue(newValue);
      }
    }
  }
};

Goodel._modelClassMethods.updateBy = function (searchHash, updateHash) {
  // Use the existing findBy to get the first matching record as a hash
  var matchingRecord = this.findBy(searchHash); // This returns a Model instance
  
  if (!matchingRecord) {
    Logger.log("No record found matching the search criteria for updateBy.");
    return; // No record to update
  }

  const rowIdxToUpdate = this.table.customManFindBy(searchHash);

  if (rowIdxToUpdate === undefined || rowIdxToUpdate === null) {
    Logger.log("Could not determine row index for updateBy.");
    return;
  }

  var sheet = this.sheet;
  var columnMap = this.table.columnMap;

  // Iterate over each column in the updateHash and set its new value
  for (var columnName in updateHash) {
    if (updateHash.hasOwnProperty(columnName)) {
      var newValue = updateHash[columnName];
      var columnIdxToSet = columnMap[columnName];

      if (columnIdxToSet === undefined) {
        this.table._throwBadAttrMsg(columnName);
      }
      sheet.getRange(rowIdxToUpdate, columnIdxToSet).setValue(newValue);
    }
  }
};

Goodel._modelClassMethods.deleteWhere = function (searchHash) {
  // Use customManFindWhere which returns an array of row indices
  var rowsIdxToDelete = this.table.customManFindWhere(searchHash);

  if (!rowsIdxToDelete || rowsIdxToDelete.length === 0) {
    Logger.log("No records found matching the search criteria for deleteWhere.");
    return;
  }

  var sheet = this.sheet;

  // IMPORTANT: When deleting multiple rows, it's most efficient to delete them
  // from the highest row index to the lowest. This prevents the row indices
  // from shifting as you delete, which would cause incorrect deletions if
  // you delete from top down.
  rowsIdxToDelete.sort(function(a, b) { return b - a; }); // Sort descending

  for (var i = 0; i < rowsIdxToDelete.length; i++) {
    var rowIdx = rowsIdxToDelete[i];
    sheet.deleteRow(rowIdx);
  }

  // Update the table's row count by the number of deleted rows
  this.table.numRows -= rowsIdxToDelete.length;
};

Goodel._modelClassMethods.deleteBy = function (searchHash) {
  // Similar to updateBy, we need the row index to delete.
  const rowIdxToDelete = this.table.customManFindBy(searchHash);

  if (rowIdxToDelete === undefined || rowIdxToDelete === null) {
    Logger.log("No record found matching the search criteria for deleteBy.");
    return;
  }

  // Delete the entire row
  this.sheet.deleteRow(rowIdxToDelete);

  // Crucially, update the table's row count after deletion
  this.table.numRows--;
};

Goodel._modelClassMethods.create = function (recordHash) {
  var newRecord = [], len = this.columns.length, i;
  
  for (i = 0; i < len; i++) {
    var column = this.columns[i],
        attribute = recordHash[column] || "";
    
    newRecord.push(attribute);
  }
  var emptyRowIdx = this.table.getEmptyRowIdx(),
      row = this.table.getRow(emptyRowIdx);
  
  row.setValues([newRecord]);
  this.table.numRows++;

  return newRecord;
}

Goodel._modelClassMethods.toString = function () {
  return "{ sheet: <sheet>, columns: <columns> }"
          .replace("<sheet>", this.sheet.getName())
          .replace("<columns>", this.columns);
}


Goodel._ModelInstance = function () {}

Goodel._ModelInstance.prototype.save = function () {
  this.model.create(this);
}

Goodel._ModelInstance.prototype.toString = function () {
  var stringifiedObj = [];
  
  for (var key in this) {
    if (this.hasOwnProperty(key)) {
      var kvPair = "<key> = <value>"
        .replace("<key>", key)
        .replace("<value>", this[key]);
      
      stringifiedObj.push(kvPair);
    }
  }
  
  stringifiedObj = "{ " + stringifiedObj.join(", ") + " }";
  
  return stringifiedObj;
}


var ALPHABET = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H',
               'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
               'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

Goodel.Table = function (sheet) {
  this.sheet = sheet;
  this.numColumns = this.getEmptyColumnIdx() - 1;
  this.numRows = this.getEmptyRowIdx() - 1;
  this.columns = this.getRow(1).getValues()[0];
  this.columnMap = null;
  this._buildColumnMap();
}

Goodel.Table.prototype.findBy = function (searchHash) {
  /* If there are more than 50 cells to check,
   * use native search with single query.
   * Otherwise, make one query per row.
   */
  
  if (this.numRows * this.numColumns > 50 ) return this.natFindBy(searchHash);
  else return this.manFindBy(searchHash);
}

Goodel.Table.prototype.findWhere = function (searchHash) {
  // Same as above
  if (this.numRows * this.numColumns > 50) return this.natFindWhere(searchHash);
  else return this.manFindWhere(searchHash);
}

Goodel.Table.prototype.manFindBy = function (searchHash) {

  // Loop through rows
  for (var rowIdx = 1; rowIdx <= this.numRows; rowIdx++) {

    // Check match for every search key
    var isAMatch = true;
    for (var searchKey in searchHash) {
      var columnIdx = this.columnMap[searchKey];
 
      if (columnIdx == undefined) this._throwBadAttrMsg(searchKey);
 
      var attribute = this.getCell(rowIdx, columnIdx).getValue();

      if (searchHash[searchKey] != attribute) isAMatch = false;
    }

    // Return record when first match is found
    if (isAMatch) {
      var row = this.getRow(rowIdx).getValues()[0],
          record = this._hashifyRow(row);
      return record;
    }

  }
  // Return null if no matches are found
  return null
}

Goodel.Table.prototype.manFindWhere = function (searchHash) {
  var searchResults = [];
  
  // Loop through rows
  for (var rowIdx = 1; rowIdx <= this.numRows; rowIdx++) {
    var isAMatch = true;

    // Check match for every search key
    for (var searchKey in searchHash) {
      var columnIdx = this.columnMap[searchKey];

      if (columnIdx == undefined) this._throwBadAttrMsg(searchKey);

      var attribute = this.getCell(rowIdx, columnIdx).getValue();
      if (searchHash[searchKey] != attribute) isAMatch = false;
    }

    // Add record to results array if it's a match
    if (isAMatch) {
      var row = this.getRow(rowIdx).getValues()[0],
          record = this._hashifyRow(row);
      searchResults.push(record);
    }
  }

  return searchResults;
}


Goodel.Table.prototype.natFindBy = function (searchHash) {
  /*
   * Create a temporary sheet with a random name
   * where the search formula is inserted.
   * Then retrieves the search results and delete the sheet.
   */
  // Same as above
  var tempSheetName = Math.random().toString(36),
      spreadsheet = this.sheet.getParent(),
      tempSheet = spreadsheet.insertSheet(tempSheetName),
      firstCell = tempSheet.getRange(1,1);

  var searchFormula = this._getSearchRange(),
      searchConditions = this._buildSearchConditions(searchHash)
  
  
  searchFormula += searchConditions + ')';

  firstCell.setFormula(searchFormula);

  var firstCellValue = firstCell.getValue();
  if (firstCellValue == "#N/A") return null;
  if (firstCellValue == "#ERROR!") throw new Error("Query error");

  var matchingRow = tempSheet.getRange(1, 1, 1, this.numColumns).getValues()[0];
  var matchingRecord = this._hashifyRow(matchingRow);
  
  spreadsheet.deleteSheet(tempSheet);
  return matchingRecord;

}

Goodel.Table.prototype.natFindWhere = function (searchHash) {
  // Same as above
  var tempSheetName = Math.random().toString(36),
      spreadsheet = this.sheet.getParent();
      var tempSheet = spreadsheet.insertSheet(tempSheetName),
      firstCell = tempSheet.getRange(1,1),
      searchResults = [];
  
  var searchFormula = this._getSearchRange(),
      searchConditions = this._buildSearchConditions(searchHash)
  
  
  searchFormula += searchConditions + ')';

  firstCell.setFormula(searchFormula);

  var firstCellValue = firstCell.getValue();
  if (firstCellValue == "#N/A") return null;
  if (firstCellValue == "#ERROR!") throw new Error("Query error");

  // If one or more records were found,
  // loop through them and add them to the results array
  var i = 1;
  var thisRow = tempSheet.getRange(i, 1, 1, this.numColumns).getValues()[0];
  while (thisRow[0] != "") {
    var foundRecord = this._hashifyRow(thisRow);
    searchResults.push(foundRecord);
    thisRow = tempSheet.getRange(++i, 1, 1, this.numColumns).getValues()[0];
  }

  spreadsheet.deleteSheet(tempSheet);
  return searchResults;
}

Goodel.Table.prototype.customManFindBy = function (searchHash) {

  // Loop through rows
  for (var rowIdx = 1; rowIdx <= this.numRows; rowIdx++) {

    // Check match for every search key
    var isAMatch = true;
    for (var searchKey in searchHash) {
      var columnIdx = this.columnMap[searchKey];
 
      if (columnIdx == undefined) this._throwBadAttrMsg(searchKey);
 
      var attribute = this.getCell(rowIdx, columnIdx).getValue();

      if (searchHash[searchKey] != attribute) isAMatch = false;
    }

    // Return record when first match is found
    if (isAMatch) return rowIdx - 1; //translate it to 0-based

  }
  // Return null if no matches are found
  return;
}

Goodel.Table.prototype.customManFindWhere = function (searchHash) {
  var rowIdxResults = [];
  
  // Loop through rows
  for (var rowIdx = 1; rowIdx <= this.numRows; rowIdx++) {
    var isAMatch = true;

    // Check match for every search key
    for (var searchKey in searchHash) {
      var columnIdx = this.columnMap[searchKey];

      if (columnIdx == undefined) this._throwBadAttrMsg(searchKey);

      var attribute = this.getCell(rowIdx, columnIdx).getValue();
      if (searchHash[searchKey] != attribute) isAMatch = false;
    }

    // Add record to results array if it's a match
    if (isAMatch) rowIdxResults.push(rowIdx);
  }

  return rowIdxResults;
}

Goodel.Table.prototype._buildSearchConditions = function (searchHash) {
  var searchConditions = "";

  for (var key in searchHash) {
    var searchColumn = ALPHABET[this.columnMap[key] - 1];

    if (searchColumn === undefined) this._throwBadAttrMsg(key);

    var condition = '<table>!<searchCol>2:<searchCol><lastRow> = "<searchVal>"'
                      .replace("<table>", this.sheet.getName())
                      .replace(/<searchCol>/g, searchColumn)
                      .replace("<lastRow>", this.getEmptyRowIdx() - 1)
                      .replace("<searchVal>", searchHash[key]);


    searchConditions += condition + ',';
  }
  
  // Remove trailing ', '
  searchConditions = searchConditions.slice(0, searchConditions.length - 1);

  return searchConditions;
}

Goodel.Table.prototype._throwBadAttrMsg = function (key) {
  var badAttributeMsg = '<sheetName> table does not have a "<key>" column.'
                              .replace("<sheetName>", this.sheet.getName())
                              .replace("<key>", key);

  throw new Error(badAttributeMsg);
}

Goodel.Table.prototype._getSearchRange = function () {
  return "=FILTER(<table>!A2:<lastRow>, "
            .replace("<table>", this.sheet.getName())
            .replace("<lastRow>", this.getEmptyRowIdx() - 1);
}

Goodel.Table.prototype._hashifyRow = function (row) {
  var record = {}, len = this.numColumns, i;
  
  for (i = 0; i < len; i++) {

    const attr = this.columns[i];
    if (row[i] == "") record[attr] = null;
    else record[attr] = row[i];

  }
  return record;
}

Goodel.Table.prototype.getRow = function (row) {
  return this.sheet.getRange(row, 1, 1, this.numColumns);
}

Goodel.Table.prototype.getCell = function (row, col) {
  return this.sheet.getRange(row, col);
}

Goodel.Table.prototype.getRange = function (row, col, nRows, nCols) {
  return this.sheet.getRange(row, col, nRows, nCols);
}

Goodel.Table.prototype.getEmptyRowIdx = function () {
  let rowIdx = 1;

  while (this.getCell(rowIdx, 1).getValue() != "") rowIdx += 1;

  return rowIdx;
}

Goodel.Table.prototype.getEmptyColumnIdx = function () {
  let columnIdx = 1;

  while (this.getCell(1, columnIdx).getValue() != "") columnIdx++;
  
  return columnIdx;
}

Goodel.Table.prototype._buildColumnMap = function () {
  var columnMap = {},
      len = this.numColumns,
      columns = this.getRow(1).getValues()[0],
      i;
  
  for (i = 0; i < len; i++) columnMap[columns[i]] = i + 1;
  
  this.columnMap = columnMap;
}
