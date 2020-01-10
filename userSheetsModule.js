/** 
 * Module for handling interactions between Spreadsheets and Drive. Includes the 
      following functions:
 * createNewSpreadsheets For each string in a list, creates a new Spreadsheet 
      with that title in the given Drive folder ID.
 * getFilesInFolder Retrieves an Array of file IDs for a Drive folder with the 
      given ID.
 * getSpreadsheetsInFolder Retrieves a list of Spreadsheets for a Drive folder 
      with the given ID.
 * spreadsheetExistsInFolder Checks if a spreadsheet with a given name exists in 
      a folder.
 */ 
var userSheetsModule_ = (function(){    
  /**
    * For each string in a list, creates a new Spreadsheet with that title in 
        the given Drive folder ID, and initializes with all of the sheets
        from the master document. Defaults to applying no permissions at this
        point.
    * @param {String} driveFolderID The ID for the Drive folder where 
        Spreadsheets will be created.
    * @param {String/Array<String>} nameList String or array of strings 
        containing the name(s) of each spreadsheet.
    * @param {String=} appendToTitle Optional string to append to the nameList. 
    */
  function createNewSpreadsheets(driveFolderID, nameList, appendToTitle){
      var defaultSheet = ['Sheet1'];
      if (!Array.isArray(nameList)) {nameList = [nameList]};
  
      if (appendToTitle){
        nameList = nameList.map(function(name) {return name + appendToTitle});
      };
      
      for (var name in nameList){
        var newSheet = SpreadsheetApp.create(nameList[name]);
        var newSheetID = newSheet.getId();    
        var driveFile = DriveApp.getFileById(newSheetID);
        var driveFolder = DriveApp.getFolderById(driveFolderID);
        driveFolder.addFile(driveFile);    
        
        spreadsheetFunctions_.updateSheetsAndRanges(driveFolderID, 
          SpreadsheetApp.getActiveSpreadsheet().getSheets(), null, null, null);
          
        spreadsheetFunctions_.clearSheetsWithName(defaultSheet, newSheet);        
      };
  };
  
  /**
   * Retrieves an Array of file IDs for a Drive folder with the given ID.
   * @param {String} driveFolderID The ID for the Drive folder where files will 
        be retrieved.
   * @return {Array} An Array of file IDs from the Drive folder matching the 
        driveFolderID. 
   */
  function getFilesInFolder(driveFolderID) {
      var files = DriveApp.getFolderById(driveFolderID).getFiles();
      var filesIDList = [];
  
      while (files.hasNext()) {
        var file = files.next();
        filesIDList.push(file);
      };
      
      return filesIDList;
  };
  
  /**
   * Retrieves a list of Spreadsheets for a Drive folder with the given ID.
   * @param {String} driveFolderID The ID for the Drive folder where 
        Spreadsheets will be retrieved.
   * @return {Array} An Array of Spreadsheets IDs from the Drive folder matching 
        the driveFolderID. 
   */
  function getSpreadsheetsInFolder(driveFolderID){
      var filesFromFolderList = getFilesInFolder(driveFolderID);
      var spreadsheetIDList = [];
  
      for (var i = 0; i < filesFromFolderList.length; i++){
        try {
          spreadsheetIDList.push(SpreadsheetApp.openById(
            filesFromFolderList[i].getId()));
          } catch(error) {
            continue;
          };
      };
  
      return spreadsheetIDList;
  }

  /** 
    * Checks if a spreadsheet with a given name exists in a folder.
    * @param {String} driveFolderID The ID for the Drive folder where 
        Spreadsheets will be retrieved.
    * @param {String} title The string to search for a matching title.
    * @return {Bool} Whether or not a spreadsheet with the matching title exists 
        in the Drive folder.
    */
  function spreadsheetExistsInFolder(driveFolderID, title){
  var spreadsheets = getSpreadsheetsInFolder(driveFolderID);
  var spreadsheetExists = false;
  
  for (var spreadsheet in spreadsheets){
    if (spreadsheet.getName() === title){
      spreadsheetExists = true;
      return spreadsheetExists;
    };
  };
  
  return spreadsheetExists;  
  };
  
  return {
    createNewSpreadsheets: createNewSpreadsheets,  
    getFilesInFolder: getFilesInFolder,
    getSpreadsheetsInFolder: getSpreadsheetsInFolder,
    spreadsheetExistsInFolder: spreadsheetExistsInFolder
  };  
})();