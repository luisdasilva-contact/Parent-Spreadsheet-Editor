/** 
 * Class for handling interactions between Spreadsheets and Drive.
 */ 
class sheetsToDriveInteractions {    
  /**
    * For each string in a list, creates a new Spreadsheet with that title in 
        the given Drive folder ID, and initializes with all of the Sheets
        from the master document. Defaults to applying no permissions at this
        point.
    * @param {String} driveFolderID The ID for the Drive folder where 
        Spreadsheets will be created.
    * @param {String/Array<String>} nameList String or array of strings 
        containing the name(s) of each spreadsheet.
    * @param {String=} appendToTitle Optional string to append to the nameList. 
    */
  createNewSpreadsheets(driveFolderID, nameList, appendToTitle){
    const defaultSheet = `Sheet1`;
    const defaultSheetCopy = `Copy of Sheet1`;
    nameList = convertToArray(nameList);
    
    if (appendToTitle){
      nameList = nameList.map(function(name) {return name + appendToTitle});
    };
    
    for (let name in nameList){
      const newSheet = SpreadsheetApp.create(nameList[name]);
      const newSheetID = newSheet.getId();    
      const driveFile = DriveApp.getFileById(newSheetID);
      const driveFolder = DriveApp.getFolderById(driveFolderID);
      driveFolder.addFile(driveFile); 
      let activeSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      let activeSheetProtections = SpreadsheetApp.getActiveSpreadsheet()
          .getProtections(SpreadsheetApp.ProtectionType.SHEET);
      
      spreadsheetFunctionsObj.updateSheetsAndRanges(
        driveFolderID, activeSheets, null, activeSheetProtections);
      
      if (spreadsheetFunctionsObj.sheetNameExistsInSpreadsheet(
        defaultSheetCopy, newSheet)){
        spreadsheetFunctionsObj.clearSheetsWithName(defaultSheet, newSheet);  
        let sheetRename = newSheet.getSheetByName(`defaultSheetCopy`);
        sheetRename.setName(defaultSheet);
      } else {
        spreadsheetFunctionsObj.clearSheetsWithName(defaultSheet, newSheet);    
      };           
    };
  };
  
  /**
   * Retrieves an Array of file IDs for a Drive folder with the given ID.
   * @param {String} driveFolderID The ID for the Drive folder where files will 
        be retrieved.
   * @return {Array} An Array of file IDs from the Drive folder matching the 
        driveFolderID. 
   */
  getFilesInFolder(driveFolderID) {
    const files = DriveApp.getFolderById(driveFolderID).getFiles();
    const filesIDList = [];
    
    while (files.hasNext()) {
      let file = files.next();
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
  getSpreadsheetsInFolder(driveFolderID){
    const filesFromFolderList = this.getFilesInFolder(driveFolderID);
    let spreadsheetIDList = [];
    
    for (let i = 0; i < filesFromFolderList.length; i++){
      try {
        spreadsheetIDList.push(SpreadsheetApp.openById(
          filesFromFolderList[i].getId()));
      } catch(error) {
        continue;
      };
    };
    
    return spreadsheetIDList;
  };

  /** 
    * Checks if a spreadsheet with a given name exists in a folder.
    * @param {String} driveFolderID The ID for the Drive folder where 
        Spreadsheets will be retrieved.
    * @param {String} title The string to search for a matching title.
    * @return {Bool} Whether or not a spreadsheet with the matching title exists 
        in the Drive folder.
    */
  spreadsheetExistsInFolder(driveFolderID, title){
    const spreadsheets = getSpreadsheetsInFolder(driveFolderID);
    
    for (let spreadsheet in spreadsheets){
      if (spreadsheet.getName() === title){
        return true;
      };
    };
    
    return false;  
  };  
  
  /**
    * Validates whether or not there is at least 1 Spreadsheet in the target Drive 
      folder.
    * @param {string} folderID The ID for the Drive folder where Spreadsheets 
      will be set. 
   @return {boolean} True if there are sheets, false if there are not.
   */ 
  sheetsInFolderValidation(folderID){
    let UiFunctions = new UiClass();
    let SHEETS_IN_DRIVE_ERROR_TEXT = `There are no Sheets in the target Drive 
        Folder. Please create at least one from the Drive Folder itself, or 
        use the Initialize Sheets menu under Update User Sheets to create them.`;
    let clearNewlineRegex = /(\r\n|\n|\r)/gm;
    SHEETS_IN_DRIVE_ERROR_TEXT = SHEETS_IN_DRIVE_ERROR_TEXT.replace(
          clearNewlineRegex," ");
    if (sheetsToDriveInteractionsObj.getSpreadsheetsInFolder(
        folderID).length === 0){
        
        UiFunctions.displayAlert(SHEETS_IN_DRIVE_ERROR_TEXT);
        return false;
    } else {
      return true;
    };
  };
};

const sheetsToDriveInteractionsObj = new sheetsToDriveInteractions();