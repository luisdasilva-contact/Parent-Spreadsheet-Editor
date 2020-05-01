/**
 * Examines the parameter for being an Array. If it is an Array, the unchanged
      Array is returned. If it is not an Array, a single-item Array containing
      the parameter is returned.
 * @param {Object/Array<Object>} toArray The object to check.toArray
 * @return {Array<Object>} The Array, either unchanged from the user input, or
      the initial object now returned as a single-item array. 
 */
function convertToArray(toArray){
  if (toArray === null){
    return null;
  } else if (toArray === undefined){
    return undefined; 
  } else {
  return (Array.isArray(toArray)? toArray : [toArray]);
  };
};

/**
 * Function to determine whether the given range is within a larger range.
 * @param {Range} innerRange The range to evaluate for being within the larger 
      range.
 * @param {Range} outerRange The range to evaluate for containing the 
      innerRange.
 @return {boolean} True if innerRange is within outerRange, false otherwise.
 */
function isRangeWithinLargerRange(innerRange, outerRange){
  if (innerRange.getRow() >= outerRange.getRow() && 
      innerRange.getLastRow() <= outerRange.getLastRow() && 
      innerRange.getColumn() >= outerRange.getColumn() && 
      innerRange.getLastColumn() <= outerRange.getLastColumn()){
              return true;
          };
  return false;
};

/**
 * Function to run checks that the range the user has chosen to update passes 
      validity checks, in the interest of avoiding errors.
 * @param {Range} rangeToUpdate The range of cells that will be updated in the 
      target sheet(s).
 * @return {boolean} True if validation passed, false if failed.
 */
function rangeValidation(rangeToUpdate){
  if (!rangeToUpdate){
  UiFunctions.displayAlert(`No Named Range found. Please make sure the selected
        cell is part of a Named Range.`);
        return false;
  }; 
  
  return true;
};

/**
 * If the active cell is part of a named range, updates that named range in all 
      Spreadsheets in the designated folder,
      using this master sheet as a template.
 */
function updateRange_(){
  const sourceSheet = SpreadsheetApp.getActiveSheet(); 
  const activeCell = sourceSheet.getActiveCell();
  const namedRangesInSheet = SpreadsheetApp.getActiveSheet().getNamedRanges();
  const folderID = folderPropertiesObj.getFolderID();  
  let sourceRangeArray = [];
  const NAMED_RANGE_ALERT_TEXT = `No Named Range found. Please make sure the 
      selected cell is part of a Named Range.`;
  
  for (let namedRange in namedRangesInSheet){
    let range = namedRangesInSheet[namedRange].getRange();
    
    if (isRangeWithinLargerRange(activeCell, range)){
      sourceRangeArray.push(range);
      sourceRange = range;
    };
  };
 
 if (sourceRangeArray.length === 0){
   UiFunctions.displayAlert(NAMED_RANGE_ALERT_TEXT);
    return false;   
 };
 
  if (!sheetsToDriveInteractionsObj.sheetsInFolderValidation(folderID)){return};  
  let protectionsArray = []; 
  if (!rangeValidation(sourceRange)) {return};
   
  if (folderPropertiesObj.getProtectionBool() === "true"){
    let protectionsItem = {};
    
    let currentRangeProtections = 
        sourceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (let rangeProtection in currentRangeProtections){
      let isProtectionInRangeToUpdate = false;
      
      for (let sourceRange in sourceRangeArray){
         if (isRangeWithinLargerRange(
              currentRangeProtections[rangeProtection].getRange(), 
              sourceRangeArray[sourceRange])){
                isProtectionInRangeToUpdate = true;
                break;
         };
      };
      
      if (isProtectionInRangeToUpdate){      
        protectionsArray.push(currentRangeProtections[rangeProtection]);
      };
    };   
  };   
  
  if (sourceRange){
    try {
      spreadsheetFunctionsObj.updateSheetsAndRanges(
        folderID, sourceSheet, sourceRange, protectionsArray);    
    } catch(error){
      UiFunctions.displayAlert(`An error has occurred. Please ensure that all 
                                  values in the Initialize Sheets menu have been 
                                  set properly.`);
    };
  };
  
};

/**
 * Updates the active sheet in all Spreadsheets in the designated folder, 
      using this master sheet as a template.
 */
function updateCurrentSheet_() {
  let sheetProtectionType = SpreadsheetApp.ProtectionType.SHEET;
  let folderID = folderPropertiesObj.getFolderID();
  let activeSheet = SpreadsheetApp.getActiveSheet();
  if (!sheetsToDriveInteractionsObj.sheetsInFolderValidation(folderID)) {
    return;
  };
  
  let currentSheetProtections = [];
  
  if (folderPropertiesObj.getProtectionBool() === "true"){
    currentSheetProtections = activeSheet.getProtections(sheetProtectionType);
  };
 
  try {
  spreadsheetFunctionsObj.updateSheetsAndRanges(
        folderID, activeSheet, null, currentSheetProtections);
    } catch(error){
      UiFunctions.displayAlert(`An error has occurred. Please ensure that a 
          target Drive folder has been set under Update User Sheets --> 
          Initialize Sheets --> Set Folder ID. `);
    };
};

/**
 * Updates all sheets in all Spreadsheets in the designated folder, using this 
      master sheet as a template.
 */
function updateAllSheets_() {
  let sheetProtectionType = SpreadsheetApp.ProtectionType.SHEET;
  let folderID = folderPropertiesObj.getFolderID();
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheets = activeSpreadsheet.getSheets();

  if (!sheetsToDriveInteractionsObj.sheetsInFolderValidation(folderID)) {
    return;
  };
  
  let protectionsArray = [];  
  if (folderPropertiesObj.getProtectionBool() === "true"){
    protectionsArray = activeSpreadsheet.getProtections(sheetProtectionType);
  };
  
  try {
    spreadsheetFunctionsObj.updateSheetsAndRanges(
          folderID, activeSheets, null, protectionsArray);
      } catch(error){
        UiFunctions.displayAlert(`An error has occurred. Please ensure that a 
            target Drive folder has been set under Update User Sheets --> 
            Initialize Sheets --> Set Folder ID. `);
  };
};



/**
 * Displays a dialogue explaining to the user how they can utilize the program.
 */
function menuHelp_(){
  const folderHTMLString = (folderPropertiesObj.getFolderID()) ? 
    `<a href='https://www.drive.google.com/drive/folders/' 
    ${folderPropertiesObj.getFolderID()}target='blank'>this Drive 
    folder.</a>` : `the Drive folder set in Update User Sheets --> Initialize
    Sheets --> Set Folder ID.`;
    
  const html = `<html><head><link rel='stylesheet' 
  href='https://fonts.googleapis.com/css?family=Roboto'>
  <style>body {font-family: 'Roboto', sans-serif; font-size: 14px}
  </style></head><body><font face='Roboto'>This master Spreadsheet is 
  used to manage changes for all spreadsheets in ${folderHTMLString}  
  These are the functions for each button in this menu: 
  <br><br>
  <b>Initialize Sheets:</b><br>
  &emsp;Create Sheets: Create all Spreadsheets for the user-defined parameters
  in the "Set" functions.<br>
  &emsp;Set User Names: Set names for each of the Spreadsheets.<br>
  &emsp;Set Drive Folder ID: Set the folder ID where each Spreadsheet will be
  stored.<br>
  &emsp;Set Title Append: Set a string to append to each Spreadsheet name.<br>
  &emsp;Clear Properties: Clear all user-entered properties from the Set functions.
  <br>
  <b>Update Sheets Sheets:</b><br>
  &emsp;Update Current Range: Update the Named Range the user currently has selected.
  <br>
  &emsp;Update Current Sheet: Update the Sheet the user is currently viewing.  
  <br>
  &emsp;Update All Sheets: Update all Sheets in the master Spreadsheet. <br>
  &emsp;NOTE: This is resource-intensive. Useonly for large-scale changes.<br>
  &emsp;Apply Permissions?: Whether or not the user wants the protections in the master
  Spreadsheet applied to <br>&emsp;all children.
  </font></body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(700)
      .setHeight(250);
  UiFunctions.displayHTMLAlert(htmlOutput, 'Help');
};

const menuItems = {
  initializeAllSheets: function() {
    initializeAllSheets_();
  },
  setUserList: function() {
    controllerObj.setContent(contentPropertyEnum.USER_LIST);
  },
  setFolderID: function() {
    controllerObj.setContent(contentPropertyEnum.FOLDER_ID);
  },
  setTitleAppend: function() {
    controllerObj.setContent(contentPropertyEnum.TITLE_APPEND);
  },
  clearAllProperties: function() {
    controllerObj.clearAllProperties();
  },
  updateRange: function(){
    updateRange_();
  },
  updateCurrentSheet: function(){
    updateCurrentSheet_();
  },
  updateAllSheets: function(){
    updateAllSheets_();
  },
  applyPermissions: function(){
    controllerObj.applyPermissions_();
  }
};

/**
 * Function to automatically build menu upon opening the document.
 * @param {Event} onOpen event containing context regarding the document upon 
      opening.
 */
function onOpen(e){  
  UiFunctions.Ui
  .createMenu('Update User Sheets')
  .addItem('Help', 'menuHelp_')
  .addSeparator()
  .addSubMenu(UiFunctions.Ui.createMenu('Initialize Sheets')
                               .addItem('Create Sheets',
                                 'menuItems.initializeAllSheets')
                               .addItem('Set User Names', 
                                  'menuItems.setUserList')
                               .addItem('Set Folder ID', 
                                  'menuItems.setFolderID')
                               .addItem('Set Title Append', 
                                  'menuItems.setTitleAppend')
                               .addSeparator()
                               .addItem('Clear Properties', 
                                'menuItems.clearAllProperties'))
  .addSubMenu(UiFunctions.Ui.createMenu('Update Sheets')
                                .addItem('Update Current Range', 
                                    'menuItems.updateRange')
  .addItem('Update Current Sheet', 'menuItems.updateCurrentSheet')
  .addItem('Update All Sheets', 'menuItems.updateAllSheets')
  .addSeparator()
  .addItem('Apply Protections?', 'menuItems.applyPermissions'))
  .addToUi();    
};