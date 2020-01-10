/**
 * Module used to display content via Google Apps Script's built-in UI features. 
      Contains the following functions: 
 * @return {function} displayAlert Displays a window with text to the user.
 * @return {function} displayPrompt Displays a text field to the user, allowing 
      them to submit text, or cancel.
 * @return {function} displayYesNoChoice Displays a window with text to the 
      user, allowing them to select "YES", "NO", or "CANCEL".
 * @return {function} displayHTMLAlert Displays a window with text to the user, 
      taking HTML content.
 */ 
var UiFunctions = (function(){
  var Ui = SpreadsheetApp.getUi();
  
  /**
   * Displays a window with text to the user.
   * @param {string} alertText The text to display to the user.
   */
  function displayAlert(alertText){
    Ui.alert(alertText);
  };
  
  /**
   * Displays a text field to the user, allowing them to submit text, or cancel.
   * @param {string} promptText The text to display to the user above the text 
        field.
   * @return {?string} The text the user has entered if they clicked "OK". Null 
        if "CANCEL" is clicked.
   */
  function displayPrompt(promptText){
    var prompt = Ui.prompt(promptText);
    
    if (prompt.getSelectedButton() === Ui.Button.OK){
      return prompt.getResponseText();
    } else {
      return null;
    };
  };
  
  /**
   * Displays a window with text to the user, allowing them to select "YES", 
        "NO", or "CANCEL".
   * @param {string} promptText The text to display to the user above the "YES", 
        "NO", and "CANCEL" buttons.
   * @return {?string} The user's choice. "YES" if they clicked "YES", "NO" if 
        they clicked "NO", null if they clicked "CANCEL".
   */
  function displayYesNoChoice(promptText){
    var prompt = Ui.alert(promptText, Ui.ButtonSet.YES_NO_CANCEL);
    
    if (prompt === Ui.Button.YES){
      return 'YES';
    } else if (prompt === Ui.Button.NO){
      return 'NO';
    } else {
      return null;
    };
  };
    
    /**
   * Displays a window with text to the user, taking HTML content.
   * @param {string} HTML content to display in the window.
   * @param {string} title The text to display as the window's title.
   */
  function displayHTMLAlert(HtmlOutput, title){
    Ui.showModelessDialog(HtmlOutput, title);            
  };
    
    /**
     * Opens the document's sidebar, and displays HTML content.
     * @param {string} HTML content to display in the window.
     */
    function displaySidebar(htmlOutput){
      Ui.showSidebar(htmlOutput);
    };
  
    return {
      getUi: Ui,
      displayPrompt: displayPrompt,
      displayAlert: displayAlert,
      displayYesNoChoice: displayYesNoChoice,
      displayHTMLAlert: displayHTMLAlert,
      displaySidebar: displaySidebar
    };
})();



/**
 * If the active cell is part of a named range, updates that named range in all 
      Spreadsheets in the designated folder,
      using this master sheet as a template.
 */
function updateRange_(){
  var activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var namedRangesInSheet = SpreadsheetApp.getActiveSheet().getNamedRanges();
  var activeCellRow = activeCell.getRow();
  var activeCellColumn = activeCell.getColumn();
  var rangeToUpdate = null;
  var permissionsObject = null;
  var sourceSheets = SpreadsheetApp.getActiveSheet(); 
  var editors = folderProperties.getAdminEmailList();
  for (var namedRange in namedRangesInSheet){
    var range = namedRangesInSheet[namedRange].getRange();
    
    if (activeCellRow >= range.getRow() && activeCellRow <= range.getLastRow() 
      && activeCellColumn >= range.getColumn() && activeCellColumn <= 
          range.getLastColumn()){
            rangeToUpdate = range.getA1Notation();
            break;
        };
    };

    var shouldProtect = UiFunctions.displayYesNoChoice('Would you like to ' +
    'assign edit protections to this range?');

    if (shouldProtect === 'YES') {
      permissionsObject = rangeToUpdate;
    } else if (shouldProtect === 'NO') {
      editors = null;
    } else if (shouldProtect === null){
      return;
    }
        
    if (rangeToUpdate){
      spreadsheetFunctions_.updateSheetsAndRanges(
        folderProperties.getFolderID(), sourceSheets, rangeToUpdate,
        permissionsObject, editors);
    };
};

/**
 * Updates the active sheet in all Spreadsheets in the designated folder, 
      using this master sheet as a template.
 */
function updateCurrentSheet_() {
  var permissionsObject = null;
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var editors = folderProperties.getAdminEmailList();
  var activeSheetRange = activeSheet.getRange(1, 1, activeSheet.getLastRow(),
    activeSheet.getLastColumn()).getA1Notation();

    var shouldProtect = UiFunctions.displayYesNoChoice('Would you like to ' +
    'assign edit protections to this range?');

    if (shouldProtect === 'YES') {
      permissionsObject = activeSheetRange;
    } else if (shouldProtect === 'NO'){
      editors = null;
    } else if (shouldProtect === null){
      return;
    }

  spreadsheetFunctions_.updateSheetsAndRanges(
    folderProperties.getFolderID(), activeSheet, activeSheetRange, 
    permissionsObject, editors);
};

/**
 * Updates the all sheets in all Spreadsheets in the designated folder, using 
      this master sheet as a template.
 */
function updateAllSheets_() {
  var allSheets = SpreadsheetApp.getActive().getSheets();
  var editors = folderProperties.getAdminEmailList();
  var shouldProtect = UiFunctions.displayYesNoChoice('Would you like to ' +
    'assign edit protections to this range?');

    if (shouldProtect === 'NO') {
      var editors = null;
    } else if (shouldProtect === null){
      return;
    } 

  spreadsheetFunctions_.updateSheetsAndRanges(folderProperties.getFolderID(), 
    allSheets, null, null, editors);
};

/**
 * Creates all the sheets from the githubSample.gs file, creating a sheet for
      each user, and assigning the editors with permission to use this 
      program to edit their sheets.
 */
function initializeAllSheets(){
  initializeAllSheetsSample_();
};

/**
 * Provides a dialogue explaining to the user how they can utilize the program.
 */
function menuHelp_(){
  var html = '<html><head><link rel=\'stylesheet\' ' + 
  'href=\'https://fonts.googleapis.com/css?family=Roboto\'>' +
  '<style>body {font-family: \'Roboto\', sans-serif; font-size: 14px}' +
  '</style></head><body><font face=\'Roboto\'>This master Spreadsheet is ' +
  'used to manage changes for all spreadsheets in ' + 
  '<a href=\'https://www.drive.google.com/drive/folders/' + 
  folderProperties.getFolderID() + '\' target=\'blank\'>this Drive ' +
  'folder.</a> These are the functions for each button in this menu: ' +
  '<br><br>' +
  'Update Current Range: <br>' +
  'Applies any text changes made in the range current being viewed to ' +
  'matching sheets in the Drive folder, including formula changes, or ' +
  'changes to cell text. Ranges can be viewed by going  to Data --> ' +
  'Named ranges. ' +
  'Update Current Sheet: <br>' +
  'Applies any text changes made in the sheet currently being viewed to ' +
  'matching sheets in the Drive folder, including formulas changes, or ' +
  'changes to cell text. <br>' +        
  '<br><br>Update All Sheets: <br>' +
  'Applies any text changes made in all sheets to the Spreadsheets in the ' +
  'Drive folder, including formulas changes, or changes to cell text. ' +
  'NOTE: This is resource-intensive. Use only for large-scale changes.' +
  '</font></body></html>';
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(700)
      .setHeight(250);
  UiFunctions.displayHTMLAlert(htmlOutput, 'Help');
};

/**
 * Function placed in UI to activate method to set user list. 
 */
function setUserList(){
  controller.setContent('USER_LIST');
};

/**
 * Function placed in UI to activate method to set admin list. 
 */
function setAdminList(){
  controller.setContent('ADMIN_EMAIL_LIST');
};

/**
 * Function placed in UI to activate method to set Drive folder ID. 
 */
function setFolderID(){
  controller.setContent('FOLDER_ID');
};

/**
 * Function placed in UI to activate method to set text to append to each
      file's title. 
 */
function setTitleAppend(){
  controller.setContent('TITLE_APPEND');
};

/**
 * Function placed in UI to activate method to clear all properties. 
 */
function clearAllProperties(){
  controller.clearAllProperties();
};

/**
 * Function to automatically build menu upon opening the document.
 * @param {Event} onOpen event containing context regarding the document upon 
      opening.
 */
function onOpen(e){  
  UiFunctions.getUi
  .createMenu('Update User Sheets')
  .addItem('Help', 'menuHelp_')
  .addSeparator()
  .addSubMenu(UiFunctions.getUi.createMenu('Initialize Sheets')
                               .addItem('Create Sheets',
                                 'initializeAllSheetsSample_')
                               .addItem('Set User Names', 'setUserList')
                               .addItem('Set Admin List', 'setAdminList')
                               .addItem('Set Folder ID', 'setFolderID')
                               .addItem('Set Title Append', 'setTitleAppend')
                               .addSeparator()
                               .addItem('Clear Properties', 
                                'clearAllProperties'))
  .addSubMenu(UiFunctions.getUi.createMenu('Update Sheets')
                                .addItem('Update Current Range', 'updateRange_')
  .addItem('Update Current Sheet', 'updateCurrentSheet_')
  .addItem('Update All Sheets', 'updateAllSheets_')) 
  .addToUi();    
};