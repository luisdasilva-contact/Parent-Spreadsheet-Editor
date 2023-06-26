/*
 * Class used to retrieve the program's user-set properties.
 */ 
class folderProperties {  
  constructor() {
    this.properties = PropertiesService.getDocumentProperties();
    this.userList = this.properties.getProperty('USER_LIST');
    this.adminEmailList = this.properties.getProperty('ADMIN_EMAIL_LIST');
    this.folderID = this.properties.getProperty('FOLDER_ID');
    this.titleAppend = this.properties.getProperty('TITLE_APPEND');
    this.protectionBool = this.properties.getProperty('PROTECTION_BOOL');
  };
    
  /*
   * Retrieves the string of user names as a comma-separated array. If the 
        string cannot be parsed as such, null is returned.
   * @return {?array<string>} String array of emails. If string can't be split 
        by comma, null is returned.
   */
 getUserList(){
      try {            
          return this.userList.split(',');
      } catch (error){
          return null;
      };
  };
  
  /*
   * Retrieves the string of admin emails as a comma-separated array. If the 
        string cannot be parsed as such, null is returned.
   * @return {?array<string>} String array of emails. If string can't be split 
        by comma, null is returned.
   */
  getAdminEmailList(){
      try {            
          return this.adminEmailList.split(',');
      } catch (error){
          return null;
      };
  };
  
  /*
   * Gets the Drive Folder ID. If there is no string, or it can't be returned, 
        null is returned.
   * @return {?string} Subject string. If it can't be returned, null is 
      returned.
   */
 getFolderID(){
      try {            
          return this.folderID;
      } catch (error){
          return null;
      };
  };

  /*
   * Gets the user-defined string to append to spreadsheets they create. 
        If there is no string, or it can't be returned, null is returned.
   * @return {?string} Title string. If it can't be returned, null is 
        returned.
   */  
  getTitleAppend(){
      try {            
          return this.titleAppend;
      } catch (error){
          return null;
      };
  };

  /*
   * Gets the user-defined protection boolean. If true, protections from the 
        parent Spreadsheet will be applied to child Spreadsheets. If false,
        protections will not be updated. If there is no string, or it can't 
        be returned, null is returned.
   * @return {?string} Boolean string. If it can't be returned, null is 
        returned.
   */   
  getProtectionBool(){
     try {            
          return this.protectionBool;
      } catch (error){
          return null;
      };
  };
          
  /*
   * Sets the given property with the given userInput as a value.
   * @param {string} property The property that will be set in Google's 
        Property Service (the "key").
   * @param {string} userInput The value that will be set in Google's Property 
        Service (the "value").
   */
  setProperty(property, userInput){
      this.properties.setProperty(property, userInput);  
  };
};

function initializeAllSheets_() {
  let nameList = folderPropertiesObj.getUserList();
  let driveFolderID = folderPropertiesObj.getFolderID();
  let appendToTitle = folderPropertiesObj.getTitleAppend(); 
  let UiFunctions = new UiClass();
  
  if (!appendToTitle){
    appendToTitle = '';
  };
  
  if (!nameList){
    UiFunctions.displayAlert(`You must set at least one title in the Initialize
    Sheets --> Set User Names menu.`);  
    return;
  };
  
  if (!driveFolderID){
    UiFunctions.displayAlert(`You must set the Drive folder's ID in the Initialize
    Sheets --> Set Drive Folder ID menu.`)
  };
  
  try {
    sheetsToDriveInteractionsObj.createNewSpreadsheets(driveFolderID, nameList, 
      appendToTitle);  
    } catch(error){
      UiFunctions.displayAlert(`An error has occurred. Please ensure that all 
      values in the Initialize Sheets menu have been set properly.`);      
    };  
  
};

const contentPropertyEnum = {
  USER_LIST: "USER_LIST",
  FOLDER_ID: "FOLDER_ID",
  TITLE_APPEND: "TITLE_APPEND"
};

const folderPropertiesObj = new folderProperties();
