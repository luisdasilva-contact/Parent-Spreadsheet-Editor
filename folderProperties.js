/*
 * Module used to retrieve properties related to the email the user will send 
      out. Contains the following functions: 
 * @return {function} getEmails Retrieves the string of emails as a 
      comma-separated array. If the string cannot be parsed as such, null is 
      returned.
 * @return {function} getCC Retrieves the string of CC emails as a 
      comma-separated array. If the string cannot be parsed as such, null is 
      returned.
 * @return {function} getSubject Gets the subject string. If there is no string, 
      or it can't be returned, null is returned.
 * @return {function} getBodyText Gets the body text string. If there is no body 
      text, or it can't be returned, null is returned.
 * @return {function} setProperty Sets the given property with the given 
      userInput as a value.
 * @return {function} Retrieves an enum object for folderProperties.
 */ 
var folderProperties = (function() {  
  var properties = PropertiesService.getDocumentProperties();
  var userList = properties.getProperty('USER_LIST');
  var adminEmailList = properties.getProperty('ADMIN_EMAIL_LIST');
  var folderID = properties.getProperty('FOLDER_ID');
  var titleAppend = properties.getProperty('TITLE_APPEND');
  
  /*
   * Retrieves the string of user names as a comma-separated array. If the 
        string cannot be parsed as such, null is returned.
   * @return {?array<string>} String array of emails. If string can't be split 
        by comma, null is returned.
   */
  function getUserList(){
      try {            
          return userList.split(',');
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
  function getAdminEmailList(){
      try {            
          return adminEmailList.split(',');
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
  function getFolderID(){
      try {            
          return folderID;
      } catch (error){
          return null;
      };
  };
  
  function getTitleAppend(){
      try {            
          return titleAppend;
      } catch (error){
          return null;
      };
  }
          
  /*
   * Sets the given property with the given userInput as a value.
   * @param {string} property The property that will be set in Google's 
        Property Service (the "key").
   * @param {string} userInput The value that will be set in Google's Property 
        Service (the "value").
   */
  function setProperty(property, userInput){
      properties.setProperty(property, userInput);  
  };

  return {
      getUserList: getUserList,
      getAdminEmailList: getAdminEmailList,
      getFolderID: getFolderID,    
      getTitleAppend: getTitleAppend,
      setProperty: setProperty,
  };  
})();

function initializeAllSheetsSample_() {
  var nameList = folderProperties.getUserList();
  var driveFolderID = folderProperties.getFolderID();
  var appendToTitle = folderProperties.getTitleAppend(); 
  
  if (!appendToTitle){
    appendToTitle = '';
  }
  try {
    userSheetsModule_.createNewSpreadsheets(driveFolderID, nameList, 
      appendToTitle);  
    } catch(error){
      UiFunctions.displayAlert('An error has occurred. Please ensure that all' + 
      'values in the Initialize Sheets menu have been set properly.');      
    };
};