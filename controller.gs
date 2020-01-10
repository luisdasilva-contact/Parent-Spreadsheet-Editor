// this was added in last minute - almost certainly a giant mess


/**
 * controller object acting as a mediator between the UI and user-stored 
      properties.
 * @return {function} setContent Sets the user's entered content as a property 
      via Google's Properties Service, dependant upon an entered enum. 
 * @return {function} clearAllProperties Deletes all of the document's
      properties related to this application.
 */
var controller = (function() {
  /**
   * Sets the user's entered content as a property via Google's Properties 
        Service, dependant upon an entered enum. 
   * @param {string} The value that will be set as a property in the Properties 
        Service.
   */
  function setContent(enum) {
    switch (enum) {
      case ('USER_LIST'):
        var currentUserList = folderProperties.getUserList();
        var UiPromptString =
          'Please enter a list of comma-separated names, or a single name, ' +
          'that will each have their own Spreadsheet, made from the Master ' +
          'Sheet. ';
        var existingListPrepend = 'The existing list is as follows: \n' +
          currentUserList;
        var propertyToSet = 'USER_LIST';

        if (currentUserList) {
          UiPromptString = UiPromptString + existingListPrepend;
        };

        setText(UiPromptString, propertyToSet);
        break;
      case ('ADMIN_EMAIL_LIST'):
        var currentAdminEmailList = folderProperties.getAdminEmailList();
        var UiPromptString =
          'Please enter a list of comma-separated emails, or a single email, ' + 
          'that will be the editors for this master sheet. ';
        var existingListPrepend = 'The existing list is as follows: \n' +
          currentAdminEmailList;
        var propertyToSet = 'ADMIN_EMAIL_LIST';

        if (currentAdminEmailList) {
          UiPromptString += existingListPrepend;
        };

        setEmails(UiPromptString, propertyToSet);
        break;
      case ('FOLDER_ID'):
        var currentFolderID = folderProperties.getFolderID();
        var UiPromptString =
          'Please enter the Drive Folder\s ID. ';
        var existingIDPrepend = 'The existing ID is as follows: \n' +
          currentFolderID;
        var propertyToSet = 'FOLDER_ID';

        if (currentFolderID) {
          UiPromptString += existingIDPrepend;
        };

        setText(UiPromptString, propertyToSet);
        break;   
      case ('TITLE_APPEND'):
        var currentTitleAppend = folderProperties.getTitleAppend();
        var UiPromptString = 
          'Please enter the title you\'d like appended to each user\s ' +
          'sheet.'
        var existingListPrepend = 'The existing text is as follows: \n' +
          currentTitleAppend;
        var propertyToSet = 'TITLE_APPEND';
        
        if (currentTitleAppend){
          UiPromptString += existingIDPrepend;
        };
        
        setText(UiPromptString, propertyToSet);
        break;
    };      
  };

  /**
   * Sets any properties related to emails. Includes regex check for valid list 
        of emails.
   * @param {string} UiPromptString The string that will be displayed to the 
        user in the window prompting them to enter a list of emails.
   * @param {string} propertyToSet The property that will be set via Properties 
        Service.
   */
  function setEmails(UiPromptString, propertyToSet) {
    const EMAIL_REGEX = /[a-zA-Z0-9._%-]+@[a-zA-Z\d\-]+./;
    var setEmailsWindowResponse = UiFunctions.displayPrompt(UiPromptString);

    if (setEmailsWindowResponse) {
      var emailsSplit = setEmailsWindowResponse.split(',');
      var validResponse = true;
      const ERROR_TEXT = 'Error! This email is not valid: ';

      for (var email in emailsSplit){
        if (!EMAIL_REGEX.test(emailsSplit[email])){
          UiFunctions.displayAlert(ERROR_TEXT + emailsSplit[email]);
          validResponse = false;
          break;
        };  
      };

      if (validResponse) {
        folderProperties.setProperty(propertyToSet, setEmailsWindowResponse);
      };
    };
  };

  /**
   * Sets any properties related to text.
   * @param {string} UiPromptString The string that will be displayed to the 
        user in the window, prompting them to enter text.
   * @param {string} propertyToSet The property that will be set via Properties 
        Service.
   */
  function setText(UiPromptString, propertyToSet) {
    var setSubjectWindowResponse = UiFunctions.displayPrompt(
    UiPromptString);
    if (setSubjectWindowResponse) {
      folderProperties.setProperty(propertyToSet, setSubjectWindowResponse);
    };
  };

  /**
   * Deletes all of the document's properties related to this application.
   */
  function clearAllProperties() {
    PropertiesService.getDocumentProperties().deleteAllProperties();
  };

  return {
    setContent: setContent,
    clearAllProperties: clearAllProperties
  };
})();
