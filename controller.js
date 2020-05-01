/**
 * Class acting as a mediator between the UI and user-stored properties.
 */
class controller {
  constructor() {
    this.existingListText_ = `The existing entry is: \n`;
  }
  /**
   * Sets the user's entered content as a property via Google's Properties 
        Service, dependant upon an entered enum. 
   * @param {string} The value that will be set as a property in the Properties 
        Service.
   */  
  setContent(enumValue) {
    switch (enumValue) {
      case (contentPropertyEnum.USER_LIST):
        const currentUserList = folderPropertiesObj.getUserList();
        let UiNamePromptString =
            `Please enter a list of comma-separated names, or a single name, 
            that will each have their own Spreadsheet, made from the Master 
            Sheet.`;
          
        if (currentUserList) {
          UiNamePromptString += this.existingListText_ + currentUserList;
        };

        this.setText_(UiNamePromptString, enumValue);
        break;
      case (contentPropertyEnum.ADMIN_LIST):
        const currentAdminEmailList = folderPropertiesObj.getAdminEmailList();
        let UiAdminEmailPromptString =
            `Please enter a list of comma-separated emails, or a single email, 
            that will be the editors for this master sheet.`;

        if (currentAdminEmailList) {
          UiAdminEmailPromptString += 
              this.existingListText_ + currentAdminEmailList;
        };

        this.setEmails_(UiAdminEmailPromptString, enumValue);
        break;
      case (contentPropertyEnum.FOLDER_ID):
        const currentFolderID = folderPropertiesObj.getFolderID();
        let UiFolderIDPromptString =
            `Please enter the Google Drive folder's ID. This is the alphanumeric code 
            in the URL when you navigate to the Drive folder in question.`;

        if (currentFolderID) {
          UiFolderIDPromptString += this.existingListText_ + currentFolderID;
        };

        this.setText_(UiFolderIDPromptString, enumValue);
        break;   
      case (contentPropertyEnum.TITLE_APPEND):
        const currentTitleAppend = folderPropertiesObj.getTitleAppend();
        let UiTitlePromptString = 
            `Please enter the title you'd like appended to each user's
            sheet.`;
        
        if (currentTitleAppend){
          UiTitlePromptString += this.existingListText_ + currentTitleAppend;
        };
        
        this.setText_(UiTitlePromptString, enumValue);
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
  setEmails_(UiPromptString, propertyToSet) {
    const EMAIL_REGEX = /[a-zA-Z0-9._%-]+@[a-zA-Z\d\-]+./;
    let setEmailsWindowResponse = UiFunctions.displayPrompt(UiPromptString);

    if (setEmailsWindowResponse) {
      let EMAIL_ERROR_TEXT = 'Error! This email is not valid: ';
      let clearNewlineRegex = /(\r\n|\n|\r)/gm;
      EMAIL_ERROR_TEXT = EMAIL_ERROR_TEXT.replace(clearNewlineRegex," ");
      const emailsSplit = setEmailsWindowResponse.split(',');
      let validResponse = true;      

      for (let email in emailsSplit){
        if (!EMAIL_REGEX.test(emailsSplit[email])){
          UiFunctions.displayAlert(EMAIL_ERROR_TEXT + emailsSplit[email]);
          validResponse = false;
          break;
        };  
      };

      if (validResponse) {
        folderPropertiesObj.setProperty(propertyToSet, setEmailsWindowResponse);
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
  setText_(UiPromptString, propertyToSet) {
    const setSubjectWindowResponse = UiFunctions.displayPrompt(UiPromptString);
    
    if (setSubjectWindowResponse) {
      folderPropertiesObj.setProperty(propertyToSet, setSubjectWindowResponse);
    };
  };

  /**
    * Sets an internal boolean to whether or not protections will be carried from 
          the master sheet to its children.
   */
  applyPermissions_() {
    let applyPermissionsBool = UiFunctions.displayYesNoChoice(`Would you like 
        changes in range and sheet protections to be applied to the target 
        range(s) and sheet(s)?`);
    
    switch(applyPermissionsBool){
      case "YES":
        folderPropertiesObj.setProperty('PROTECTION_BOOL', "true");
        break;
      case "NO":
        folderPropertiesObj.setProperty('PROTECTION_BOOL', "false");
        break;
      default:
        return;
    };
  };

  /**
   * Deletes all of the document's properties related to this application.
   */
  clearAllProperties() {
    PropertiesService.getDocumentProperties().deleteAllProperties();
  };
};

const controllerObj = new controller();