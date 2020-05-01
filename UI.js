/**
 * Class used to display content via Google Apps Script's built-in UI features. 
 */ 
class UiClass {
  constructor(){
    this.Ui = SpreadsheetApp.getUi();
  };  
  
  /**
   * Displays a window with text to the user.
   * @param {string} alertText The text to display to the user.
   */
  displayAlert(alertText){
    this.Ui.alert(alertText);
  };
  
  /**
   * Displays a text field to the user, allowing them to submit text, or cancel.
   * @param {string} promptText The text to display to the user above the text 
        field.
   * @return {?string} The text the user has entered if they clicked "OK". Null 
        if "CANCEL" is clicked.
   */
  displayPrompt(promptText){
    const prompt = this.Ui.prompt(promptText);
    
    if (prompt.getSelectedButton() === this.Ui.Button.OK){
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
  displayYesNoChoice(promptText){
    const prompt = this.Ui.alert(promptText, this.Ui.ButtonSet.YES_NO_CANCEL);
    
    if (prompt === this.Ui.Button.YES){
      return 'YES';
    } else if (prompt === this.Ui.Button.NO){
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
  displayHTMLAlert(HtmlOutput, title){
    this.Ui.showModelessDialog(HtmlOutput, title);            
  };
    
  /**
  * Opens the document's sidebar, and displays HTML content.
  * @param {string} HTML content to display in the window.
  */
  displaySidebar(htmlOutput){
    this.Ui.showSidebar(htmlOutput);
  };
};

const UiFunctions = new UiClass();