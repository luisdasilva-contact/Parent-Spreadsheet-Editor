/** 
  * Class for performing functions related to Spreadsheet manipulation. 
  */
 class spreadsheetFunctions {
  /**
   * For a Sheet, Removes the "Copy of " prefix that Google Sheets applies to 
        duplicated sheets.   
    * @param {Sheet/Array<Sheet>} targetSheet The Sheet that will have the 
        prefix removed. Takes either a single Sheet or an array of Sheet 
        objects. 
    */
  removeCopyOfPrefixFromSheets(targetSheets) {    
    targetSheets = convertToArray(targetSheets);
    
    for (let targetSheet in targetSheets) {
      const REGEX = /^Copy\sof\s/;
      targetSheets[targetSheet].setName(
        targetSheets[targetSheet].getName().replace(REGEX,''));
    };
  };

  /**
    * Checks if a Sheet with a specific name already exists in a target 
        Spreadsheet.
    * @param {string} sheetName The name to check for in the target 
        Spreadsheet's Sheets. 
    * @param {Spreadsheet} targetSpreadsheet The Spreadsheet that will be 
        searched for a Sheet titled the sheetName string. 
    * @return {boolean} True if a sheet with the sheetName exists 
        in the targetSpreadsheet, false if not.
    */
  sheetNameExistsInSpreadsheet(sheetName, targetSpreadsheet) {
    const targetSheets = targetSpreadsheet.getSheets();

    for (let i = 0; i < targetSheets.length; i++){
        if (targetSheets[i].getName() === sheetName) {
            return true;
        };
    };      
    return false;
  };

    /**
      * Utilizes sheetNameExistsInSpreadsheet to delete a Sheet if its title 
          exists in a Spreadsheet.
      * @param {string/Array<string>} sheetNamesToDelete String or array of 
          strings checked for matching titles with those in targetSpreadsheet; 
          if matching titles are found, targetSpreadsheet is deleted.
      * @param {Spreadsheet} targetSpreadsheet The spreadsheet that will be 
          searched for a sheet with the sheetNamesToDelete strings. 
      */
    clearSheetsWithName(sheetNamesToDelete, targetSpreadsheet) {
      sheetNamesToDelete = convertToArray(sheetNamesToDelete);
      
      for (let sheetName in sheetNamesToDelete){
          if (this.sheetNameExistsInSpreadsheet(
              sheetNamesToDelete[sheetName], targetSpreadsheet)){
                targetSpreadsheet.deleteSheet(
                  targetSpreadsheet.getSheetByName(
                    sheetNamesToDelete[sheetName]));
          };
      };
    };

    /** 
    * Given a Range, unprotects it if it has protections applied.
    * @param {Range/Array<Range>} rangesToUnprotect A range or Array of Ranges to unprotect.
    */
    unprotectRange(rangesToUnprotect){
      rangesToUnprotect = convertToArray(rangesToUnprotect);
      for (let rangeToUnprotect in rangesToUnprotect){
        let sheet = rangesToUnprotect[rangeToUnprotect].getSheet();
        let sheetProtectedRanges = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
       
        for (let range in sheetProtectedRanges){
          if (sheetProtectedRanges[range].getRange().getA1Notation() ===
              rangesToUnprotect[rangeToUnprotect].getA1Notation()){
            sheetProtectedRanges[range].remove();
          };
        };
      };
    };    
       
    /**
     * Clears existing protections for a given range in a sheet, replacing them 
        with the list of users in NewEditors.
     * @param {Range/Array<Range>} ranges The Range(s) that will have its/their
            existing editors removed, replaced with those in newEditors.
     * @param {Array<User>} newEditors An array of User objects representing the 
        new editors for the range.}
     */
    replaceRangeProtections(ranges, newEditors){
      let UiFunctions = new UiClass();
      ranges = convertToArray(ranges);

      try {
        for (let range in ranges){
          const protectionObject = ranges[range].protect();
          protectionObject.addEditor(Session.getEffectiveUser().getEmail());
          newEditors = convertToArray(newEditors);
          protectionObject.addEditors(newEditors);
          let removeEditorsList = protectionObject.getEditors().map(
              editor => editor.getEmail());
          protectionObject.removeEditors(
            removeEditorsList).addEditors(newEditors);  
        };
      } catch(error){
      UiFunctions.displayAlert('Cannot apply protections for admin.');
      return;
      };    
    };
    
    /**   
     * Sets protections for a Sheet, allowing users from a list of emails
          to edit the Sheet. Note that this removes the existing set of editors
          and replaces them with those from the editorEmails Array!
     * @param {Array<string>} editorEmails List of emails to add as editors
          to the sheetsToProtect.
     * @param {Sheet/Array<Sheet>} sheetsToProtect A sheet or array of Sheets 
          to apply protections to.
     */
   setSheetProtections(editorEmails, sheetsToProtect){   
     sheetsToProtect = convertToArray(sheetsToProtect);

     for (let sheet in sheetsToProtect){
       const protectionObj = sheetsToProtect[sheet].getProtections(
            SpreadsheetApp.ProtectionType.SHEET);
       if (protectionObj[0]) {protectionObj[0].remove()};
       const newProtect = sheetsToProtect[sheet].protect();
       newProtect.addEditors(editorEmails);   
     };           
   };

    /**
      * Sets protections for a range, allowing users from a list of emails to 
            edit those ranges. Note that this removes the existing set of 
            editors and replaces them with those from the editorEmails array!
      * @param {Array<string>} editorEmails An array of editors to grant access 
          to edit ranges.      
      * @param {Range/Array<Range>} rangesToProtect A range or array of 
          specific ranges to apply protections to. If this is undefined, 
          protections are applied to the entire sheet.
      */
    setRangeProtections(editorEmails, rangesToProtect){
        rangesToProtect = convertToArray(rangesToProtect);
        editorEmails = convertToArray(editorEmails);

        for (let rangeToProtect in rangesToProtect){
          const sheetOfRangeToProtect = 
              rangesToProtect[rangeToProtect].getSheet();
          
          const existingRangeProtections = 
              sheetOfRangeToProtect.getProtections(
                  SpreadsheetApp.ProtectionType.RANGE);
          let existingRangeProtectionEditors = [];
          let existingRangeProtectionsA1 = [];
          
          for (let existingRangeProtection in existingRangeProtections) {
            existingRangeProtectionEditors.push(
                existingRangeProtections[existingRangeProtection].getEditors());
            existingRangeProtectionsA1.push(
                existingRangeProtections[existingRangeProtection]
                .getRange().getA1Notation());
          };     
          
          const indexOfRange = existingRangeProtectionsA1.indexOf(
              rangesToProtect[rangeToProtect].getA1Notation());          
          
          if (existingRangeProtectionEditors[indexOfRange]) {
            if (existingRangeProtectionEditors[indexOfRange]
                .concat().sort().toString() !== 
                    editorEmails.concat().sort().toString()) {     
            existingRangeProtections[indexOfRange].remove();
            this.replaceRangeProtections(
                rangesToProtect[rangeToProtect], editorEmails);              
          };                      
          } else {
            this.replaceRangeProtections(
                rangesToProtect[rangeToProtect], editorEmails);
          };                
        };
    };      
    
   /**
     * Uses the parent sheet as a template to update all sheets in the given 
          folderID at once.
     * @param {String} folderID The ID for the Drive folder where Spreadsheets 
          will be set.   
     * @param {Array<Sheet>} sourceSheets The sheets that will be copied into 
          every Spreadsheet in the given folder. If given as Array, will 
          linearly loop through each sheet.  
     * @param {Array<Range>/Range=} ranges Optional Range or array of Ranges, 
          used to determine which cells will be updated. If no range is given, 
          the entire sheet will be updated.    
     * @param {Array<Protection>/Protection=} protectionsObj Optional 
            Protection or Array of Protections, used to update and replace
            the existing protections in the target Range/Sheets.
     */   
   updateSheetsAndRanges(
     folderID, sourceSheets, ranges, protectionsObj) {
    const IDList = sheetsToDriveInteractionsObj.getSpreadsheetsInFolder(folderID);
    const COPY_PREFIX = 'Copy of ';

    sourceSheets = convertToArray(sourceSheets);
    ranges = convertToArray(ranges);
    protectionsObj = convertToArray(protectionsObj);
     
    for (let ID in IDList) {       
      for (let sheet in sourceSheets) {
        let sourceSheetName = sourceSheets[sheet].getName();        
        let copiedSourceSheet = sourceSheets[sheet].copyTo(IDList[ID]);
        let sheetToUpdate = IDList[ID].getSheetByName(sourceSheetName);        
        
        if (!this.sheetNameExistsInSpreadsheet(sourceSheetName, IDList[ID])) {
          this.removeCopyOfPrefixFromSheets(
              IDList[ID].getSheetByName(COPY_PREFIX + sourceSheetName));
          sheetToUpdate = IDList[ID].getSheetByName(sourceSheetName);
          copiedSourceSheet = sheetToUpdate.copyTo(IDList[ID]);
        };        
      
        if (ranges){
          for (let range in ranges) {
            let copiedSourceSheetRange = 
                copiedSourceSheet.getRange(ranges[range].getA1Notation());
            let sheetToUpdateRange = 
                sheetToUpdate.getRange(ranges[range].getA1Notation());
            copiedSourceSheetRange.copyTo(sheetToUpdateRange);                
            };
        } else {
          let sourceSheetDuplicateRange = 
            copiedSourceSheet.getRange(1, 1, sourceSheets[sheet].getMaxRows(), 
            sourceSheets[sheet].getMaxColumns());
                                          
          if (sheetToUpdate.getMaxRows() > copiedSourceSheet.getMaxRows()) {
            let rowDifference = 
              sheetToUpdate.getMaxRows() - copiedSourceSheet.getMaxRows();
            sheetToUpdate.deleteRows(
              copiedSourceSheet.getMaxRows() + 1, rowDifference);
          };

          if (sheetToUpdate.getMaxColumns() > 
            copiedSourceSheet.getMaxColumns()) {
              let columnDifference = sheetToUpdate.getMaxColumns() - 
                copiedSourceSheet.getMaxColumns();
              sheetToUpdate.deleteColumns(
                copiedSourceSheet.getMaxColumns() + 1, columnDifference);
          };
          sourceSheetDuplicateRange.copyTo(sheetToUpdate.getRange(
            sourceSheetDuplicateRange.getA1Notation()));              
          };

        if (protectionsObj) {
          if (protectionsObj.length > 0){
          for (let protection in protectionsObj){          
            if (ranges){                   
              const targetRangeToProtect = sheetToUpdate.getRange(
                protectionsObj[protection].getRange().getA1Notation());             
              const editorEmails = protectionsObj[protection].getEditors();
              this.setRangeProtections(editorEmails, targetRangeToProtect);              
            } else {
                if (sheetToUpdate.getName() === 
                        protectionsObj[protection].getRange().getSheet().getName()){
                  const editorEmails = protectionsObj[protection].getEditors();
                  this.setSheetProtections(editorEmails, sheetToUpdate);
                };
            };
          };
        } else {
          if (ranges){
            let targetRangeToUnprotect = 
              sheetToUpdate.getRange(ranges[range].getA1Notation());
            this.unprotectRange(targetRangeToUnprotect);            
          } else {            
            let sheetProtection = 
              sheetToUpdate.getProtections(SpreadsheetApp.ProtectionType.SHEET);
            if (sheetProtection[0]){sheetProtection[0].remove()};            
          };
        };
        };
        IDList[ID].deleteSheet(copiedSourceSheet);
      };       
    };
  };
};

const spreadsheetFunctionsObj = new spreadsheetFunctions(); 
