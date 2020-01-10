/** Module for performing functions related to Spreadsheet manipulation. 
      Includes the following functions:
  * @return {function} removeCopyOfPrefixFromSheets For a Sheet or array of 
      Sheet objects, Removes the "Copy of " prefix Google Sheets applies to 
      duplicated sheets.   
  * @return {function} sheetNameExistsInSpreadsheet Checks if a sheet with a 
      specific name already exists in a target spreadsheet.
  * @return {function} clearSheetsWithName Utilizes sheetNameExistsInSpreadsheet 
      to delete a sheet if its title exists in a spreadsheet.
  * @return {function} replaceRangePermissions Clears existing permissions for a 
      given range in a sheet, replacing them with the list of users in 
      NewEditors.
  * @return {function} setProtections Sets protections for a sheet and or/range, 
      allowing users from a list of emails to edit those ranges. 
      Note that this removes the existing set of editors!  
  * @return {function} updateSheetsAndRanges Uses the master sheet as a template 
      to update all sheets in the given folderID at once.
  */
 var spreadsheetFunctions_ = (function() {
  /**
   * For a Sheet, Removes the "Copy of " prefix Google Sheets applies to 
      duplicated sheets.   
    * @param {Sheet/Array<Sheet>} targetSheet The sheet that will have the 
      prefix removed. Takes either a single Sheet or an array of Sheet 
      objects. 
    */
  function removeCopyOfPrefixFromSheets(targetSheets) {    
    if (!Array.isArray(targetSheets)) {targetSheets = [targetSheets]};
    
    for (var targetSheet in targetSheets) {
      const REGEX = /^Copy\sof\s/;
      targetSheets[targetSheet].setName(
        targetSheets[targetSheet].getName().replace(REGEX,''));
    };
  };

  /**
    * Checks if a sheet with a specific name already exists in a target 
        spreadsheet.
    * @param {string} sheetName The name to check for in the target 
        spreadsheet's sheets. 
    * @param {Spreadsheet} targetSpreadsheet The spreadsheet that will be 
        searched for a sheet titled the sheetName string. 
    * @return {boolean} sheetExists True if a sheet with the sheetName exists 
        in the targetSpreadsheet, false if not.
    */
  function sheetNameExistsInSpreadsheet(sheetName, targetSpreadsheet) {
    var sheetExists = false;
    var targetSheets = targetSpreadsheet.getSheets();

    for (var i = 0; i < targetSheets.length; i++){
        if (targetSheets[i].getName() === sheetName) {
            sheetExists = true;
            return sheetExists;
        };
    };      
    return sheetExists;
  };

    /**
      * Utilizes sheetNameExistsInSpreadsheet to delete a sheet if its title 
          exists in a spreadsheet.
      * @param {string/Array<string>} sheetNamesToDelete String or array of 
          strings checked for matching titles with those in targetSpreadsheet; 
          if matching titles are found, targetSpreadsheet is deleted.
      * @param {Spreadsheet} targetSpreadsheet The spreadsheet that will be 
          searched for a sheet with the sheetNamesToDelete strings. 
      */
    function clearSheetsWithName(sheetNamesToDelete, targetSpreadsheet) {
      if (!Array.isArray(sheetNamesToDelete)) {sheetNamesToDelete = 
        [sheetNamesToDelete]};
      
      for (var sheetName in sheetNamesToDelete){
          if (sheetNameExistsInSpreadsheet(sheetNamesToDelete[sheetName], 
              targetSpreadsheet)){
                targetSpreadsheet.deleteSheet(
                  targetSpreadsheet.getSheetByName(
                    sheetNamesToDelete[sheetName]));
          };
      };
    };
    
    /**
     * Clears existing protections for a given range in a sheet, replacing them 
        with the list of users in NewEditors.
     * @param {Range/Array<Range>} ranges The range that will have its existing 
        editors removed, replaced with those in newEditors.
     * @param {Array<User>} newEditors An array of User objects representing the 
        new editors for the range.}
     */
    function replaceRangeProtections(ranges, newEditors){
        if (!Array.isArray(ranges)) {ranges = [ranges]};
        
        try {
          for (var range in ranges){
            var protectionObject = ranges[range].protect();
            protectionObject.addEditor(Session.getEffectiveUser())
                            .removeEditors(protectionObject.getEditors())
                            .addEditors(newEditors);  
          };
        } catch(error){
          UiFunctions.displayAlert('Cannot apply protections for admin.');
          return;
        };
    };

    /**
      * Sets protections for a sheet or range, allowing users from a list of 
          emails to edit those ranges. Note that this removes the existing set 
          of editors!
      * @param {Array<string>} editorEmails An array of editors to grant access 
          to edit ranges.
      * @param {Sheet/Array<Sheet>} sheetsToProtect A sheet or array of Sheets 
          to apply protections to.
      * @param {Range/Array<Range>=} rangesToProtect An optional array of 
          specific ranges to apply protections to. If this is null, protections 
          are applied to the entire sheet.
      */
    function setProtections(editorEmails, sheetsToProtect, rangesToProtect) {
      if ((sheetsToProtect) && (!rangesToProtect)){
        if (!Array.isArray(sheetsToProtect)) {sheetsToProtect = 
          [sheetsToProtect]};
        try {
          for (var sheet in sheetsToProtect){    
            var existingSheetProtection = sheetsToProtect[sheet]
              .getProtections(SpreadsheetApp.ProtectionType.SHEET);
              
            if (existingSheetProtection[0]) {
              if (JSON.stringify(
                existingSheetProtection[0].getEditors()) !== 
                  JSON.stringify(editorEmails)) {
                  sheetsToProtect[sheet].protect()
                                        .addEditor(Session.getEffectiveUser())
                                        .removeEditors(
                                          sheetsToProtect[sheet].protect()
                                          .getEditors())
                                        .addEditors(editorEmails);
              }
            } else {
                sheetsToProtect[sheet].protect()
                                      .addEditor(Session.getEffectiveUser())
                                      .removeEditors(
                                        sheetsToProtect[sheet].protect()
                                        .getEditors())
                                      .addEditors(editorEmails);
            };
          };      
        } catch (error){
          return;
        };
      };
      
      if (rangesToProtect){
        if (!Array.isArray(sheetsToProtect)) {sheetsToProtect = 
          [sheetsToProtect]};
        if (!Array.isArray(rangesToProtect)) {rangesToProtect = 
          [rangesToProtect]};
      
        for (var sheet in sheetsToProtect){        
          if (!Array.isArray(rangesToProtect)) {rangesToProtect = 
            [rangesToProtect]};
          var existingRangeProtections = sheetsToProtect[sheet]
            .getProtections(SpreadsheetApp.ProtectionType.RANGE);
          var existingRangeProtectionEditors = [];
          var existingRangeProtectionsA1 = [];
          
          for (var existingRangeProtection in existingRangeProtections) {
            existingRangeProtectionEditors.push(
              existingRangeProtections[existingRangeProtection].getEditors());
            existingRangeProtectionsA1.push(
              existingRangeProtections[existingRangeProtection]
                .getRange().getA1Notation());
          };
          
          for (var rangeToProtect in rangesToProtect) {
            var indexOfRange = existingRangeProtectionsA1.indexOf(
              rangesToProtect[rangeToProtect]);
            
            if (existingRangeProtectionEditors[indexOfRange]) {
              if (existingRangeProtectionEditors[indexOfRange]
                  .concat().sort() !== editorEmails.concat().sort()) {              
                    if (existingRangeProtections[indexOfRange]) {
                      existingRangeProtections[indexOfRange].remove();
                    };
                    replaceRangeProtections(sheetsToProtect[sheet]
                      .getRange(rangesToProtect[rangeToProtect]), 
                        editorEmails);                          
              };                      
            } else {
              replaceRangeProtections(sheetsToProtect[sheet]
                .getRange(rangesToProtect[rangeToProtect]), 
                editorEmails);
            };                
          };
        };          
      };
    };      
    
   /**
     * Uses the master sheet as a template to update all sheets in the given 
          folderID at once.
     * @param {String} folderID The ID for the Drive folder where Spreadsheets 
          will be set.   
     * @param {Array/Object} sourceSheets The sheets that will be copied into 
          every Spreadsheet in the given folder. If given as Array, will 
          linearly loop through each sheet. If given as object, will loop 
          through every key to retrieve sheets. 
     * @param {Array<Range>/Object=} ranges Optional array of Range objects, 
          used to determine which cells will be updated. If no range is given, 
          the entire sheet will be updated.    
     * @param {Array<Range>} rangesToProtect Optional array of ranges to assign 
          protections. If both it and editorEmails are null, no protections
          are assigned. If it is null and editorEmails is not, protections will 
          be applied to the whole sheet. If this has content but editorEmails
          does not, nothing will be done.
     * @param {Array<string>=} editorEmails Optional array of emails of users to 
          provide editing protections for. These users will be able to edit any 
          ranges given by rangesToProtect. 
     */   
   function updateSheetsAndRanges(
     folderID, sourceSheets, ranges, rangesToProtect, editorEmails) {
    var IDList = userSheetsModule_.getSpreadsheetsInFolder(folderID);
    const COPY_PREFIX = 'Copy of ';
    
    if (!Array.isArray(sourceSheets)) {sourceSheets = [sourceSheets]};
    if ((ranges) && (!Array.isArray(ranges))) {ranges = [ranges]};
    if ((rangesToProtect) && (!Array.isArray(rangesToProtect))) {
      rangesToProtect = [rangesToProtect]
    };
    if ((editorEmails) && (!Array.isArray(editorEmails))) {
      editorEmails = [editorEmails]
    };
        
    for (var ID in IDList) {        
      for (var sheet in sourceSheets) {
        var sourceSheetName = sourceSheets[sheet].getName();
        var copiedSourceSheet = sourceSheets[sheet].copyTo(IDList[ID]);
        var sheetToUpdate = IDList[ID].getSheetByName(sourceSheetName);
        
        if (!sheetNameExistsInSpreadsheet(sourceSheetName, IDList[ID])) {
          removeCopyOfPrefixFromSheets(
            IDList[ID].getSheetByName(COPY_PREFIX + sourceSheetName));
          sheetToUpdate = IDList[ID].getSheetByName(sourceSheetName);
          copiedSourceSheet = sheetToUpdate.copyTo(IDList[ID]);
        };

        if (ranges) {
          for (var range in ranges) {
            var copiedSourceSheetRange = 
              copiedSourceSheet.getRange(ranges[range]);
            var sheetToUpdateRange = 
              sheetToUpdate.getRange(ranges[range]);
            copiedSourceSheetRange.copyTo(sheetToUpdateRange);                
            };
        } else {
          var sourceSheetDuplicateRange = 
            copiedSourceSheet.getRange(1, 1, sourceSheets[sheet].getMaxRows(), 
            sourceSheets[sheet].getMaxColumns());
                                          
          if (sheetToUpdate.getMaxRows() > copiedSourceSheet.getMaxRows()) {
            var rowDifference = 
              sheetToUpdate.getMaxRows() - copiedSourceSheet.getMaxRows();
            sheetToUpdate.deleteRows(
              copiedSourceSheet.getMaxRows() + 1, rowDifference);
          };

          if (sheetToUpdate.getMaxColumns() > 
            copiedSourceSheet.getMaxColumns()) {
              var columnDifference = sheetToUpdate.getMaxColumns() - 
                copiedSourceSheet.getMaxColumns();
              sheetToUpdate.deleteColumns(
                copiedSourceSheet.getMaxColumns() + 1, columnDifference);
          };
          sourceSheetDuplicateRange.copyTo(sheetToUpdate.getRange(
            sourceSheetDuplicateRange.getA1Notation()));              
          };
        
        if (rangesToProtect && editorEmails) {
            setProtections(editorEmails, sheetToUpdate, rangesToProtect);
        } else if (editorEmails) {
            setProtections(editorEmails, sheetToUpdate);
        };
        
        IDList[ID].deleteSheet(copiedSourceSheet);
      };       
    };
};

return {
    removeCopyOfPrefixFromSheets: removeCopyOfPrefixFromSheets,
    sheetNameExistsInSpreadsheet: sheetNameExistsInSpreadsheet,
    clearSheetsWithName: clearSheetsWithName,
    replaceRangeProtections: replaceRangeProtections,
    setProtections: setProtections,
    updateSheetsAndRanges: updateSheetsAndRanges
}
})();