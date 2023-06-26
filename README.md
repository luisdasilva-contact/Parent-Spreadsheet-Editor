### Parent Spreadsheet Editor

![Artwork for Parent Spreadsheet Editor](https://cdnb.artstation.com/p/assets/images/images/064/354/539/large/luis-dasilva-title.jpg?1687752445)

 A program for those in need of templated Sheets files that can be updated automatically, across a large number of users. For example, those who lead data entry teams in their organization can create a parent spreadsheet, and use this program to apply formula or sheet changes to their teammates with a single button press, rather than having to manually copy/paste changes for every single user. Use of this program will not interrupt teammates' workflows; as soon as the user clicks a button to apply updates, changes will be applied from the parent spreadsheet to the other sheets without any further input needed.

##### Instructions:
1. In a Google Sheets file, go to Tools --> Script Editor. 
2. Copy and paste each of the .gs files included in this repository. Click Save or hit CMD + S/CTRL + S upon completion. The file can be given any name. Close the Script Editor window.
3. The program will automatically run upon refreshing the Spreadsheet's page, appearing as the final tab in the tool bar.
4. Set information for the program by going to Update User Sheets --> Initialize Sheets, and clicking the options to set values. The options are as follows:
	* Set User Names: (Required) Set the names of the Spreadsheets you'll create.
	* Set Admin List: (Optional) Set the emails for who will have access to apply edits to the sheets once created.
	* Set Folder ID: (Required) The ID for the Drive folder Spreadsheets will be managed in (the ID can be found by going to a folder on Drive, and getting the generated value from the URL. For example, in "drive.google.com/drive/folders/12345", "12345" is the ID.     
	* Set Title Append: (Optional) Text to append to the names of the Spreadsheets. 
Note that these values persist, even after the Spreadsheet is closed. If you'd like to clear these properties at any time, go to Update User Sheets --> Initialize Sheets --> Clear Properties.
5. After values have been set, click Update User Sheets --> Initialize Sheets --> Create Sheets. The script will run, and your Spreadsheets will be created!
6. After you've made changes to your Spreadsheet and want them applied to all the duplicated sheets, see the Update Sheets menu under Update User Sheets. Options are broken down by size: 	
	* Update Current Range: Used if the user has a cell in a Named Range as their key selection. If this is clicked, only the Named Range is updated. If the cell is not part of a Named Range, the whole sheet is updated. If the sheet does not yet exist in any of the duplicates, it is created.
	* Update Current Sheet: Updates the current sheet. If it does not exist in any of the duplicated sheets, it is created.
	* Update All Sheets: Updates all sheets in the Spreadsheet. If any do not exist in any of the duplicates, they are created.
