/** Google Slides Lesson Planning Automation
 * Creates a presentation from a template and a Google Sheet.
 * Copies a Google Slides template and replaces placeholder text with
 * content from a specified Google Sheet. Also, adds speaker notes from
 * another sheet.
 * 
 * Globals:
 * - TEMPLATE_ID: ID of the Google Slides template.
 * - TARGET_FOLDER_ID: ID of the folder where the new presentation will be saved.
 * 
 * @param {none} No parameters.
 * @return {void} Does not return a value.
 */
// Global Variables
const TEMPLATE_ID = "YOUR_TEMPLATE_ID_HERE";
const TARGET_FOLDER_ID = "YOUR_TARGET_FOLDER_ID_HERE"; // Folder changes weekly to keep things organized
var newPresentation;

function createPresentation() {
  try {
    // IDs and URLs
    var spreadsheetId = "YOUR_SPREADSHEET_ID_HERE"; // Google Sheet ID
    var sheetName = "Slide Content"; // Replace with the actual sheet name. This sheet holds the template variables which is separate from the speaker notes.

    // Get the template presentation as a file
    var templateFile = DriveApp.getFileById(TEMPLATE_ID);
    if (!templateFile) throw new Error("Template file not found.");

    // Create a copy of the template presentation in the target folder
    var targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    if (!targetFolder) throw new Error("Target folder not found.");
    var newFile = templateFile.makeCopy("NAME_YOUR_PRESENTATION", targetFolder); // Name presentation

    // Open the new presentation
    newPresentation = SlidesApp.openById(newFile.getId());

    // Open the Google Sheet
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

    // Read data from the spreadsheet
    var values = sheet.getDataRange().getValues();
    if (values.length == 0) throw new Error("No data found in the sheet.");

    // Replace template variables in the presentation with values
    values.forEach(function(row, index) {
      var templateVariable = row[0]; // First column contains variable names
      var templateValue = row[1]; // Assumes values are in the second column
      if (templateVariable && templateValue) {
        newPresentation.replaceAllText(templateVariable, templateValue);
      } else {
        Logger.log("Missing data in row " + (index + 1) + "; skipping this row.");
      }
    });

    // Log the ID of the new file to verify it's being created
    Logger.log("New File ID: " + newFile.getId());

    // Log the number of slides to verify the presentation is being opened correctly
    Logger.log("Number of slides in newPresentation: " + newPresentation.getSlides().length);

    // Add speaker notes to the slides
    addSpeakerNotesToSlides(newPresentation, spreadsheetId, "Weekly Notes", 2);//(Values are 2,5,8,11,14)
  } catch (e) {
    Logger.log("Error in createPresentation: " + e.message);
  }
}

/**
 * Adds speaker notes to slides based on content from a Google Sheet.
 * The function reads notes from a specified column and inserts them into the speaker notes section of each slide.
 * Assumes a header row and combines specific rows for custom note placement.
 *
 * @param {SlidesApp.Presentation} presentation - The presentation object to add notes to.
 * @param {string} spreadsheetId - The ID of the Google Spreadsheet containing the notes.
 * @param {string} sheetName - The name of the sheet within the Spreadsheet to read.
 * @param {number} notesColumn - The column index from which to read the speaker notes.
 */
function addSpeakerNotesToSlides(presentation, spreadsheetId, sheetName, notesColumn) {
  try {
    // Open the Google Sheet containing speaker notes
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' not found for speaker notes.");

    // Define the range, starting from the fourth row
    var firstRow = 4;
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(firstRow, notesColumn, lastRow - firstRow + 1, 1);

    // Read speaker notes data from the specified range
    var notes = range.getValues();
    if (notes.length == 0) throw new Error("No speaker notes found in the sheet.");

    // Log the number of notes fetched
    Logger.log("Number of notes fetched from the sheet: " + notes.length);

    // Get the slides of the presentation
    var slides = presentation.getSlides();

    // Iterate over the slides and add the speaker notes
    for (var i = 0; i < slides.length; i++) {
      var notesText = "";

      // Custom logic for combining notes for specific slides
      if (i == 2) { // For the 3rd slide, combine notes from row 6 and 7.
        notesText = combineSpecificNotes(notes, 2, 3);
      } else {
        // Calculate the index for notes array
        var notesIndex = i < 2 ? i : i + 1; // Adjust index after slide 3
        if (notesIndex < notes.length && notes[notesIndex]) {
          notesText = notes[notesIndex][0] ? notes[notesIndex][0] : "";
        }
      }

      if (notesText) {
        var notesPage = slides[i].getNotesPage();
        var notesTextRange = notesPage.getSpeakerNotesShape().getText();
        notesTextRange.setText(notesText);
        Logger.log("Added note to slide " + (i + 1) + ": " + notesText);
      } else {
        Logger.log("No note available for slide " + (i + 1));
      }
    }
  } catch (e) {
    Logger.log("Error in addSpeakerNotesToSlides: " + e.message);
  }
}

/**
 * Combines notes for specific slides based on given indices.
 * @param {Array} notes - The array containing notes.
 * @param {number} firstIndex - The index of the first note to combine.
 * @param {number} secondIndex - The index of the second note to combine.
 * @return {string} The combined notes text.
 */
function combineSpecificNotes(notes, firstIndex, secondIndex) {
  var firstNote = notes[firstIndex] && notes[firstIndex][0] ? notes[firstIndex][0] + "\n" : "";
  var secondNote = notes[secondIndex] && notes[secondIndex][0] ? notes[secondIndex][0] : "";
  return firstNote + secondNote;
}




