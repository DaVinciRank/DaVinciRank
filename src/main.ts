// main.gs

// Import other files
// @import "utils.gs"
// @import "spreadsheetUtils.gs"
// @import "folderUtils.gs"
// @import "scoringUtils.gs"
// @import "slides.gs"

/**
 * Runs when the Google Sheets document is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu("Science Olympiad Tournament Functions")
    .addItem("1. Create Only Event Tabs", "duplicateProtectedSheet")
    .addItem("2. Create Event Spreadsheets", "createNewScoringSpreadsheets")
    .addItem("3. Create Grading Scoresheets", "getTemplateFilesByEvent")
    .addItem("4. Share Scoring Folder with ES", "shareScoringFoldersWithEmails")
    .addItem("5. Create Slides Presentation", "createOneSlidePerRow")
    .addToUi();
}
