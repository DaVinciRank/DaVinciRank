import { FolderUtils } from "./folderUtils";
import { ScoringUtils } from "./scoringUtils";
import { SpreadsheetUtils } from "./spreadsheetUtils";
import { Utils } from "./utils";
import { Slides } from "./slides";

/**
 * Runs when the Google Sheets document is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu("Science Olympiad Tournament Functions")
    .addItem("1. Create Only Event Tabs", "SpreadsheetUtils.duplicateProtectedSheet")
    .addItem("2. Create Event Spreadsheets", "ScoringUtils.createNewScoringSpreadsheets")
    .addItem("3. Create Grading Scoresheets", "SpreadsheetUtils.getTemplateFilesByEvent")
    .addItem("4. Share Scoring Folder with ES", "FolderUtils.shareScoringFoldersWithEmails")
    .addItem("5. Create Slides Presentation", "Slides.createOneSlidePerRow")
    .addToUi();
}

const util = new Utils();
const scoringUtils = new ScoringUtils();
const folderUtils = new FolderUtils();
const spreadsheetUtils = new SpreadsheetUtils();
const slides = new Slides();
