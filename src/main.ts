import { FolderUtils } from "./folderUtils";
import { ScoringUtils } from "./scoringUtils";
import { SpreadsheetUtils } from "./spreadsheetUtils";
import { Utils } from "./utils";
import { Slides } from "./slides";
import { CacheLogger } from "./cacheLogger";

/**
 * Runs when the Google Sheets document is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu("Science Olympiad Tournament Functions")
    .addItem("1. Create Only Event Tabs", "createEventTabs")
    .addItem("2. Create Event Spreadsheets", "createEventSpreadsheets")
    .addItem("3. Create Grading Scoresheets", "createGradingScoresheets")
    .addItem("4. Share Scoring Folder with ES", "shareScoringFoldersWithEmails")
    .addItem("5. Create Slides Presentation", "createOneSlidePerRow")
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Test: Send Many Logger', 'sendManyLogs')
    .addToUi();
}

function createEventTabs() {
  showSidebar();
  SpreadsheetUtils.duplicateProtectedSheet();
}

function createEventSpreadsheets() {
  showSidebar();
  ScoringUtils.createNewScoringSpreadsheets();
}

function createGradingScoresheets() {
  showSidebar();
  SpreadsheetUtils.getTemplateFilesByEvent();
}

function shareScoringFoldersWithEmails() {
  showSidebar();
  FolderUtils.shareScoringFoldersWithEmails();
}

function createOneSlidePerRow() {
  showSidebar();
  Slides.createOneSlidePerRow();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Log Messages");
  SpreadsheetApp.getUi().showSidebar(html);
}

function sendManyLogs() {
  showSidebar();
  for (let i = 1; i <= 20; i++) {
  (function (logNumber) {
      Utilities.sleep(1000);
      appendLog("Log " + logNumber);
  })(i);
  }
}

function appendLog(message: string): string {
  return CacheLogger.appendLog(message);
}

function getLogs(): string {
  return CacheLogger.getLogs();
}

function clearLogs(): string {
  return CacheLogger.clearLogs();
}


// Define imports here so when built, they all combine into 1 js file for clasp push
const util = new Utils();
const scoringUtils = new ScoringUtils();
const folderUtils = new FolderUtils();
const spreadsheetUtils = new SpreadsheetUtils();
const slides = new Slides();
const cacheLogger = new CacheLogger();
