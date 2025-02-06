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
      .addSeparator()
      .addSubMenu(
        ui.createMenu('Debugging Only')
        .addItem("Enable Debug Logs", "enableDebugLogs")
        .addItem("Disable Debug Logs", "disableDebugLogs")
        .addItem("Delete Event Tabs", "deleteEventTabs")
        .addItem('Test: Send Many Logger', 'sendManyLogs')
    ).addToUi();
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

function deleteEventTabs() {
  showSidebar();
  // Add a ui alert to confirm deletion
  var userChoice = SpreadsheetApp.getUi().alert(
    'Are you sure you want to delete all event tabs?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (userChoice == SpreadsheetApp.getUi().Button.NO) {
    SpreadsheetApp.getUi().alert('Stopping delete operation.');
    return;
  }

  CacheLogger.appendLog("Deleting Event Tabs");
  const eventNames = SpreadsheetUtils.getEventNames();
  if (!eventNames || eventNames.length === 0) {
    CacheLogger.appendLog("No event names found to delete.");
    return;
  }

  const sheetsToDelete: GoogleAppsScript.Spreadsheet.Sheet[] = [];
  eventNames.forEach(eventName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eventName);
    if (sheet) {
      sheetsToDelete.push(sheet);
    }
  });

  if (sheetsToDelete.length > 0) {
    sheetsToDelete.forEach(sheet => {
      CacheLogger.appendLog(`Deleting sheet: ${sheet.getName()}`);
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    });
  }
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

function enableDebugLogs() {
  CacheLogger.setDebugMode(true);
}

function disableDebugLogs() {
  CacheLogger.setDebugMode(false);
}


// Define imports here so when built, they all combine into 1 js file for clasp push
const util = new Utils();
const scoringUtils = new ScoringUtils();
const folderUtils = new FolderUtils();
const spreadsheetUtils = new SpreadsheetUtils();
const slides = new Slides();
const cacheLogger = new CacheLogger();
