import { FolderUtils } from "./folderUtils";
import { ScoringUtils } from "./scoringUtils";
import { SpreadsheetUtils } from "./spreadsheetUtils";
import { Utils } from "./utils";
import { Slides } from "./slides";
import { CacheLogger } from "./cacheLogger";
import { Constants } from "./constants";

/**
 * Runs when the Google Sheets document is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu("Science Olympiad Tournament Functions")
    .addItem("1. First time setup", "tournamentSetup")
    .addItem("2. Create Only Event Tabs", "createEventTabs")
    .addItem("3. Create Event Spreadsheets", "createEventSpreadsheets")
    .addItem("4. Create Grading Scoresheets", "createGradingScoresheets")
    .addItem("5. Share Scoring Folder with ES", "shareScoringFoldersWithEmails")
    .addItem("6. Create Slides Presentation", "createOneSlidePerRow")
    .addItem("Show Sidebar", "showSidebar")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Debugging Only")
        .addItem("Set Default Tournament Values", "setDefaultTournamentValues")
        .addItem("Delete Event Tabs", "deleteEventTabs")
        .addItem("Test: Send Many Logger", "sendManyLogs"),
    )
    .addToUi();
  closeSidebar();
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

function setDefaultTournamentValues() {
  var userChoice = SpreadsheetApp.getUi().alert(
    "Are you sure you want to set all tournament named ranges to default?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO,
  );
  if (userChoice == SpreadsheetApp.getUi().Button.NO) {
    return;
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet
    .getRangeByName(Constants.TOURNAMENT_NAME_RANGE_NAME)
    ?.setValue(Constants.DEFAULT_TOURNAMENT_NAME);
  spreadsheet
    .getRangeByName(Constants.TOURNAMENT_DATE_RANGE_NAME)
    ?.setValue(Constants.DEFAULT_TOURNAMENT_DATE);
  spreadsheet
    .getRangeByName(Constants.LOCATION_RANGE_NAME)
    ?.setValue(Constants.DEFAULT_TOURNAMENT_LOCATION);
  spreadsheet
    .getRangeByName(Constants.DIVISION_RANGE_NAME)
    ?.setValue(Constants.DEFAULT_TOURNAMENT_DIVISION);
  spreadsheet
    .getRangeByName(Constants.NUMBER_OF_TEAMS_RANGE_NAME)
    ?.setValue(Constants.DEFAULT_NUMBER_OF_TEAMS);
}

function deleteEventTabs() {
  showSidebar();
  // Add a ui alert to confirm deletion
  var userChoice = SpreadsheetApp.getUi().alert(
    "Are you sure you want to delete all event tabs?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO,
  );
  if (userChoice == SpreadsheetApp.getUi().Button.NO) {
    SpreadsheetApp.getUi().alert("Stopping delete operation.");
    return;
  }

  CacheLogger.appendLog("Deleting Event Tabs");
  const eventNames = SpreadsheetUtils.getEventNames();
  if (!eventNames || eventNames.length === 0) {
    CacheLogger.appendLog("No event names found to delete.");
    return;
  }

  const sheetsToDelete: GoogleAppsScript.Spreadsheet.Sheet[] = [];
  eventNames.forEach((eventName) => {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eventName);
    if (sheet) {
      sheetsToDelete.push(sheet);
    }
  });

  if (sheetsToDelete.length > 0) {
    sheetsToDelete.forEach((sheet) => {
      CacheLogger.appendLog(`Deleting sheet: ${sheet.getName()}`);
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    });
  }
}

function showSidebar(force: boolean = false) {
  const cache = CacheService.getDocumentCache();
  var sidebarStatus: string | null = "on";
  if (cache) {
    sidebarStatus = cache.get("sidebar");
  }
  if (sidebarStatus !== "on" || force) {
    var html =
      HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
        "Tournament Manager",
      );
    SpreadsheetApp.getUi().showSidebar(html);
  }
  openSidebar();
}

function openSidebar() {
  const cache = CacheService.getDocumentCache();
  if (cache) {
    cache.put("sidebar", "on", 11);
  }
}

function closeSidebar() {
  const cache = CacheService.getDocumentCache();
  if (cache) {
    cache.put("sidebar", "off", 11);
  }
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

function toggleDebugMode(enable: boolean): boolean {
  CacheLogger.setDebugMode(enable);
  return CacheLogger.isDebugMode();
}

function isDebugMode(): boolean {
  return CacheLogger.isDebugMode();
}

function enableDebugLogs() {
  CacheLogger.setDebugMode(true);
}

function disableDebugLogs() {
  CacheLogger.setDebugMode(false);
}

function tournamentSetup() {
  showSidebar();

  // Get current tournament info
  const info = getTournamentInfo();

  // Check each field and prompt if not populated
  for (const item of info) {
    if (!item.populated) {
      promptForMissingInfo(item.label);
    }
  }
}

function getTournamentInfo() {
  try {
    // Get values from named ranges
    const tournamentName = Utils.getTournamentName();
    const tournamentDate = Utils.getTournamentDate();
    const tournamentLocation = Utils.getTournamentLocation();
    const division = Utils.getTournamentDivision();
    const numTeams = Utils.getNumberOfTeams();

    const scoresheetFolder = Utils.getScoreSheetFolder(false);
    const templateFolder = Utils.getTemplateFolder(false);

    // Return object with specific order (order matters for display)
    return [
      {
        label: "Tournament Name",
        value: tournamentName,
        populated: Boolean(tournamentName),
        isLink: false,
      },
      {
        label: "Tournament Date",
        value: tournamentDate
          ? Utilities.formatDate(
              new Date(tournamentDate),
              Session.getScriptTimeZone(),
              "MMMM d, yyyy",
            )
          : "",
        populated: Boolean(tournamentDate),
        isLink: false,
      },
      {
        label: "Tournament Location",
        value: tournamentLocation,
        populated: Boolean(tournamentLocation),
        isLink: false,
      },
      {
        label: "Division",
        value: division,
        populated: Boolean(division),
        isLink: false,
      },
      {
        label: "Number of Teams",
        value: numTeams,
        populated: Boolean(numTeams),
        isLink: false,
      },
      {
        label: "Scoresheet Folder",
        value: scoresheetFolder ? scoresheetFolder.getUrl() : "",
        populated: Boolean(scoresheetFolder),
        isLink: true,
      },
      {
        label: "Template Folder",
        value: templateFolder ? templateFolder.getUrl() : "",
        populated: Boolean(templateFolder),
        isLink: true,
      },
    ];
  } catch (error) {
    console.error("Error in getTournamentInfo:", error);
    return [
      { label: "Tournament Name", value: "", populated: false, isLink: false },
      { label: "Tournament Date", value: "", populated: false, isLink: false },
      {
        label: "Tournament Location",
        value: "",
        populated: false,
        isLink: false,
      },
      { label: "Division", value: "", populated: false, isLink: false },
      { label: "Number of Teams", value: "", populated: false, isLink: false },
      { label: "Scoresheet Folder", value: "", populated: false, isLink: true },
      { label: "Template Folder", value: "", populated: false, isLink: true },
    ];
  }
}

function promptForMissingInfo(field: string) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let prompt: string = "",
    range: GoogleAppsScript.Spreadsheet.Range | null = null;

  switch (field) {
    case "Tournament Name":
      prompt = "Please enter the tournament name:";
      range = spreadsheet.getRangeByName(Constants.TOURNAMENT_NAME_RANGE_NAME);
      break;
    case "Tournament Date":
      prompt = "Please enter the tournament date (MM/DD/YYYY):";
      range = spreadsheet.getRangeByName(Constants.TOURNAMENT_DATE_RANGE_NAME);
      break;
    case "Tournament Location":
      prompt =
        "Please enter the tournament location (ie school/university name):";
      range = spreadsheet.getRangeByName(Constants.LOCATION_RANGE_NAME);
      break;
    case "Division":
      prompt = "Please enter the division (B or C):";
      range = spreadsheet.getRangeByName(Constants.DIVISION_RANGE_NAME);
      break;
    case "Number of Teams":
      prompt = "Please enter the number of teams (used for scoring):";
      range = spreadsheet.getRangeByName(Constants.NUMBER_OF_TEAMS_RANGE_NAME);
      break;
    case "Scoresheet Folder":
      Utils.getScoreSheetFolder(true);
      return;
    case "Template Folder":
      Utils.getTemplateFolder(true);
      return;
  }

  if (!prompt || prompt == "" || !range) {
    return;
  }

  const result = Utils.showPrompt(prompt);

  if (result) {
    if (field === "Tournament Date") {
      // Parse date string to date object
      const date = new Date(result);
      if (!isNaN(date.getTime())) {
        range.setValue(date);
      } else {
        ui.alert("Invalid date format. Please use MM/DD/YYYY");
      }
    } else if (field === "Number of Teams") {
      const numTeams = parseInt(result);
      if (!isNaN(numTeams) && numTeams > 0) {
        range.setValue(numTeams);
      } else {
        ui.alert("Please enter a valid number");
      }
    } else {
      range.setValue(result);
    }
  }
}

function getEventScoringFiles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const events =
      ss.getRangeByName("Events")?.getValues().flat().filter(String) || [];
    const tournamentName = Utils.getTournamentNameParsed(false);
    const scoresheetFolder = Utils.getScoreSheetFolder(false);

    if (!scoresheetFolder) {
      console.error("Scoresheet folder not created yet:");
      return [];
    }

    return events.map((eventName) => {
      const scoringFileName = eventName + " Event Scoring - " + tournamentName;
      let fileUrl = null;

      const folder = FolderUtils.findFolderBySubstrings(
        scoresheetFolder,
        eventName,
      );

      try {
        if (folder) {
          const files = folder.getFilesByName(scoringFileName);
          if (files.hasNext()) {
            fileUrl = files.next().getUrl();
          }
        }
      } catch (e) {
        console.error(`Error finding file for ${eventName}: ${e}`);
      }

      return {
        name: eventName,
        fileUrl: fileUrl,
      };
    });
  } catch (error) {
    console.error("Error in getEventScoringFiles:", error);
    return [];
  }
}

/**
 * Retrieves a list of events along with whether each is finalized.
 */
function getEventsForCheckOff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsRange = ss.getRangeByName("Events");
  if (!eventsRange) return [];

  // Get a flat array of event names from the named range.
  const events = eventsRange.getValues().flat().filter(String);

  // Assume the finalized checkmarks are in the "Final Rankings" sheet.
  const finalRankSheet = ss.getSheetByName("Final Rankings");
  if (!finalRankSheet) {
    throw new Error("Final Rankings sheet not found.");
  }

  const result = events.map((eventName) => {
    // Use SpreadsheetUtils to find the row where this event is listed.
    const rowNum = SpreadsheetUtils.findCellRowWithTextInSheet(
      finalRankSheet,
      eventName,
    );
    let finalized = false;
    if (rowNum) {
      // Use the same logic as in your Slides class to check column "G"
      // (i.e. get the cell value at that row with an offset of 0).
      const cellValue = Slides.getCellValueByColumnRowAndOffset(
        finalRankSheet,
        "G",
        rowNum,
        0,
      );
      finalized = cellValue === true;
    }
    return {
      name: eventName,
      finalized: finalized,
    };
  });

  return result;
}

/**
 * Marks a given event as finalized.
 * @param {string} eventId - The event identifier (in this case, the event name).
 */
function finalizeEvent(eventId: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalRankSheet = ss.getSheetByName("Final Rankings");
  if (!finalRankSheet) {
    throw new Error("Final Rankings sheet not found.");
  }

  // Find the row for the event using the event name.
  const rowNum = SpreadsheetUtils.findCellRowWithTextInSheet(
    finalRankSheet,
    eventId,
  );
  if (!rowNum) {
    throw new Error(
      `Event "${eventId}" not found in the Final Rankings sheet.`,
    );
  }

  finalRankSheet.getRange("G" + rowNum).setValue(true);
}

// Define imports here so when built, they all combine into 1 js file for clasp push
const util = new Utils();
const scoringUtils = new ScoringUtils();
const folderUtils = new FolderUtils();
const spreadsheetUtils = new SpreadsheetUtils();
const slides = new Slides();
const cacheLogger = new CacheLogger();
