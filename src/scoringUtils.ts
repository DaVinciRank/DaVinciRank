import { Utils } from "./utils";
import { FolderUtils } from "./folderUtils";
import { SpreadsheetUtils } from "./spreadsheetUtils";
import { CacheLogger } from "./cacheLogger";
import { TournamentUtils } from "./tournamentUtils";

export class ScoringUtils {
  /**
   * Creates new scoring spreadsheets.
   */
  static createNewScoringSpreadsheets() {
    const startTime = new Date();
    const tournamentName = TournamentUtils.getTournamentNameParsed();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = spreadsheet.getSheetByName("Blank Score Sheet");
    if (!templateSheet) {
      SpreadsheetApp.getUi().alert(
        'Template sheet "Blank Score Sheet" not found.',
      );
      return;
    }

    const eventNames = SpreadsheetUtils.getEventNames();
    if (!eventNames || eventNames.length === 0) {
      return;
    }

    var teamNumbers = spreadsheet.getRangeByName("Team_Numbers")?.getValues();
    if (!teamNumbers || teamNumbers[0][0] == "") {
      SpreadsheetApp.getUi().alert(
        "You have not entered any team numbers. Please try again",
      );
      return;
    }

    var parentFolderId = FolderUtils.getParentFolderId();
    var scoreSheetFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
      parentFolderId,
      "Event Specific Score Sheets - " + tournamentName,
    );

    const highLowScoreWinsRange =
      spreadsheet.getRangeByName("HighLowScoreWins");
    if (!highLowScoreWinsRange) {
      SpreadsheetApp.getUi().alert('Named range "HighLowScoreWins" not found.');
      return;
    }
    const highLowScoreWins = highLowScoreWinsRange
      .getValues()
      .flat()
      .filter(String);

    eventNames.forEach((eventName, index) => {
      const startEventTime = new Date();
      CacheLogger.appendLog(
        `Creating scoring sheet for ${index + 1}/${eventNames.length}: ${eventName}`,
      );

      var spreadSheetName = eventName + " Event Scoring - " + tournamentName;
      var spreadSheetFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
        scoreSheetFolder,
        spreadSheetName,
      );
      var newSpreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet =
        SpreadsheetUtils.createNewSpreadSheetUnderSpecificFolder(
          spreadSheetFolder,
          spreadSheetName,
        );
      var newSpreadSheetSheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetUtils.duplicateProtectedSheetToNewSpreadsheet(
          templateSheet,
          newSpreadSheet,
          eventName,
        );
      Utils.moveRows(
        templateSheet,
        newSpreadSheetSheet,
        eventName,
        highLowScoreWins[index],
      );

      ScoringUtils.pasteLookupFormulasToSourceScoringSheets(
        spreadsheet,
        newSpreadSheet.getUrl(),
        eventName,
      );

      const endEventTime = new Date();
      CacheLogger.appendLog(
        `Total time taken for ${eventName}: ${(endEventTime.getTime() - startEventTime.getTime()) / 1000} seconds`,
        true,
      );
    });

    SpreadsheetApp.flush();

    const htmlOutput = HtmlService.createHtmlOutput(
      '<p>Click to view <a href="' +
        scoreSheetFolder.getUrl() +
        '" target="_blank">' +
        "Event ScoreSheets" +
        "</a></p>",
    )
      .setWidth(800)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(
      htmlOutput,
      "Created " + eventNames.length + " Event Sheets for Scoring",
    );
    const endTime = new Date();
    CacheLogger.appendLog(
      `Total time taken: ${(endTime.getTime() - startTime.getTime()) / 1000} seconds`,
      true,
    );
  }

  /**
   * Pastes lookup formulas to source scoring sheets.
   * @param {Spreadsheet} currentSheet - The current spreadsheet.
   * @param {string} newSheetUrl - The URL of the new sheet.
   * @param {string} eventName - The name of the event.
   */
  static pasteLookupFormulasToSourceScoringSheets(
    currentSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    newSheetUrl: string,
    eventName: string,
  ) {
    var scoreSheet = currentSheet.getSheetByName(eventName);
    if (!scoreSheet) {
      SpreadsheetApp.getUi().alert(
        "Sheet for event '" + eventName + "' does not exist.",
      );
      return;
    }
    var columns = ["C", "D", "E"];
    for (const i in columns) {
      var col = columns[i];
      scoreSheet
        .getRange(col + "2")
        .setFormula(
          '=IMPORTRANGE("' +
            newSheetUrl +
            '", "' +
            eventName +
            "!" +
            col +
            "2:" +
            col +
            '104")',
        );
    }
  }
}
