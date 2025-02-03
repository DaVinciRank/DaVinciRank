import { Utils } from "./utils";
import { FolderUtils } from "./folderUtils";
import { SpreadsheetUtils } from "./spreadsheetUtils";

export class ScoringUtils {
  /**
   * Creates new scoring spreadsheets.
   */
  static createNewScoringSpreadsheets() {
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = currentSheet.getSheetByName("Blank Score Sheet");
    if (!templateSheet) {
      SpreadsheetApp.getUi().alert("`Blank Score Sheet` sheet does not exist.");
      return;
    }

    var range = currentSheet.getRangeByName("Events");
    if (!range) {
      SpreadsheetApp.getUi().alert("The named range 'Events' does not exist.");
      return;
    }
    var values = range.getValues();
    var sNames = values.flat().filter(function (cell) {
      return cell !== "";
    });

    var teamNumbers = currentSheet.getRangeByName("Team_Numbers")?.getValues();
    if (!teamNumbers || teamNumbers[0][0] == "") {
      SpreadsheetApp.getUi().alert(
        "You have not entered any team numbers. Please try again",
      );
      return;
    }

    var parentFolderId = FolderUtils.getParentFolderId();
    var scoreSheetFolderId = FolderUtils.createFolderUnderRootFolder(
      parentFolderId,
      Utils.getTournamentNameParsed() + " - Event Specific Score Sheets",
    );

    for (const j in sNames) {
      var eventName = sNames[j];
      var spreadSheetName =
        eventName + " Event Scoring - " + Utils.getTournamentNameParsed();
      var spreadSheetFolderId = FolderUtils.createFolderUnderRootFolder(
        scoreSheetFolderId,
        spreadSheetName,
      );
      var spreadSheetId = SpreadsheetUtils.createNewSpreadSheetUnderSpecificFolder(
        spreadSheetFolderId,
        spreadSheetName,
      );
      var newSheet = SpreadsheetUtils.duplicateProtectedSheetToNewSpreadsheet(
        templateSheet,
        spreadSheetId,
        eventName,
      );
      Utils.moveRows(templateSheet, newSheet, eventName);
      ScoringUtils.pasteLookupFormulasToSourceScoringSheets(
        currentSheet,
        SpreadsheetApp.openById(spreadSheetId).getUrl(),
        eventName,
      );
    }

    const htmlOutput = HtmlService.createHtmlOutput(
      '<p>Click to view <a href="' +
        DriveApp.getFolderById(scoreSheetFolderId).getUrl() +
        '" target="_blank">' +
        "Event ScoreSheets" +
        "</a></p>",
    )
      .setWidth(800)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(
      htmlOutput,
      "Created " + sNames.length + " Event Sheets for Scoring",
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