import { Utils } from "./utils";
import { FolderUtils } from "./folderUtils";

export class SpreadsheetUtils {

  /**
   * Creates a new spreadsheet under a specific folder.
   * @param {string} folderId - The folder ID.
   * @param {string} spreadSheetName - The name of the spreadsheet.
   * @returns {string} - The ID of the new spreadsheet.
   */
  static createNewSpreadSheetUnderSpecificFolder(
    folderId: string,
    spreadSheetName: string,
  ): string {
    const folder: GoogleAppsScript.Drive.Folder =
      DriveApp.getFolderById(folderId);
    var existing_ss = folder.getFilesByName(spreadSheetName);
    if (existing_ss.hasNext()) {
      DriveApp.getFileById(existing_ss.next().getId()).setTrashed(true);
    }
    var ss = SpreadsheetApp.create(spreadSheetName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    return ss.getId();
  }

  /**
   * Copies a template to a spreadsheet.
   * @param {Sheet} templateSheet - The template sheet.
   * @param {string} spreadSheetId - The spreadsheet ID.
   * @param {string} sheetTabName - The name of the sheet tab.
   */
  static copyTemplateToSpreadsheet(
    templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    spreadSheetId: string,
    sheetTabName: string,
  ) {
    const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.openById(spreadSheetId);
    templateSheet.copyTo(spreadsheet).setName(sheetTabName);
  }

  /**
   * Duplicates a protected sheet to a new spreadsheet.
   * @param {Sheet} templateSheet - The template sheet.
   * @param {string} newSpreadsheetId - The ID of the new spreadsheet.
   * @param {string} eventName - The name of the event.
   * @returns {Sheet} - The duplicated sheet.
   */
  static duplicateProtectedSheetToNewSpreadsheet(
    templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    spreadSheetId: string,
    sheetTabName: string,
  ): GoogleAppsScript.Spreadsheet.Sheet {
    var ss = SpreadsheetApp.openById(spreadSheetId);

    // Create the new sheet
    var sheet2 = templateSheet.copyTo(ss).setName(sheetTabName);
    // Copy over all the permissions
    var p = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    var p2 = sheet2.protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());
    if (!p.isWarningOnly()) {
      p2.removeEditors(p2.getEditors().map((editor) => editor.getEmail()));
      p2.addEditors(p.getEditors().map((editor) => editor.getEmail()));
    }
    var ranges = p.getUnprotectedRanges();
    var newRanges = [];
    for (const range of ranges) {
      newRanges.push(sheet2.getRange(range.getA1Notation()));
    }
    p2.setUnprotectedRanges(newRanges);

    var blank_sheet = ss.getSheetByName("Sheet1");
    if (blank_sheet) {
      ss.deleteSheet(blank_sheet);
    }

    return sheet2;
  }

  /**
   * Duplicates the protected sheet for each event.
   * @param {string} eventName - The name of the event.
   */
  static duplicateProtectedSheet() {
    Utils.getTournamentNameParsed();

    const ss: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();
    const sheet: GoogleAppsScript.Spreadsheet.Sheet | null =
      ss.getSheetByName("Blank Score Sheet");
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Template sheet "Blank Score Sheet" not found.',
      );
      return;
    }

    var range = SpreadsheetApp.getActive().getRangeByName("Events");
    if (!range) {
      SpreadsheetApp.getUi().alert('Named range "Events" not found.');
      return;
    }
    var values = range.getValues();
    var sNames = values.flat().filter(function (cell) {
      return cell !== "";
    });

    var highLowScoreWinsRange =
      SpreadsheetApp.getActive().getRangeByName("HighLowScoreWins");
    if (!highLowScoreWinsRange) {
      SpreadsheetApp.getUi().alert('Named range "HighLowScoreWins" not found.');
      return;
    }
    var highLowScoreWins = highLowScoreWinsRange
      .getValues()
      .flat()
      .filter(function (cell) {
        return cell !== "";
      });

    for (const j in sNames) {
      // Remove the sheet if it already exists and then re-create it
      var cur_sheet = ss.getSheetByName(sNames[j]);
      if (cur_sheet) {
        ss.deleteSheet(cur_sheet);
      }
      // Create the new sheet
      const sheet2 = sheet.copyTo(ss).setName(sNames[j]);

      // Copy over the event name
      sheet2.getRange("L2:O2").setValue(sNames[j]);

      sheet2.getRange("L4:O4").setValue(highLowScoreWins[j]);
      sheet2.getRange("L5:O5").setValue(highLowScoreWins[j]);

      // Copy over all the permissions
      var p = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      var p2 = sheet2.protect();
      p2.setDescription(p.getDescription());
      p2.setWarningOnly(p.isWarningOnly());
      if (!p.isWarningOnly()) {
        p2.removeEditors(p2.getEditors().map((editor) => editor.getEmail()));
        p2.addEditors(p.getEditors().map((editor) => editor.getEmail()));
      }
      var ranges = p.getUnprotectedRanges();
      var newRanges = [];
      for (var i = 0; i < ranges.length; i++) {
        newRanges.push(sheet2.getRange(ranges[i].getA1Notation()));
      }
      p2.setUnprotectedRanges(newRanges);
    }
    SpreadsheetUtils.forceRefreshSheetFormulas("Master Scoresheet", 32);
    SpreadsheetApp.getUi().alert("Created event tabs");
  }

  /**
   * Forces the refresh of formulas in a specified range on a sheet.
   *
   * This function iterates over a specified range on a sheet and forces the refresh
   * of formulas in that range by temporarily clearing and then setting them back.
   * It ensures that formulas dependent on external data sources are updated.
   *
   * @param {string} sheetName - The name of the sheet where formulas need to be refreshed.
   * @param {number} maxColumns - The maximum number of columns in the range to refresh.
   */
  static forceRefreshSheetFormulas(sheetName: string, maxColumns: number) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = activeSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet with name ${sheetName} not found.`);
    }
    var range = sheet.getDataRange();
    const numCols: number = maxColumns;
    const numRows: number = range.getNumRows();
    const rowOffset: number = range.getRow();
    const colOffset: number = range.getColumn();

    // Change formulas then change them back to refresh it
    var originalFormulas = range.getFormulas();

    //Loop through each column and each row in the sheet
    //`row` and `col` are relative to the range, not the sheet
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        if (originalFormulas[row][col] != "") {
          range.getCell(row + rowOffset, col + colOffset).setFormula("");
        }
      }
    }
    SpreadsheetApp.flush();
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        if (originalFormulas[row][col] != "") {
          range
            .getCell(row + rowOffset, col + colOffset)
            .setFormula(originalFormulas[row][col]);
        }
      }
    }
    SpreadsheetApp.flush();
  }

  /**
   * Retrieves template files for each event and copies them into event-specific folders.
   * @param {string} tournamentName - The parsed full tournament name.
   */
  static getTemplateFilesByEvent(tournamentName: string) {
    tournamentName = Utils.getTournamentNameParsed();

    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = currentSheet.getSheetByName("Blank Score Sheet");

    var values = currentSheet.getRangeByName("Events")?.getValues();
    if (!values) {
      SpreadsheetApp.getUi().alert('Named range "Events" not found.');
      return;
    }
    var sNames = values.flat().filter(function (cell) {
      return cell !== "";
    });

    var parentFolderId = FolderUtils.getParentFolderId();
    var scoreSheetFolderId = FolderUtils.createFolderUnderRootFolder(
      parentFolderId,
      Utils.getTournamentNameParsed() + " - Event Specific Score Sheets",
    );

    var templateFolderId = FolderUtils.createFolderUnderRootFolder(
      parentFolderId,
      Utils.getTournamentNameParsed() + " - Template Files",
    );
    var allTemplateFiles = FolderUtils.getFilesUnderRootRolder(templateFolderId);
    if (allTemplateFiles.length == 0) {
      const htmlOutput = HtmlService.createHtmlOutput(
        '<p>Click to open <a href="' +
          DriveApp.getFolderById(templateFolderId).getUrl() +
          '" target="_blank">' +
          Utils.getTournamentNameParsed() +
          " - Template Files" +
          "</a></p>",
      )
        .setWidth(800)
        .setHeight(100);
      SpreadsheetApp.getUi().showModalDialog(
        htmlOutput,
        "You have not uploaded template scoring sheets, please do so",
      );
      return;
    }

    for (const j in sNames) {
      var eventName = sNames[j];
      Logger.log(eventName);

      var eventScoringFolderName =
        eventName + " Event Scoring - " + Utils.getTournamentNameParsed();
      var eventScoringFolderId = FolderUtils.createFolderUnderRootFolder(
        scoreSheetFolderId,
        eventScoringFolderName,
      );
      var eventScoringFolder = DriveApp.getFolderById(eventScoringFolderId);

      var templateFiles = FolderUtils.getTemplateFilesWithSubstring(
        eventName,
        allTemplateFiles,
      );
      for (const i in templateFiles) {
        var templateFile = templateFiles[i];

        // Copy the template file into event specific scoring folder
        // Also need to clean-up the name if needed
        var fileType = templateFile.getMimeType();

        if (fileType == "application/vnd.google-apps.spreadsheet") {
          var scoreSheetName =
            tournamentName +
            ": " +
            eventName +
            " - Scoresheet (Use this for grading)";
          FolderUtils.removeFileIfExists(eventScoringFolder, scoreSheetName);
          var copiedFile = templateFile.makeCopy(
            scoreSheetName,
            eventScoringFolder,
          );

          // Need to copy team names over to scoresheet
          if (templateSheet) {
            SpreadsheetUtils.copyTeamNames(templateSheet, copiedFile);
          } else {
            throw new Error('Template sheet "Blank Score Sheet" not found.');
          }

          var scoringSpreadSheetName =
            eventName + " Event Scoring - " + Utils.getTournamentNameParsed();

          // Add IMPORTRANGE into the scoring spreadsheet
          SpreadsheetUtils.pasteLookupFormulasToScoringSheets(
            eventScoringFolder,
            scoringSpreadSheetName,
            SpreadsheetApp.openById(copiedFile.getId()),
          );
        } else {
          FolderUtils.removeFileIfExists(eventScoringFolder, templateFile.getName());
          var copiedFile = templateFile.makeCopy(eventScoringFolder);
        }
      }
    }
  }

  /**
   * Copies team names from a template sheet to a new sheet.
   * @param {Sheet} templateSheet - The template sheet.
   * @param {File} newFile - The new file.
   */
  static copyTeamNames(
    templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    newFile: GoogleAppsScript.Drive.File,
  ) {
    const newSheet = SpreadsheetApp.openById(newFile.getId());
    const startingRow = SpreadsheetUtils.findCellRowWithTextInSpreadsheet(newSheet, "Team #");

    if (typeof startingRow === "number") {
      newSheet
        .getRange("B" + (startingRow + 1) + ":B" + (startingRow + 103))
        .setValues(templateSheet.getRange("Team_Numbers").getValues());
      newSheet
        .getRange("C" + (startingRow + 1) + ":C" + (startingRow + 103))
        .setValues(templateSheet.getRange("Schools").getValues());
      newSheet
        .getRange("D" + (startingRow + 1) + ":D" + (startingRow + 103))
        .setValues(templateSheet.getRange("Team_Names").getValues());
    } else {
      const startingRow = SpreadsheetUtils.findCellRowWithTextInSpreadsheet(
        newSheet,
        "Team Name and State",
      );
      if (typeof startingRow === "number") {
        newSheet
          .getRange("C" + (startingRow + 1) + ":C" + (startingRow + 103))
          .setValues(templateSheet.getRange("Team_Numbers").getValues());
      }
    }
  }

  /**
   * Finds the row number of a cell containing specific text in a spreadsheet.
   * @param {Spreadsheet} spreadsheet - The spreadsheet.
   * @param {string} textToFind - The text to find.
   * @returns {number|boolean} - The row number or false if not found.
   */
  static findCellRowWithTextInSpreadsheet(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    textToFind: string,
  ): number | boolean {
    var sheet = spreadsheet.getSheetByName("Scoring");
    if (!sheet) {
      sheet = spreadsheet.getSheetByName("Sheet1");
      if (!sheet) {
        return false;
      }
    }
    return SpreadsheetUtils.findCellRowWithTextInSheet(sheet, textToFind);
  }

  /**
   * Finds the row number of a cell containing specific text in a sheet.
   * @param {Sheet} sheet - The sheet.
   * @param {string} textToFind - The text to find.
   * @returns {number|boolean} - The row number or false if not found.
   */
  static findCellRowWithTextInSheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    textToFind: string,
  ): number | boolean {
    var textFinder = sheet.createTextFinder(textToFind);
    var matchedRanges = textFinder.findAll();

    for (const i in matchedRanges) {
      var range = matchedRanges[i];
      if (range.getColumn() < 5) {
        return range.getRow();
      }
    }

    return false;
  }

  /**
   * Finds the row and column number of a cell containing specific text.
   * @param {Spreadsheet} spreadsheet - The spreadsheet.
   * @param {string} textToFind - The text to find.
   * @returns {Array|boolean} - An array containing [column, row] or false if not found.
   */
  static findCellRowAndColumnWithText(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    textToFind: string,
  ): number[] | boolean {
    // Create a text finder instance
    let sheet: GoogleAppsScript.Spreadsheet.Sheet | null =
      spreadsheet.getSheetByName("Scoring");

    if (!sheet) {
      sheet = spreadsheet.getSheetByName("Sheet1");
    }

    if (!sheet) {
      return false;
    }

    const textFinder = sheet.createTextFinder(textToFind);

    // Account for discrepency in formatting
    const minRange = sheet
      .createTextFinder("Final Scores")
      .matchEntireCell(true)
      .findNext();
    if (minRange) {
      var minCol = minRange.getColumn();
    } else {
      const finalRankingsRange = sheet
        .createTextFinder("Final Rankings")
        .matchEntireCell(true)
        .findNext();
      if (finalRankingsRange) {
        minCol = finalRankingsRange.getColumn() - 6;
      } else {
        return false;
      }
    }

    var maxCol = minCol + 5;

    // Find all occurrences of the text
    var matchedRanges = textFinder.matchEntireCell(true).findAll();
    var maxRowRange;

    for (const i in matchedRanges) {
      var range = matchedRanges[i];
      if (
        range.getColumn() >= minCol &&
        range.getColumn() <= maxCol &&
        range.getRow() < 20
      ) {
        if (!maxRowRange) {
          maxRowRange = range;
        }
        if (range.getRow() > maxRowRange.getRow()) {
          maxRowRange = range;
        }
      }
    }

    if (maxRowRange) {
      return [
        maxRowRange.getColumn(),
        SpreadsheetUtils.findFirstNonMergedRow(
          sheet,
          maxRowRange.getColumn(),
          maxRowRange.getRow(),
        ),
      ];
    } else {
      return false;
    }
  }

  /**
   * Finds the first non-merged row for a specific cell.
   * @param {Sheet} sheet - The sheet.
   * @param {number} startColumn - The start column.
   * @param {number} startRow - The start row.
   * @returns {number} - The row number.
   */
  static findFirstNonMergedRow(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    startColumn: number,
    startRow: number,
  ): number {
    var row = startRow + 1; // Start checking from the row after the merged cell

    var cell = sheet.getRange(row, startColumn, 5, 20);
    var mergedRanges = cell.getMergedRanges();

    while (row < 1000) {
      var cell = sheet.getRange(row, startColumn, 5, 20);
      var mergedRanges = cell.getMergedRanges();
      var isMerged = false;

      for (const i in mergedRanges) {
        if (Utils.rangeIntersect(cell, mergedRanges[i])) {
          isMerged = true;
          break;
        }
      }

      if (!isMerged) {
        return row; // Found the first non-merged cell, return its row number
      }

      row++; // Move to the next row
    }
    throw new Error(
      "Exceeded maximum iterations while searching for non-merged cell.",
    );
  }

  /**
   * Adds IMPORTRANGE formulas to the scoring spreadsheets.
   * @param {Folder} targetFolder - The target folder.
   * @param {string} targetSheetName - The name of the target sheet.
   * @param {Spreadsheet} sourceSheet - The source sheet.
   */
  static pasteLookupFormulasToScoringSheets(
    targetFolder: GoogleAppsScript.Drive.Folder,
    targetSheetName: string,
    sourceSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  ) {
    /*
    Columns to pull
    Score
    Tier
    Tiebreaker
    */

    var existing_ss = targetFolder.getFilesByName(targetSheetName);
    if (existing_ss.hasNext()) {
      var targetSheet = SpreadsheetApp.openById(existing_ss.next().getId());
    } else {
      return;
    }

    var sourceSheetUrl = sourceSheet.getUrl();

    var columnsToTransfer = ["Score", "Tier", "Tiebreaker"];
    var columnsToTransferIndex = ["C", "D", "E"];

    for (const i in columnsToTransfer) {
      var columnName = columnsToTransfer[i];
      var targetColumnIndex = columnsToTransferIndex[i];
      const cell: number[] | boolean = SpreadsheetUtils.findCellRowAndColumnWithText(
        sourceSheet,
        columnName,
      );

      if (!cell || !Array.isArray(cell)) {
        continue;
      }

      const row = cell[1];
      const column = Utils.getColumnLetters(cell[0]);

      var formula =
        '=IMPORTRANGE("' +
        sourceSheetUrl +
        '", "' +
        column +
        row +
        ":" +
        column +
        (row + 102) +
        '")';
      Logger.log(columnName + " " + row + " " + column + " " + formula);

      targetSheet.getRange(targetColumnIndex + "2").setFormula(formula);
    }
  }
}