import { Constants } from "./constants";
import { FolderUtils } from "./folderUtils";

export class Utils {
  /**
   * Displays a prompt to the user and returns the response.
   * @param {string} prompt - The prompt message.
   * @returns {string} - The user's response.
   */
  static showPrompt(prompt: string): string | null {
    var ui = SpreadsheetApp.getUi();

    var result = ui.prompt(
      prompt ? prompt : "Give an input",
      "Input:",
      ui.ButtonSet.OK_CANCEL,
    );

    var button = result.getSelectedButton();
    var response = result.getResponseText();
    if (button == ui.Button.OK) {
      // call function and pass the value
      Logger.log(response);
      return response;
    } else {
      return null;
    }
  }

  /**
   * Retrieves the parsed full tournament name based on the spreadsheet data.
   * @returns {string} - The full tournament name.
   */
  static getTournamentNameParsed(prompt_user: boolean = true): string {
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!currentSheet) {
      throw new Error("No active spreadsheet found.");
    }

    var tournamentName = Utils.getTournamentName();
    if (!tournamentName && prompt_user) {
      tournamentName = Utils.showPrompt(
        "You have not entered a tournament name. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.TOURNAMENT_NAME_RANGE_NAME)
        ?.setValue(tournamentName);
    }

    var tournamentDate = Utils.getTournamentDate();
    while (!tournamentDate && prompt_user) {
      tournamentDate = Utils.showPrompt(
        "You have not entered a tournament date. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.TOURNAMENT_DATE_RANGE_NAME)
        ?.setValue(tournamentDate);
      tournamentDate = Utils.getTournamentDate();
    }

    var tournamentDivision = Utils.getTournamentDivision();
    if (!tournamentDivision && prompt_user) {
      tournamentDivision = Utils.showPrompt(
        "You have not entered a tournament division. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.DIVISION_RANGE_NAME)
        ?.setValue(tournamentDivision);
    }

    var tournamentLocation = Utils.getTournamentLocation();
    if (!tournamentLocation && prompt_user) {
      tournamentLocation = Utils.showPrompt(
        "You have not entered a tournament location. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.LOCATION_RANGE_NAME)
        ?.setValue(tournamentLocation);
    }

    if (
      !tournamentName ||
      !tournamentDate ||
      !tournamentDivision ||
      !tournamentLocation
    ) {
      return "";
    }

    var fullTournamentDate =
      tournamentDate +
      " " +
      tournamentName +
      " Division-" +
      tournamentDivision +
      " @ " +
      tournamentLocation;

    return fullTournamentDate;
  }

  /**
   * Converts column index to letter format.
   * @param {number} columnIndexStartFromOne - The column index starting from one.
   * @returns {string} - The column letter.
   */
  static getColumnLetters(columnIndexStartFromOne: number): string {
    // https://www.allstacksdeveloper.com/2021/08/how-to-convert-column-index-into-letters-with-google-apps-script.html
    const ALPHABETS = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W",
      "X",
      "Y",
      "Z",
    ];

    if (columnIndexStartFromOne < 27) {
      return ALPHABETS[columnIndexStartFromOne - 1];
    } else {
      var res = columnIndexStartFromOne % 26;
      var div = Math.floor(columnIndexStartFromOne / 26);
      if (res === 0) {
        div = div - 1;
        res = 26;
      }
      return Utils.getColumnLetters(div) + ALPHABETS[res - 1];
    }
  }

  /**
   * Checks if two ranges intersect.
   * @param {Range} R1 - The first range.
   * @param {Range} R2 - The second range.
   * @returns {boolean} - True if they intersect, otherwise false.
   */
  static rangeIntersect(
    R1: GoogleAppsScript.Spreadsheet.Range,
    R2: GoogleAppsScript.Spreadsheet.Range,
  ): boolean {
    var LR1 = R1.getLastRow();
    var Ro2 = R2.getRow();
    if (LR1 < Ro2) return false;

    var LR2 = R2.getLastRow();
    var Ro1 = R1.getRow();
    if (LR2 < Ro1) return false;

    var LC1 = R1.getLastColumn();
    var C2 = R2.getColumn();
    if (LC1 < C2) return false;

    var LC2 = R2.getLastColumn();
    var C1 = R1.getColumn();
    if (LC2 < C1) return false;

    return true;
  }

  /**
   * Moves rows from the template to the new spreadsheet.
   * @param {Sheet} templateSheet - The template sheet.
   * @param {Sheet} newSheet - The new sheet.
   * @param {string} eventName - The name of the event.
   */
  static moveRows(
    templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    newSheet: GoogleAppsScript.Spreadsheet.Sheet,
    eventName: string,
    highLowScoreValues: string,
  ): void {
    // Define the ranges we need to copy over
    const replacementRanges = ["A2:B104", "AE7:AE9", "AA8", "U3:U103", "K1:O1"];

    // Copy over the event name
    newSheet.getRange("L2:O2").setValue(eventName);
    // Copy over high/low score wins
    newSheet.getRange("L4").setValue(highLowScoreValues);
    newSheet.getRange("L5").setValue(highLowScoreValues);

    for (let i = 0; i < replacementRanges.length; i++) {
      const range = replacementRanges[i];
      newSheet
        .getRange(range)
        .setValues(templateSheet.getRange(range).getValues());
    }
  }

  /**
   * Gets the tournament name from the spreadsheet.
   * @returns {string} - The tournament name.
   */
  static getTournamentName(): string | null {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const tournamentName = currentSheet
      .getRangeByName(Constants.TOURNAMENT_NAME_RANGE_NAME)
      ?.getValue();
    if (
      tournamentName == "" ||
      tournamentName == Constants.DEFAULT_TOURNAMENT_NAME
    ) {
      return null;
    }
    return tournamentName;
  }

  /**
   * Gets the tournament division from the spreadsheet.
   * @returns {string} - The tournament division.
   */
  static getTournamentDivision(): string | null {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const tournamentDivision = currentSheet
      .getRangeByName(Constants.DIVISION_RANGE_NAME)
      ?.getValue();
    if (
      tournamentDivision == "" ||
      tournamentDivision == Constants.DEFAULT_TOURNAMENT_DIVISION
    ) {
      return null;
    }
    return tournamentDivision;
  }

  /**
   * Gets the tournament location from the spreadsheet.
   * @returns {string} - The tournament location.
   */
  static getTournamentLocation(): string | null {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const tournamentLocation = currentSheet
      .getRangeByName(Constants.LOCATION_RANGE_NAME)
      ?.getValue();
    if (
      tournamentLocation == "" ||
      tournamentLocation == Constants.DEFAULT_TOURNAMENT_LOCATION
    ) {
      return null;
    }
    return tournamentLocation;
  }

  /**
   * Gets the tournament date from the spreadsheet.
   * @returns {string} - The tournament date.
   */
  static getTournamentDate(): string | null {
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    var tournamentDate = currentSheet
      .getRangeByName(Constants.TOURNAMENT_DATE_RANGE_NAME)
      ?.getValue();
    var parsedDate = Utilities.formatDate(
      tournamentDate,
      "America/Los_Angeles",
      "d-MMMM-YYYY",
    );
    if (parsedDate == Constants.DEFAULT_TOURNAMENT_DATE) {
      return null;
    }
    return parsedDate;
  }

  /**
   * Gets the number of teams from the spreadsheet.
   * @returns {string} - The number of teams.
   */
  static getNumberOfTeams(): string | null {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const tournamentLocation = currentSheet
      .getRangeByName(Constants.NUMBER_OF_TEAMS_RANGE_NAME)
      ?.getValue();
    if (
      tournamentLocation == "" ||
      tournamentLocation == Constants.DEFAULT_NUMBER_OF_TEAMS
    ) {
      return null;
    }
    return tournamentLocation;
  }

  /**
   * Gets the scoresheet folder from the spreadsheet.
   * @returns {Folder} - The scoresheet folder.
   */
  static getTemplateFolder(
    prompt_user: boolean = true,
  ): GoogleAppsScript.Drive.Folder | null {
    const rootFolder = FolderUtils.getParentFolderId();
    const tournamentName = Utils.getTournamentNameParsed(prompt_user);
    if (
      tournamentName == "" &&
      !FolderUtils.findFolderBySubstrings(rootFolder, "Template File")
    ) {
      return null;
    }
    const templateFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
      rootFolder,
      "Template Files - " + tournamentName,
      "Template File",
    );
    return templateFolder;
  }

  /**
   * Gets the scoresheet folder from the spreadsheet.
   * @returns {Folder} - The scoresheet folder.
   */
  static getScoreSheetFolder(
    prompt_user: boolean = true,
  ): GoogleAppsScript.Drive.Folder | null {
    const rootFolder = FolderUtils.getParentFolderId();
    const tournamentName = Utils.getTournamentNameParsed(prompt_user);
    if (
      tournamentName == "" &&
      !FolderUtils.findFolderBySubstrings(
        rootFolder,
        "Event Specific Score Sheets - " + tournamentName,
      )
    ) {
      return null;
    }
    const scoreSheetFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
      rootFolder,
      "Event Specific Score Sheets - " + tournamentName,
    );
    return scoreSheetFolder;
  }
}
