import { Constants } from "./constants";
import { UiUtils } from "./uiUtils";

export class TournamentUtils {
  /**
   * Retrieves the parsed full tournament name based on the spreadsheet data.
   * @returns {string} - The full tournament name.
   */
  static getTournamentNameParsed(prompt_user: boolean = true): string {
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!currentSheet) {
      throw new Error("No active spreadsheet found.");
    }

    var tournamentName = TournamentUtils.getTournamentName();
    if (!tournamentName && prompt_user) {
      tournamentName = UiUtils.showPrompt(
        "You have not entered a tournament name. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.TOURNAMENT_NAME_RANGE_NAME)
        ?.setValue(tournamentName);
    }

    var tournamentDate = TournamentUtils.getTournamentDate();
    while (!tournamentDate && prompt_user) {
      tournamentDate = UiUtils.showPrompt(
        "You have not entered a tournament date. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.TOURNAMENT_DATE_RANGE_NAME)
        ?.setValue(tournamentDate);
      tournamentDate = TournamentUtils.getTournamentDate();
    }

    var tournamentDivision = TournamentUtils.getTournamentDivision();
    if (!tournamentDivision && prompt_user) {
      tournamentDivision = UiUtils.showPrompt(
        "You have not entered a tournament division. Please enter one now",
      );
      currentSheet
        .getRangeByName(Constants.DIVISION_RANGE_NAME)
        ?.setValue(tournamentDivision);
    }

    var tournamentLocation = TournamentUtils.getTournamentLocation();
    if (!tournamentLocation && prompt_user) {
      tournamentLocation = UiUtils.showPrompt(
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
   * Gets the number of event medals from the spreadsheet.
   * @returns {number} - The number of medals to present.
   */
  static getNumberOfEventMedals(): number {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const numberOfEventMedals = currentSheet
      .getRangeByName(Constants.NUMBER_OF_EVENT_MEDALS)
      ?.getValue();
    if (numberOfEventMedals == "")
      return Constants.DEFAULT_NUMBER_OF_EVENT_MEDALS;
    return Number(numberOfEventMedals);
  }

  /**
   * Gets the number of team trophies from the spreadsheet.
   * @returns {int} - The number of team trophies to present.
   */
  static getNumberOfTeamTrophies(): number {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const numberOfTeamTrophies = currentSheet
      .getRangeByName(Constants.NUMBER_OF_TEAM_TROPHIES)
      ?.getValue();
    if (numberOfTeamTrophies == "")
      return Constants.DEFAULT_NUMBER_OF_TEAM_TROPIES;
    return Number(numberOfTeamTrophies);
  }
}
