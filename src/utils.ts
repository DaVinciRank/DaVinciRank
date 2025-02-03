/**
 * Displays a prompt to the user and returns the response.
 * @param {string} prompt - The prompt message.
 * @returns {string} - The user's response.
 */
function showPrompt(prompt: string): string {
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
    return showPrompt(prompt);
  }
}

/**
 * Retrieves the parsed full tournament name based on the spreadsheet data.
 * @returns {string} - The full tournament name.
 */
function getTournamentNameParsed(): string {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!currentSheet) {
    throw new Error("No active spreadsheet found.");
  }

  var tournamentName = currentSheet
    .getRangeByName("TournamentName")
    ?.getValue();
  if (tournamentName == "" || tournamentName == "Tournament Name") {
    tournamentName = showPrompt(
      "You have not entered a tournament name. Please enter one now",
    );
    currentSheet.getRangeByName("TournamentName")?.setValue(tournamentName);
  }

  var tournamentDate = currentSheet
    .getRangeByName("TournamentDate")
    ?.getValue();
  var parsedDate = Utilities.formatDate(
    tournamentDate,
    "America/Los_Angeles",
    "d-MMMM-YYYY",
  );
  if (parsedDate == "1-January-2000") {
    tournamentDate = showPrompt(
      "You have not entered a tournament date. Please enter one now",
    );
    currentSheet.getRangeByName("TournamentDate")?.setValue(tournamentDate);
    var tournamentDate = currentSheet
      .getRangeByName("TournamentDate")
      ?.getValue();
    var parsedDate = Utilities.formatDate(
      tournamentDate,
      "America/Los_Angeles",
      "d-MMMM-YYYY",
    );
  }

  var tournamentDivision = currentSheet.getRangeByName("Division")?.getValue();
  if (tournamentDivision == "" || tournamentDivision == "__") {
    tournamentDivision = showPrompt(
      "You have not entered a tournament division. Please enter one now",
    );
    currentSheet.getRangeByName("Division")?.setValue(tournamentDivision);
  }

  var tournamentLocation = currentSheet.getRangeByName("Location")?.getValue();
  if (tournamentLocation == "" || tournamentLocation == "School_Name") {
    tournamentLocation = showPrompt(
      "You have not entered a tournament location. Please enter one now",
    );
    currentSheet.getRangeByName("Location")?.setValue(tournamentLocation);
  }

  var fullTournamentDate =
    parsedDate +
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
function getColumnLetters(columnIndexStartFromOne: number): string {
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
    return getColumnLetters(div) + ALPHABETS[res - 1];
  }
}

/**
 * Checks if two ranges intersect.
 * @param {Range} R1 - The first range.
 * @param {Range} R2 - The second range.
 * @returns {boolean} - True if they intersect, otherwise false.
 */
function rangeIntersect(
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
function moveRows(
  templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
  newSheet: GoogleAppsScript.Spreadsheet.Sheet,
  eventName: string,
): void {
  // Define the ranges we need to copy over

  const replacementRanges = ["A2:B104", "AE7:AE9", "AA8", "U3:U103", "K1:O1"];

  // Copy over the event name
  newSheet.getRange("L2:O2").setValue(eventName);

  for (let i = 0; i < replacementRanges.length; i++) {
    const range = replacementRanges[i];
    newSheet
      .getRange(range)
      .setValues(templateSheet.getRange(range).getValues());
  }
}

export {
  showPrompt,
  getTournamentNameParsed,
  moveRows,
  rangeIntersect,
  getColumnLetters,
};
