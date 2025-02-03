import {
  showPrompt,
  getTournamentNameParsed,
  rangeIntersect,
  getColumnLetters,
} from "../src/utils";

test("showPrompt should return user response", () => {
  // Mock the SpreadsheetApp and Logger
  global.SpreadsheetApp = {
    getUi: jest.fn().mockReturnValue({
      prompt: jest.fn().mockReturnValue({
        getSelectedButton: jest.fn().mockReturnValue("OK"),
        getResponseText: jest.fn().mockReturnValue("Test Response"),
      }),
      ButtonSet: {
        OK_CANCEL: "OK_CANCEL",
      },
      Button: {
        OK: "OK",
      },
    }),
  } as unknown as typeof SpreadsheetApp;
  global.Logger = {
    log: jest.fn(),
  } as unknown as typeof Logger;

  const response = showPrompt("Test Prompt");
  expect(response).toBe("Test Response");
});

test("getTournamentNameParsed should return tournament name", () => {
  // Mock the SpreadsheetApp and Utilities
  global.SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn().mockReturnValue({
      getRangeByName: jest.fn().mockReturnValue({
        getValue: jest.fn().mockReturnValue("Test Tournament"),
      }),
    }),
  } as unknown as typeof SpreadsheetApp;
  global.Utilities = {
    formatDate: jest.fn().mockReturnValue("1-January-2023"),
  } as unknown as typeof Utilities;

  const tournamentName = getTournamentNameParsed();
  expect(tournamentName).toBe(
    "1-January-2023 Test Tournament Division-Test Tournament @ Test Tournament",
  );
});

test("getColumnLetters should return correct column letter", () => {
  expect(getColumnLetters(1)).toBe("A");
  expect(getColumnLetters(26)).toBe("Z");
  expect(getColumnLetters(27)).toBe("AA");
  expect(getColumnLetters(52)).toBe("AZ");
  expect(getColumnLetters(53)).toBe("BA");
});

test("rangeIntersect should return true if ranges intersect", () => {
  const mockRange1 = {
    getLastRow: jest.fn().mockReturnValue(10),
    getRow: jest.fn().mockReturnValue(1),
    getLastColumn: jest.fn().mockReturnValue(5),
    getColumn: jest.fn().mockReturnValue(1),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  const mockRange2 = {
    getLastRow: jest.fn().mockReturnValue(15),
    getRow: jest.fn().mockReturnValue(5),
    getLastColumn: jest.fn().mockReturnValue(10),
    getColumn: jest.fn().mockReturnValue(3),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  expect(rangeIntersect(mockRange1, mockRange2)).toBe(true);
});

test("rangeIntersect should return true if ranges intersect", () => {
  const mockRange1 = {
    getLastRow: jest.fn().mockReturnValue(10),
    getRow: jest.fn().mockReturnValue(1),
    getLastColumn: jest.fn().mockReturnValue(5),
    getColumn: jest.fn().mockReturnValue(1),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  const mockRange2 = {
    getLastRow: jest.fn().mockReturnValue(15),
    getRow: jest.fn().mockReturnValue(5),
    getLastColumn: jest.fn().mockReturnValue(10),
    getColumn: jest.fn().mockReturnValue(3),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  expect(rangeIntersect(mockRange1, mockRange2)).toBe(true);
});

test("rangeIntersect should return false if ranges do not intersect", () => {
  const mockRange1 = {
    getLastRow: jest.fn().mockReturnValue(10),
    getRow: jest.fn().mockReturnValue(1),
    getLastColumn: jest.fn().mockReturnValue(5),
    getColumn: jest.fn().mockReturnValue(1),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  const mockRange2 = {
    getLastRow: jest.fn().mockReturnValue(15),
    getRow: jest.fn().mockReturnValue(11),
    getLastColumn: jest.fn().mockReturnValue(10),
    getColumn: jest.fn().mockReturnValue(6),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
  expect(rangeIntersect(mockRange1, mockRange2)).toBe(false);
});
