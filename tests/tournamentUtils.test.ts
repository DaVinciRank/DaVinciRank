import { TournamentUtils } from "../src/tournamentUtils";

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

  const tournamentName = TournamentUtils.getTournamentNameParsed();
  expect(tournamentName).toBe(
    "1-January-2023 Test Tournament Division-Test Tournament @ Test Tournament",
  );
});
