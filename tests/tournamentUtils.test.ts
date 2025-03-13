import { Constants } from "../src/constants";
import { TournamentUtils } from "../src/tournamentUtils";
import { UiUtils } from "../src/uiUtils";

describe("TournamentUtils", () => {
  beforeEach(() => {
    // Reset mocks before each test
    jest.clearAllMocks();
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

    const tournamentName = TournamentUtils.getTournamentNameParsed();
    expect(tournamentName).toBe(
      "1-January-2023 Test Tournament Division-Test Tournament @ Test Tournament",
    );
  });

  describe("getTournamentName", () => {
    test("should return tournament name when valid", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue("Test Tournament"),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentName();
      expect(result).toBe("Test Tournament");
    });

    test("should return null for empty tournament name", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentName();
      expect(result).toBeNull();
    });
  });

  describe("getTournamentDate", () => {
    test("should return formatted date when valid", () => {
      const testDate = new Date("2024-01-15");
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(testDate),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      global.Utilities = {
        formatDate: jest.fn().mockReturnValue("15-January-2024"),
      } as unknown as typeof Utilities;

      const result = TournamentUtils.getTournamentDate();
      expect(result).toBe("15-January-2024");
      expect(Utilities.formatDate).toHaveBeenCalledWith(
        testDate,
        "America/Los_Angeles",
        "d-MMMM-YYYY",
      );
    });

    test("should return null for default date", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(new Date()),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      global.Utilities = {
        formatDate: jest
          .fn()
          .mockReturnValue(Constants.DEFAULT_TOURNAMENT_DATE),
      } as unknown as typeof Utilities;

      const result = TournamentUtils.getTournamentDate();
      expect(result).toBeNull();
    });
  });

  describe("getTournamentDivision", () => {
    test("should return division when valid", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue("Division A"),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentDivision();
      expect(result).toBe("Division A");
    });

    test("should return null for empty division", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentDivision();
      expect(result).toBeNull();
    });
  });

  describe("getTournamentLocation", () => {
    test("should return location when valid", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue("Test Location"),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentLocation();
      expect(result).toBe("Test Location");
    });

    test("should return null for empty location", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getTournamentLocation();
      expect(result).toBeNull();
    });
  });

  describe("getNumberOfTeams", () => {
    test("should return number of teams when valid", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue("8"),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getNumberOfTeams();
      expect(result).toBe("8");
    });

    test("should return null for empty teams value", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      const result = TournamentUtils.getNumberOfTeams();
      expect(result).toBeNull();
    });
  });

  describe("getTournamentNameParsed additional cases", () => {
    test("should return empty string when missing required fields", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
          }),
        }),
      } as unknown as typeof SpreadsheetApp;

      global.Utilities = {
        formatDate: jest.fn().mockReturnValue(""),
      } as unknown as typeof Utilities;

      const result = TournamentUtils.getTournamentNameParsed(false);
      expect(result).toBe("");
    });

    test("should handle user prompts when fields are missing", () => {
      // Mock the user prompts in the correct order based on the received output
      const mockUiUtils = {
        showPrompt: jest
          .fn()
          .mockReturnValueOnce("Test Tournament") // Tournament name
          .mockReturnValueOnce("Division A") // Division (changed order)
          .mockReturnValueOnce("Test Location"), // Location (changed order)
      };

      const mockPrompt = jest.fn().mockReturnValue({
        getResponseText: jest.fn().mockReturnValue("Test Tournament"),
        getSelectedButton: jest.fn().mockReturnValue("OK"),
      });

      const mockButtonSet = {
        OK_CANCEL: "OK_CANCEL",
      };

      const mockButton = {
        OK: "OK",
        CANCEL: "CANCEL",
      };

      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getRangeByName: jest.fn().mockReturnValue({
            getValue: jest.fn().mockReturnValue(""),
            setValue: jest.fn(),
          }),
        }),
        getUi: jest.fn().mockReturnValue({
          prompt: mockPrompt,
          ButtonSet: mockButtonSet,
          Button: mockButton,
        }),
      } as unknown as typeof SpreadsheetApp;

      global.Utilities = {
        formatDate: jest.fn().mockReturnValue("15-January-2024"),
      } as unknown as typeof Utilities;

      // Inject the mock UiUtils
      jest
        .spyOn(UiUtils, "showPrompt")
        .mockImplementation(mockUiUtils.showPrompt);

      const result = TournamentUtils.getTournamentNameParsed(true);
      expect(result).toBe(
        "15-January-2024 Test Tournament Division-Division A @ Test Location",
      );
      expect(mockUiUtils.showPrompt).toHaveBeenCalledTimes(3);
    });

    test("should throw error when no active spreadsheet", () => {
      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnValue(null),
      } as unknown as typeof SpreadsheetApp;

      expect(() => {
        TournamentUtils.getTournamentNameParsed(false);
      }).toThrow("No active spreadsheet found.");
    });
  });
});
