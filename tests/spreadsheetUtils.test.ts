import { SpreadsheetUtils } from "../src/spreadsheetUtils";

const mockCache = {
  get: jest.fn(),
  put: jest.fn(),
  remove: jest.fn(),
};

global.CacheService = {
  getUserCache: jest.fn(() => mockCache),
} as unknown as GoogleAppsScript.Cache.CacheService;

const mockSheet = {
  getSheetName: jest.fn(),
} as unknown as jest.Mocked<GoogleAppsScript.Spreadsheet.Sheet>;

const mockSpreadsheet = {
  getSheetByName: jest.fn(),
} as unknown as jest.Mocked<GoogleAppsScript.Spreadsheet.Spreadsheet>;

describe("SpreadsheetUtils", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe("findScoringSheetWithinSpreadsheet", () => {
    test("should return Export sheet when available", () => {
      mockSheet.getSheetName.mockReturnValue("Export");
      mockSpreadsheet.getSheetByName
        .mockReturnValueOnce(mockSheet)
        .mockReturnValue(null);

      const result =
        SpreadsheetUtils.findScoringSheetWithinSpreadsheet(mockSpreadsheet);

      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Export");
      expect(result).toBe(mockSheet);
    });

    test("should return Scoring sheet when Export not available", () => {
      mockSheet.getSheetName.mockReturnValue("Scoring");
      mockSpreadsheet.getSheetByName
        .mockReturnValueOnce(null)
        .mockReturnValueOnce(mockSheet);

      const result =
        SpreadsheetUtils.findScoringSheetWithinSpreadsheet(mockSpreadsheet);

      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Export");
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Scoring");
      expect(result).toBe(mockSheet);
    });

    test("should return Sheet1 when Export and Scoring not available", () => {
      mockSheet.getSheetName.mockReturnValue("Sheet1");
      mockSpreadsheet.getSheetByName
        .mockReturnValueOnce(null)
        .mockReturnValueOnce(null)
        .mockReturnValueOnce(mockSheet);

      const result =
        SpreadsheetUtils.findScoringSheetWithinSpreadsheet(mockSpreadsheet);

      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Export");
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Scoring");
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("Sheet1");
      expect(result).toBe(mockSheet);
    });

    test("should return null when no sheets found", () => {
      mockSpreadsheet.getSheetByName.mockReturnValue(null);

      const result =
        SpreadsheetUtils.findScoringSheetWithinSpreadsheet(mockSpreadsheet);

      expect(result).toBeNull();
    });
  });
});
