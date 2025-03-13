import { Utils } from "../src/utils";

test("getColumnLetters should return correct column letter", () => {
  expect(Utils.getColumnLetters(1)).toBe("A");
  expect(Utils.getColumnLetters(26)).toBe("Z");
  expect(Utils.getColumnLetters(27)).toBe("AA");
  expect(Utils.getColumnLetters(52)).toBe("AZ");
  expect(Utils.getColumnLetters(53)).toBe("BA");
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
  expect(Utils.rangeIntersect(mockRange1, mockRange2)).toBe(true);
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
  expect(Utils.rangeIntersect(mockRange1, mockRange2)).toBe(true);
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
  expect(Utils.rangeIntersect(mockRange1, mockRange2)).toBe(false);
});
