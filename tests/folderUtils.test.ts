import { FolderUtils } from "../src/folderUtils";
import { TournamentUtils } from "../src/tournamentUtils";
import { CacheLogger } from "../src/cacheLogger";

const mockFolder = {
  getId: jest.fn(),
  getName: jest.fn().mockReturnValue("Event Specific Score Sheets"),
  getFilesByName: jest.fn(),
  getFiles: jest.fn(),
  getFolders: jest.fn(),
  getParents: jest.fn(),
  createFolder: jest.fn(),
  getEditors: jest.fn().mockReturnValue([
    {
      getEmail: jest.fn().mockReturnValue("test@example.com"),
      getDomain: jest.fn().mockReturnValue("example.com"),
      getName: jest.fn().mockReturnValue("Test User"),
      getPhotoUrl: jest.fn().mockReturnValue("http://example.com/photo.jpg"),
      getUserLoginId: jest.fn().mockReturnValue("user123"),
    },
  ]),
  addEditor: jest.fn(),
  addEditors: jest.fn(),
  addFile: jest.fn(),
  addFolder: jest.fn(),
  addViewer: jest.fn(),
  addViewers: jest.fn(),
  removeEditor: jest.fn(),
  removeViewer: jest.fn(),
} as unknown as jest.Mocked<GoogleAppsScript.Drive.Folder>;

// Ensure folderIterator is a proper mock
const folderIterator = {
  hasNext: jest.fn().mockReturnValue(true),
  next: jest.fn().mockReturnValue(mockFolder),
} as unknown as jest.Mocked<GoogleAppsScript.Drive.FolderIterator>;

// Ensure mockFile is a proper mock
const mockFile = {
  getName: jest.fn(),
  setTrashed: jest.fn(),
} as unknown as jest.Mocked<GoogleAppsScript.Drive.File>;

// Ensure fileIterator is a proper mock
const fileIterator = {
  hasNext: jest.fn().mockReturnValue(true),
  next: jest.fn().mockReturnValue(mockFile),
} as unknown as jest.Mocked<GoogleAppsScript.Drive.FileIterator>;

// Spreadsheet mocks
const mockSpreadsheet = {
  getId: jest.fn(),
  getRangeByName: jest.fn(),
  getActiveSpreadsheet: jest.fn(),
};

const mockRange = {
  getValues: jest.fn(),
};

const mockCache = {
  get: jest.fn(),
  put: jest.fn(),
  remove: jest.fn(),
};

// Global mocks for Google Apps Script services
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => mockSpreadsheet),
} as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

global.DriveApp = {
  getFileById: jest.fn(() => ({
    getParents: jest.fn(() => ({
      next: jest.fn(() => mockFolder),
    })),
  })),
} as unknown as GoogleAppsScript.Drive.DriveApp;

global.CacheService = {
  getUserCache: jest.fn(() => mockCache),
} as unknown as GoogleAppsScript.Cache.CacheService;

describe("FolderUtils", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  test("getParentFolderId should return the parent folder", () => {
    const result = FolderUtils.getParentFolderId();

    expect(DriveApp.getFileById).toHaveBeenCalled();
    expect(result).toBe(mockFolder);
  });

  test("findFolderBySubstrings should return matching folder", () => {
    folderIterator.hasNext.mockReturnValueOnce(true).mockReturnValueOnce(false);
    folderIterator.next.mockReturnValue(mockFolder);
    mockFolder.getFolders.mockReturnValue(folderIterator);
    mockFolder.getName.mockReturnValue("Event Specific Score Sheets");

    const result = FolderUtils.findFolderBySubstrings(
      mockFolder,
      "Score Sheets",
    );

    expect(result).toBe(mockFolder);
  });

  test("createFolderUnderRoot should create and return a new folder", () => {
    mockFolder.createFolder.mockReturnValue(mockFolder);

    const result = FolderUtils.createFolderUnderRoot(mockFolder, "New Folder");

    expect(mockFolder.createFolder).toHaveBeenCalledWith("New Folder");
    expect(result).toBe(mockFolder);
  });

  test("findOrCreateFolderUnderRootFolder should find existing folder", () => {
    jest
      .spyOn(FolderUtils, "findFolderBySubstrings")
      .mockReturnValue(mockFolder);

    const result = FolderUtils.findOrCreateFolderUnderRootFolder(
      mockFolder,
      "Existing Folder",
    );

    expect(result).toBe(mockFolder);
  });

  test("findOrCreateFolderUnderRootFolder should create folder if not found", () => {
    jest.spyOn(FolderUtils, "findFolderBySubstrings").mockReturnValue(null);
    jest
      .spyOn(FolderUtils, "createFolderUnderRoot")
      .mockReturnValue(mockFolder);

    const result = FolderUtils.findOrCreateFolderUnderRootFolder(
      mockFolder,
      "New Folder",
    );

    expect(FolderUtils.createFolderUnderRoot).toHaveBeenCalledWith(
      mockFolder,
      "New Folder",
    );
    expect(result).toBe(mockFolder);
  });

  // test("removeFileIfExists should trash the file if it exists", () => {
  //   mockFolder.getFilesByName.mockReturnValue(fileIterator);

  //   FolderUtils.removeFileIfExists(mockFolder, "Old File");

  //   expect(mockFile.setTrashed).toHaveBeenCalledWith(true);
  // });

  test("getTemplateFilesWithSubstring should return matching files", () => {
    const files = [
      { getName: () => "Template 1" },
      { getName: () => "Other File" },
    ];
    const result = FolderUtils.getTemplateFilesWithSubstring(
      "Template",
      files as GoogleAppsScript.Drive.File[],
    );

    expect(result.length).toBe(1);
    expect(result[0].getName()).toBe("Template 1");
  });

  // test("getFilesUnderRootFolder should return all files", () => {
  //   mockFolder.getFiles.mockReturnValue(fileIterator);

  //   const result = FolderUtils.getFilesUnderRootFolder(mockFolder);

  //   expect(result.length).toBe(1);
  //   expect(result[0]).toBe(mockFile);
  // });

  test("addEditorToFolder should add new editor", () => {
    mockFolder.getEditors.mockReturnValue([
      {
        getEmail: jest.fn().mockReturnValue("existing@example.com"),
        getDomain: jest.fn().mockReturnValue("example.com"),
        getName: jest.fn().mockReturnValue("Existing User"),
        getPhotoUrl: jest.fn().mockReturnValue("http://example.com/photo.jpg"),
        getUserLoginId: jest.fn().mockReturnValue("user123"),
      },
    ]);

    FolderUtils.addEditorToFolder(mockFolder, ["new@example.com"]);

    expect(mockFolder.addEditor).toHaveBeenCalledWith("new@example.com");
  });

  test("shareScoringFoldersWithEmails should share folders", () => {
    jest
      .spyOn(TournamentUtils, "getTournamentNameParsed")
      .mockReturnValue("Test Tournament");
    mockSpreadsheet.getRangeByName.mockReturnValue(mockRange);
    mockRange.getValues.mockReturnValue([["Event 1", "editor1@example.com"]]);

    jest.spyOn(FolderUtils, "getParentFolderId").mockReturnValue(mockFolder);
    jest
      .spyOn(FolderUtils, "findOrCreateFolderUnderRootFolder")
      .mockReturnValue(mockFolder);
    jest.spyOn(FolderUtils, "addEditorToFolder").mockImplementation();

    FolderUtils.shareScoringFoldersWithEmails();

    expect(FolderUtils.addEditorToFolder).toHaveBeenCalledWith(mockFolder, [
      "editor1@example.com",
    ]);
  });

  test("getScoreSheetFolder should return scoresheet folder", () => {
    jest.spyOn(FolderUtils, "getParentFolderId").mockReturnValue(mockFolder);
    jest
      .spyOn(TournamentUtils, "getTournamentNameParsed")
      .mockReturnValue("Test Tournament");
    jest
      .spyOn(FolderUtils, "findOrCreateFolderUnderRootFolder")
      .mockReturnValue(mockFolder);

    const result = FolderUtils.getScoreSheetFolder();

    expect(result).toBe(mockFolder);
  });

  test("getTemplateFolder should return template folder", () => {
    jest.spyOn(FolderUtils, "getParentFolderId").mockReturnValue(mockFolder);
    jest
      .spyOn(TournamentUtils, "getTournamentNameParsed")
      .mockReturnValue("Test Tournament");
    jest
      .spyOn(FolderUtils, "findOrCreateFolderUnderRootFolder")
      .mockReturnValue(mockFolder);

    const result = FolderUtils.getTemplateFolder();

    expect(result).toBe(mockFolder);
  });
});
