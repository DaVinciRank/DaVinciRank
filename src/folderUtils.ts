import { TournamentUtils } from "./tournamentUtils";
import { CacheLogger } from "./cacheLogger";

export class FolderUtils {
  /**
   * Retrieves the ID of the parent folder of the current spreadsheet.
   * @returns {GoogleAppsScript.Drive.Folder} - The ID of the parent folder.
   */
  static getParentFolderId(): GoogleAppsScript.Drive.Folder {
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    return spreadsheetFile.getParents().next();
  }

  /**
   * Finds a folder that matches the given substrings under the root folder.
   * @param {GoogleAppsScript.Drive.Folder} rootFolder - The root folder to search in.
   * @param {string} folderSubString - The primary substring to match.
   * @param {string} [secondarySubString] - Optional secondary substring to match.
   * @returns {GoogleAppsScript.Drive.Folder | null} - The matching folder or null if not found.
   */
  static findFolderBySubstrings(
    rootFolder: GoogleAppsScript.Drive.Folder,
    folderSubString: string,
    secondarySubString?: string,
  ): GoogleAppsScript.Drive.Folder | null {
    const subfolders = rootFolder.getFolders();
    const matchingFolders: GoogleAppsScript.Drive.Folder[] = [];

    // Collect all folders matching the first substring
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      if (folder.getName().includes(folderSubString)) {
        matchingFolders.push(folder);
      }
    }

    // If secondary substring provided, filter matching folders further
    if (matchingFolders.length > 0 && secondarySubString) {
      const secondaryMatches = matchingFolders.filter((folder) =>
        folder.getName().includes(secondarySubString),
      );

      if (secondaryMatches.length > 0) {
        return secondaryMatches[0];
      }
    }

    // Return first match if any, otherwise null
    return matchingFolders.length > 0 ? matchingFolders[0] : null;
  }

  /**
   * Creates a new folder under the root folder with the given name.
   * @param {GoogleAppsScript.Drive.Folder} rootFolder - The root folder.
   * @param {string} folderName - The name of the new folder.
   * @returns {GoogleAppsScript.Drive.Folder} - The newly created folder.
   */
  static createFolderUnderRoot(
    rootFolder: GoogleAppsScript.Drive.Folder,
    folderName: string,
  ): GoogleAppsScript.Drive.Folder {
    return rootFolder.createFolder(folderName);
  }

  /**
   * Finds or creates a folder under the root folder.
   * @param {GoogleAppsScript.Drive.Folder} rootFolder - The root folder.
   * @param {string} folderName - The name of the folder to create if substring does not exist.
   * @param {string} folderSubString - The substring of the folder name.
   * @param {string} secondarySubString - The if first substring has multiple matches, use this as secondary match
   * @returns {GoogleAppsScript.Drive.Folder} - The folder.
   */
  static findOrCreateFolderUnderRootFolder(
    rootFolder: GoogleAppsScript.Drive.Folder,
    folderName: string,
    folderSubString?: string,
    secondarySubString?: string,
  ): GoogleAppsScript.Drive.Folder {
    const existingFolder = FolderUtils.findFolderBySubstrings(
      rootFolder,
      folderSubString ? folderSubString : folderName,
      secondarySubString,
    );

    return (
      existingFolder ||
      FolderUtils.createFolderUnderRoot(rootFolder, folderName)
    );
  }

  /**
   * Removes a file if it already exists in a folder.
   * @param {Folder} folder - The folder.
   * @param {string} fileName - The name of the file.
   */
  static removeFileIfExists(
    folder: GoogleAppsScript.Drive.Folder,
    fileName: string,
  ) {
    // Try to find a file with the same name in the destination folder
    var existingFiles = folder.getFilesByName(fileName);
    if (existingFiles.hasNext()) {
      // There is an existing file with the same name, so delete it
      var existingFile = existingFiles.next();
      existingFile.setTrashed(true); // Move to trash
    }
  }

  /**
   * Gets template files containing a specific substring.
   * @param {string} substring - The substring.
   * @param {File[]} allTemplateFiles - All template files.
   * @returns {File[]} - Template files containing the substring.
   */
  static getTemplateFilesWithSubstring(
    substring: string,
    allTemplateFiles: GoogleAppsScript.Drive.File[],
  ): GoogleAppsScript.Drive.File[] {
    var templateFiles = [];
    for (const i in allTemplateFiles) {
      var fileName = allTemplateFiles[i].getName();
      if (fileName.includes(substring)) {
        templateFiles.push(allTemplateFiles[i]);
      }
    }
    return templateFiles;
  }

  /**
   * Gets all files under a root folder.
   * @param {string} rootFolder - The root folder
   * @returns {File[]} - All files under the root folder.
   */
  static getFilesUnderRootFolder(rootFolder: GoogleAppsScript.Drive.Folder) {
    var files = [];
    var filesIterator = rootFolder.getFiles();
    while (filesIterator.hasNext()) {
      // Iterate through the files
      var file = filesIterator.next();
      files.push(file);
    }
    return files;
  }

  /**
   * Adds editors to a folder if not already added.
   * @param {GoogleAppsScript.Drive.Folder} folder - The folder.
   * @param {string[]} emails - The email addresses of the editors.
   */
  static addEditorToFolder(
    folder: GoogleAppsScript.Drive.Folder,
    emails: string[],
  ) {
    var existingEditors = folder.getEditors().map(function (editor) {
      return editor.getEmail();
    });

    for (var i in emails) {
      var email = emails[i];
      if (email !== "" && existingEditors.indexOf(email) === -1) {
        try {
          folder.addEditor(email); // Attempt to add the editor
        } catch (e: unknown) {
          if (e instanceof Error) {
            Logger.log(
              "Error adding editor: " + email + ". Error: " + e.message,
            );
          } else {
            Logger.log(
              "Error adding editor: " + email + ". Unknown error occurred.",
            );
          }
        }
      }
    }
  }

  /**
   * Shares all scoring folders with specified emails.
   */
  static shareScoringFoldersWithEmails() {
    const tournamentName = TournamentUtils.getTournamentNameParsed();
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = currentSheet.getRangeByName("EventsAndEmailSharing");
    if (!range) {
      Logger.log("Range 'EventsAndEmailSharing' not found.");
      return;
    }
    var values = range.getValues();
    if (!values) {
      Logger.log("No values found in the range.");
      return;
    }
    var rangeValues = values.filter(function (subList) {
      return subList[0] !== "";
    });

    var parentFolder = FolderUtils.getParentFolderId();
    var scoreSheetFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
      parentFolder,
      "Event Specific Score Sheets - " + tournamentName,
    );

    for (const j in rangeValues) {
      var eventName = rangeValues[j][0];
      CacheLogger.appendLog("Adding ES emails for " + eventName);
      var spreadSheetName = eventName + " Event Scoring - " + tournamentName;
      var spreadSheetFolder = FolderUtils.findOrCreateFolderUnderRootFolder(
        scoreSheetFolder,
        spreadSheetName,
      );
      FolderUtils.addEditorToFolder(
        spreadSheetFolder,
        rangeValues[j].slice(1, 4),
      );
    }
  }

  /**
   * Gets the scoresheet folder from the spreadsheet.
   * @returns {Folder} - The scoresheet folder.
   */
  static getScoreSheetFolder(
    prompt_user: boolean = true,
  ): GoogleAppsScript.Drive.Folder | null {
    const rootFolder = FolderUtils.getParentFolderId();
    const tournamentName = TournamentUtils.getTournamentNameParsed(prompt_user);
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

  /**
   * Gets the scoresheet folder from the spreadsheet.
   * @returns {Folder} - The scoresheet folder.
   */
  static getTemplateFolder(
    prompt_user: boolean = true,
  ): GoogleAppsScript.Drive.Folder | null {
    const rootFolder = FolderUtils.getParentFolderId();
    const tournamentName = TournamentUtils.getTournamentNameParsed(prompt_user);
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
}
