import { Utils } from "./utils";

export class FolderUtils {
  /**
   * Retrieves the ID of the parent folder of the current spreadsheet.
   * @returns {string} - The ID of the parent folder.
   */
  static getParentFolderId(): string {
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    var folderId = spreadsheetFile.getParents().next().getId();
    return folderId;
  }

  /**
   * Creates a folder under the specified root folder.
   * @param {string} rootFolderId - The ID of the root folder.
   * @param {string} folderName - The name of the folder to create.
   * @returns {string} - The ID of the created folder.
   */
  static createFolderUnderRootFolder(
    rootFolderId: string,
    folderName: string,
  ): string {
    var rootFolder = DriveApp.getFolderById(rootFolderId);

    const folderIterator = rootFolder.getFoldersByName(folderName);
    if (folderIterator.hasNext()) {
      // When the folder exists
      return folderIterator.next().getId();
    } else {
      // When the folder doesn't exist
      return rootFolder.createFolder(folderName).getId();
    }
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
   * @param {string} rootFolderId - The root folder ID.
   * @returns {File[]} - All files under the root folder.
   */
  static getFilesUnderRootRolder(rootFolderId: string) {
    var rootFolder = DriveApp.getFolderById(rootFolderId);
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
   * @param {string} folderId - The folder ID.
   * @param {string[]} emails - The email addresses of the editors.
   */
  static addEditorToFolder(folderId: string, emails: string[]) {
    var folder = DriveApp.getFolderById(folderId);
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
            Logger.log("Error adding editor: " + email + ". Error: " + e.message);
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

    var parentFolderId = FolderUtils.getParentFolderId();
    var scoreSheetFolderId = FolderUtils.createFolderUnderRootFolder(
      parentFolderId,
      Utils.getTournamentNameParsed() + " - Event Specific Score Sheets",
    );

    for (const j in rangeValues) {
      var eventName = rangeValues[j][0];
      Logger.log(eventName);
      var spreadSheetName =
        eventName + " Event Scoring - " + Utils.getTournamentNameParsed();
      var spreadSheetFolderId = FolderUtils.createFolderUnderRootFolder(
        scoreSheetFolderId,
        spreadSheetName,
      );
      FolderUtils.addEditorToFolder(spreadSheetFolderId, rangeValues[j].slice(1, 4));
    }
  }

}