import { findCellRowWithTextInSheet } from "./spreadsheetUtils";
import {
  getParentFolderId,
  getFilesUnderRootRolder,
  getTemplateFilesWithSubstring,
} from "./folderUtils";

/**
 * Finds the slide show presentation ID.
 * @returns {string} - The slide show presentation ID.
 */
function findSlideShowPresentation(): string {
  var parentFolderId = getParentFolderId();
  var files = getFilesUnderRootRolder(parentFolderId);
  // var division = currentSheet.getRangeByName("Division").getValue();
  var files = getTemplateFilesWithSubstring("Medals", files);
  // var files = getTemplateFilesWithSubstring(division, files)
  return files[0].getId();
}

/**
 * Retrieves data corresponding to an event name.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} spreadsheet - The spreadsheet.
 * @param {string} eventName - The event name.
 * @param {number} maxVal - The maximum value.
 * @returns {string[]} - The data corresponding to the event name.
 */
function getDataCorrespondingToEventName(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  eventName: string,
  maxVal: number,
): string[] {
  const rowNum = findCellRowWithTextInSheet(sheet, eventName);
  if (!rowNum || typeof rowNum !== "number") {
    Logger.log("Event name not found in the spreadsheet.");
    return [];
  }
  /*
  1st: A{row+2} & B{row+2} & C{row+2}
  2nd: A{row+3} & B{row+3} & C{row+3}
  3rd: A{row+4} & B{row+4} & C{row+4}
  4th: A{row+5} & B{row+5} & C{row+5}
  */
  var entryList = [];
  for (var i = 2; i <= maxVal; i++) {
    entryList.push(
      getCellValueByColumnRowAndOffset(sheet, "A", rowNum, i) +
        "\t\t" +
        getCellValueByColumnRowAndOffset(sheet, "B", rowNum, i) +
        "\t" +
        getCellValueByColumnRowAndOffset(sheet, "C", rowNum, i),
    );
  }
  return entryList;
}

/**
 * Retrieves the cell value by column, row, and offset.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet.
 * @param {string} column - The column.
 * @param {number} row - The row.
 * @param {number} offset - The offset.
 * @returns {string} - The cell value.
 */
function getCellValueByColumnRowAndOffset(
  spreadsheet: GoogleAppsScript.Spreadsheet.Sheet,
  column: string,
  row: number,
  offset: number,
): string {
  // double check this logic
  return spreadsheet
    .getRange(column + (row + offset) + ":" + column + (row + offset))
    .getValues()[0][0];
}

/**
 * Removes slides after a specified index.
 * @param {number} nIndex - The index after which slides will be removed.
 * @param {GoogleAppsScript.Slides.Presentation} deck - The presentation deck.
 */
function removeSlidesAfterIndex(
  nIndex: number,
  deck: GoogleAppsScript.Slides.Presentation,
): void {
  const slides = deck.getSlides();
  slides.slice(nIndex).forEach((s) => s.remove());
}

/**
 * Creates one slide per row in the "Final Rankings" sheet.
 */
function createOneSlidePerRow() {
  // Replace <INSERT_SLIDE_DECK_ID> wih the ID of your
  // Google Slides presentation.
  const masterDeckID = findSlideShowPresentation();
  // Open the presentation and get the slides in it.
  const deck = SlidesApp.openById(masterDeckID);
  const slides: GoogleAppsScript.Slides.Slide[] = deck.getSlides();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
    spreadsheet.getSheetByName("Final Rankings");
  if (!currentSheet || typeof currentSheet === "undefined") {
    throw new Error("Final Ranking Sheet not found in the spreadsheet.");
  }

  var range = spreadsheet.getRangeByName("Events");
  if (!range) {
    throw new Error("Range 'Events' not found in the spreadsheet.");
  }
  var values = range.getValues();
  var eventNames = values.flat().filter(function (cell) {
    return cell !== "";
  });

  // The 2nd slide is the template that will be duplicated
  // once per row in the spreadsheet.
  const eventSlides: GoogleAppsScript.Slides.Slide = slides[1];
  const teamSlides: GoogleAppsScript.Slides.Slide = slides[2];
  eventSlides.setSkipped(true);
  teamSlides.setSkipped(true);

  removeSlidesAfterIndex(3, deck);

  for (var i = eventNames.length - 1; i >= 0; i--) {
    const eventName = eventNames[i];
    const eventData = getDataCorrespondingToEventName(
      currentSheet,
      eventName,
      5,
    );

    const slide = eventSlides.duplicate();
    slide.setSkipped(false);

    // Populate data in the slide that was created
    slide.replaceAllText("EVENT_NAME", eventName);
    slide.replaceAllText("1. __", eventData[0]);
    slide.replaceAllText("2. __", eventData[1]);
    slide.replaceAllText("3. __", eventData[2]);
    slide.replaceAllText("4. __", eventData[3]);
  }

  // Create the final ranking slide
  const slide = teamSlides.duplicate();
  slide.setSkipped(false);
  const eventData = getDataCorrespondingToEventName(
    currentSheet,
    "Overall Team Results",
    9,
  );

  slide.replaceAllText("1. __", eventData[0]);
  slide.replaceAllText("2. __", eventData[1]);
  slide.replaceAllText("3. __", eventData[2]);
  slide.replaceAllText("4. __", eventData[3]);
  slide.replaceAllText("5. __", eventData[4]);
  slide.replaceAllText("6. __", eventData[5]);
  slide.replaceAllText("7. __", eventData[6]);
  slide.replaceAllText("8. __", eventData[7]);

  teamSlides.move(2);
}
