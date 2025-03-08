import { SpreadsheetUtils } from "./spreadsheetUtils";
import { FolderUtils } from "./folderUtils";
import { CacheLogger } from "./cacheLogger";

export class Slides {
  /**
   * Finds the slide show presentation ID.
   * @returns {string} - The slide show presentation ID.
   */
  static findSlideShowPresentation(): string {
    var parentFolderId = FolderUtils.getParentFolderId();
    var files = FolderUtils.getFilesUnderRootFolder(parentFolderId);
    // var division = currentSheet.getRangeByName("Division").getValue();
    var files = FolderUtils.getTemplateFilesWithSubstring("Medals", files);
    // var files = FolderUtils.getTemplateFilesWithSubstring(division, files)
    return files[0].getId();
  }

  /**
   * Retrieves data corresponding to an event name.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} spreadsheet - The spreadsheet.
   * @param {string} eventName - The event name.
   * @param {number} maxVal - The maximum value.
   * @returns {string[] | bolean} - The data corresponding to the event name. Returns false is event is not done scoring.
   */
  static getDataCorrespondingToEventName(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    eventName: string,
    maxVal: number,
  ): string[] | boolean {
    const rowNum = SpreadsheetUtils.findCellRowWithTextInSheet(
      sheet,
      eventName,
    );
    if (!rowNum || typeof rowNum !== "number") {
      Logger.log("Event name not found in the spreadsheet.");
      return [];
    }

    const isEventCompleted = Slides.getCellValueByColumnRowAndOffset(
      sheet,
      "G",
      rowNum,
      0,
    );

    // If the event is not marked as complete, then skip slides
    if (!isEventCompleted) {
      return false;
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
        Slides.getCellValueByColumnRowAndOffset(sheet, "A", rowNum, i) +
          "  \t" +
          Slides.getCellValueByColumnRowAndOffset(sheet, "B", rowNum, i) +
          "\t" +
          Slides.getCellValueByColumnRowAndOffset(sheet, "C", rowNum, i),
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
  static getCellValueByColumnRowAndOffset(
    spreadsheet: GoogleAppsScript.Spreadsheet.Sheet,
    column: string,
    row: number,
    offset: number,
  ): string | boolean {
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
  static removeSlidesAfterIndex(
    nIndex: number,
    deck: GoogleAppsScript.Slides.Presentation,
  ): void {
    const slides = deck.getSlides();
    slides.slice(nIndex).forEach((s) => s.remove());
  }

  /**
   * Creates one slide per row in the "Final Rankings" sheet, ensuring smart appending.
   */
  static createOneSlidePerRow() {
    // Google Slides presentation.
    const masterDeckID = Slides.findSlideShowPresentation();
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
    const eventSlideTemplate: GoogleAppsScript.Slides.Slide = slides.find(
      (slide) =>
        slide
          .getNotesPage()
          .getSpeakerNotesShape()
          .getText()
          .asString()
          .includes("Event Slide Template"),
    );

    const finalRankingSlideTemplate: GoogleAppsScript.Slides.Slide =
      slides.find((slide) =>
        slide
          .getNotesPage()
          .getSpeakerNotesShape()
          .getText()
          .asString()
          .includes("Final Ranking Slide Template"),
      );

    if (!eventSlideTemplate || !finalRankingSlideTemplate) {
      throw new Error(
        "Event Slide Template or Final Ranking Slide Template not found.",
      );
    }

    const currentEventSlides = slides.filter((slide) =>
      slide
        .getNotesPage()
        .getSpeakerNotesShape()
        .getText()
        .asString()
        .includes("Event Slide:"),
    );

    const eventSlidePosition = slides.indexOf(eventSlideTemplate) + 1;

    // Iterate over the event names and add new slides, ensuring no duplicate event slides
    eventNames.reverse().forEach((eventName) => {
      const eventData = Slides.getDataCorrespondingToEventName(
        currentSheet,
        eventName,
        5,
      );

      if (!eventData || !Array.isArray(eventData)) {
        return;
      }

      // Check if a slide with the same tag already exists
      const existingEventSlide = currentEventSlides.find((slide) =>
        slide
          .getNotesPage()
          .getSpeakerNotesShape()
          .getText()
          .asString()
          .includes(`Event Slide: ${eventName}`),
      );

      if (existingEventSlide) {
        // Skip creating the slide if it already exists
        CacheLogger.appendLog(
          `Event slide for "${eventName}" already exists, skipping.`,
        );
        return;
      }

      CacheLogger.appendLog("Adding slides for " + eventName);

      // Create a new event slide by duplicating the template
      const newEventSlide = eventSlideTemplate.duplicate();
      newEventSlide.setSkipped(false);

      // Populate data in the slide
      newEventSlide.replaceAllText("EVENT_NAME", eventName);
      newEventSlide.replaceAllText("1. __", eventData[0]);
      newEventSlide.replaceAllText("2. __", eventData[1]);
      newEventSlide.replaceAllText("3. __", eventData[2]);
      newEventSlide.replaceAllText("4. __", eventData[3]);

      // Set the tag for this new slide
      newEventSlide
        .getNotesPage()
        .getSpeakerNotesShape()
        .getText()
        .setText("Event Slide: " + eventName);

      // Move the new event slide to the correct position (after the last event slide and before the final ranking slide)
      const newSlideIndex = eventSlidePosition + currentEventSlides.length + 1;
      newEventSlide.move(newSlideIndex);
    });

    // Create the final ranking slide
    const finalRankingData = Slides.getDataCorrespondingToEventName(
      currentSheet,
      "Overall Team Results",
      9,
    );

    if (!finalRankingData || !Array.isArray(finalRankingData)) {
      CacheLogger.appendLog("Skipping slides for overall results");
      return;
    }

    // Add the final ranking slide
    const newFinalRankingSlide = finalRankingSlideTemplate.duplicate();
    newFinalRankingSlide.setSkipped(false);

    newFinalRankingSlide.replaceAllText("1. __", finalRankingData[0]);
    newFinalRankingSlide.replaceAllText("2. __", finalRankingData[1]);
    newFinalRankingSlide.replaceAllText("3. __", finalRankingData[2]);
    newFinalRankingSlide.replaceAllText("4. __", finalRankingData[3]);
    newFinalRankingSlide.replaceAllText("5. __", finalRankingData[4]);
    newFinalRankingSlide.replaceAllText("6. __", finalRankingData[5]);
    newFinalRankingSlide.replaceAllText("7. __", finalRankingData[6]);
    newFinalRankingSlide.replaceAllText("8. __", finalRankingData[7]);

    // Set the tag for the final ranking slide
    newFinalRankingSlide
      .getNotesPage()
      .getSpeakerNotesShape()
      .getText()
      .setText("Final Ranking Slide");
  }
}
