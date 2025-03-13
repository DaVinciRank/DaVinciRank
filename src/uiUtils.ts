export class UiUtils {
  /**
   * Displays a prompt to the user and returns the response.
   * @param {string} prompt - The prompt message.
   * @returns {string} - The user's response.
   */
  static showPrompt(prompt: string): string | null {
    var ui = SpreadsheetApp.getUi();

    var result = ui.prompt(
      prompt ? prompt : "Give an input",
      "Input:",
      ui.ButtonSet.OK_CANCEL,
    );

    var button = result.getSelectedButton();
    var response = result.getResponseText();
    if (button == ui.Button.OK) {
      // call function and pass the value
      Logger.log(response);
      return response;
    } else {
      return null;
    }
  }
}
