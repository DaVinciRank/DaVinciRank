import { UiUtils } from "../src/uiUtils";

test("showPrompt should return user response", () => {
  // Mock the SpreadsheetApp and Logger
  global.SpreadsheetApp = {
    getUi: jest.fn().mockReturnValue({
      prompt: jest.fn().mockReturnValue({
        getSelectedButton: jest.fn().mockReturnValue("OK"),
        getResponseText: jest.fn().mockReturnValue("Test Response"),
      }),
      ButtonSet: {
        OK_CANCEL: "OK_CANCEL",
      },
      Button: {
        OK: "OK",
      },
    }),
  } as unknown as typeof SpreadsheetApp;
  global.Logger = {
    log: jest.fn(),
  } as unknown as typeof Logger;

  const response = UiUtils.showPrompt("Test Prompt");
  expect(response).toBe("Test Response");
});
