import { CacheLogger } from "../src/cacheLogger";

const mockCache = {
  get: jest.fn(),
  put: jest.fn(),
  remove: jest.fn(),
};

const mockProperties = {
  getProperty: jest.fn(),
  setProperty: jest.fn(),
};

global.CacheService = {
  getUserCache: jest.fn(() => mockCache),
} as unknown as GoogleAppsScript.Cache.CacheService;

global.PropertiesService = {
  getScriptProperties: jest.fn(() => mockProperties),
} as unknown as GoogleAppsScript.Properties.PropertiesService;

describe("CacheLogger", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("appendLog should store logs and return updated logs", () => {
    mockCache.get.mockReturnValue("Existing log\n");

    const result = CacheLogger.appendLog("New log");

    expect(mockCache.get).toHaveBeenCalledWith("logs");
    expect(mockCache.put).toHaveBeenCalledWith(
      "logs",
      "Existing log\nNew log\n",
      600,
    );
    expect(result).toBe("Existing log\nNew log\n");
  });

  test("appendLog should return empty string if debug mode is off and only_log_if_debug is true", () => {
    jest.spyOn(CacheLogger, "isDebugMode").mockReturnValue(false);

    const result = CacheLogger.appendLog("Debug log", true);

    expect(mockCache.get).not.toHaveBeenCalled();
    expect(mockCache.put).not.toHaveBeenCalled();
    expect(result).toBe("");
  });

  test("getLogs should return stored logs", () => {
    mockCache.get.mockReturnValue("Stored logs");

    const result = CacheLogger.getLogs();

    expect(mockCache.get).toHaveBeenCalledWith("logs");
    expect(result).toBe("Stored logs");
  });

  test("getLogs should return an empty string if no logs exist", () => {
    mockCache.get.mockReturnValue(null);

    const result = CacheLogger.getLogs();

    expect(result).toBe("");
  });

  test("clearLogs should remove stored logs", () => {
    const result = CacheLogger.clearLogs();

    expect(mockCache.remove).toHaveBeenCalledWith("logs");
    expect(result).toBe("");
  });

  test("setDebugMode should store the debug mode in properties", () => {
    CacheLogger.setDebugMode(true);

    expect(mockProperties.setProperty).toHaveBeenCalledWith(
      "DEBUG_MODE",
      "true",
    );

    CacheLogger.setDebugMode(false);

    expect(mockProperties.setProperty).toHaveBeenCalledWith(
      "DEBUG_MODE",
      "false",
    );
  });

  test("isDebugMode should return false if DEBUG_MODE is set to 'false' or undefined", () => {
    mockProperties.getProperty.mockReturnValue("false");
    expect(CacheLogger.isDebugMode()).toBe(false);

    mockProperties.getProperty.mockReturnValue(null);
    expect(CacheLogger.isDebugMode()).toBe(false);
  });
});
