export class CacheLogger {
    static appendLog(message: string, only_log_if_debug: boolean | null = null): string{
        if (only_log_if_debug && !CacheLogger.isDebugMode()) {
            return "";
        }
        var cache = CacheService.getUserCache();
        var existingLogs = cache.get("logs") || ""; // Retrieve stored logs
        var newLogs = existingLogs + message + "\n"; // Append new message
        cache.put("logs", newLogs, 600); // Store logs with a 10-minute expiry
        return newLogs;
    }
    
    static getLogs(): string {
        var cache = CacheService.getUserCache();
        return cache.get("logs") || ""; // Return stored logs
    }
    
    static clearLogs(): string {
        var cache = CacheService.getUserCache();
        cache.remove("logs"); 
        return ""; 
    }

    /**
     * Sets the debug mode.  Typically called from a separate setup script or manually
     * via the Script properties in the Apps Script editor.
     * @param {boolean} debug - True to enable debug logging, false to disable.
     */
    static setDebugMode(debug: boolean) {
        PropertiesService.getScriptProperties().setProperty('DEBUG_MODE', String(debug));
    }

    /**
     * Gets the debug mode.
     * @return {boolean} - True if debug logging is enabled, false otherwise.
     */
    static isDebugMode() {
        const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE');
        return debugMode === 'true'; // Explicit comparison to string
    }
}