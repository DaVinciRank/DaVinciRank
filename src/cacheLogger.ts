export class CacheLogger {
    static appendLog(message: string): string{
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
}