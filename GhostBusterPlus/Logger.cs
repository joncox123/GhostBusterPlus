using System;
using System.IO;
using System.Text;

namespace ScreenRefreshApp
{
    /// <summary>
    /// Static class that handles application logging to a file
    /// </summary>
    public static class Logger
    {
        private static StreamWriter logWriter;
        private static string logFilePath;
        
        /// <summary>
        /// Master flag to enable/disable logging
        /// </summary>
        public static bool ENABLE_FILE_LOG = true;
        
        /// <summary>
        /// Initializes the logger and creates/overwrites the log file
        /// </summary>
        public static void Initialize()
        {
            if (!ENABLE_FILE_LOG) return;
            
            try
            {
                // Create log directory in AppData\Local\GhostBusterPlus
                string logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GhostBusterPlus");
                
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                
                // Set up log file path
                logFilePath = Path.Combine(logDir, "GhostBusterPlus.log");
                
                // Create or overwrite the log file with UTF-8 encoding
                logWriter = new StreamWriter(logFilePath, false, Encoding.UTF8);
                logWriter.AutoFlush = true;
                
                // Write initial log entry
                Log($"GhostBusterPlus started");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to initialize logger: {ex.Message}");
                ENABLE_FILE_LOG = false;
            }
        }
        
        /// <summary>
        /// Logs a message with timestamp
        /// </summary>
        /// <param name="message">Message to log</param>
        public static void Log(string message)
        {
            if (!ENABLE_FILE_LOG || logWriter == null) return;
            
            try
            {
                // Format: [yyyy-MM-dd HH:mm:ss.fff] Message
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                logWriter.WriteLine(logEntry);
                logWriter.Flush();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging failed: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Closes the log file
        /// </summary>
        public static void Close()
        {
            if (!ENABLE_FILE_LOG || logWriter == null) return;
            
            try
            {
                Log("GhostBusterPlus shutting down");
                logWriter.Flush();
                logWriter.Close();
                logWriter.Dispose();
                logWriter = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to close log file: {ex.Message}");
            }
        }
    }
}