using System;
using System.IO;

namespace ScannerApp
{
    public static class Logger
    {
        private static string _logFilePath;
        private static readonly object _lock = new object();

        public static void Initialize(string outputPdfPath)
        {
            _logFilePath = Path.ChangeExtension(outputPdfPath, ".log");

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(_logFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        public static void Log(string message)
        {
            if (string.IsNullOrEmpty(_logFilePath))
            {
                throw new InvalidOperationException("Log file path not initialized. Call Logger.Initialize first.");
            }

            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";
            lock (_lock)
            {
                File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
            }
        }
    }
}
