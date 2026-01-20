using System.IO;
using Microsoft.Extensions.Logging;

namespace DesktopMcpServer
{
    public class FileLoggerProvider : ILoggerProvider
    {
        private readonly string _filePath;

        public FileLoggerProvider(string filePath)
        {
            _filePath = filePath;
        }

        public ILogger CreateLogger(string categoryName)
        {
            return new FileLogger(categoryName, _filePath);
        }

        public void Dispose() { }
    }

    public class FileLogger : ILogger
    {
        private readonly string _categoryName;
        private readonly string _filePath;
        private static readonly object _lock = new object();

        public FileLogger(string categoryName, string filePath)
        {
            _categoryName = categoryName;
            _filePath = filePath;
        }

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;

            var message = formatter(state, exception);
            var logRecord = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{logLevel}] [{_categoryName}] {message}";

            if (exception != null)
            {
                logRecord += Environment.NewLine + exception;
            }

            // Simple thread-safe write
            lock (_lock)
            {
                try
                {
                    File.AppendAllText(_filePath, logRecord + Environment.NewLine);
                }
                catch { } // Ignore file access errors to prevent app crash
            }
        }
    }
}