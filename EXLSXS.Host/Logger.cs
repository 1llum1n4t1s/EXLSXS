using System.Diagnostics;

namespace EXLSXS.Host;

internal enum LogLevel
{
    Debug,
    Info,
    Warning,
    Error
}

internal static class Logger
{
    private static readonly object Gate = new();
    private static StreamWriter? writer;
    private static bool writerFailed;

    public static void Log(string message, LogLevel level = LogLevel.Info)
    {
        var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] {message}";

        lock (Gate)
        {
            try
            {
                var currentWriter = GetWriter();
                if (currentWriter != null)
                {
                    currentWriter.WriteLine(line);
                    return;
                }
            }
            catch
            {
                writerFailed = true;
            }
        }

        Trace.WriteLine(line);
    }

    public static void LogException(string message, Exception exception)
    {
        Log($"{message} {exception}", LogLevel.Error);
    }

    public static void Dispose()
    {
        lock (Gate)
        {
            writer?.Dispose();
            writer = null;
        }
    }

    private static StreamWriter? GetWriter()
    {
        if (writer != null)
        {
            return writer;
        }

        if (writerFailed)
        {
            return null;
        }

        try
        {
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EXLSXS",
                "Logs");
            Directory.CreateDirectory(logDir);

            var logPath = Path.Combine(logDir, "EXLSXS.Host.log");
            writer = new StreamWriter(File.Open(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            {
                AutoFlush = true
            };
            return writer;
        }
        catch
        {
            writerFailed = true;
            return null;
        }
    }
}
