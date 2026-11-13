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
    // ログオン時に Run (--update-check) と RunOnce (--register) のホストが同時に起動しうる。
    // .NET の FileMode.Append はハンドルごとの位置へ書くため、別プロセスの追記を上書きしてしまう。
    // 同じユーザーセッション内のホスト同士をこの名前付きミューテックスで直列化する。
    private const string CrossProcessMutexName = @"Local\EXLSXS.Host.Log";
    private static readonly TimeSpan CrossProcessWaitTimeout = TimeSpan.FromSeconds(2);

    private static readonly Lock Gate = new();
    private static StreamWriter? writer;
    private static bool writerFailed;
    private static Mutex? crossProcessMutex;
    private static bool crossProcessMutexFailed;

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
                    var acquired = TryEnterCrossProcessLock();
                    try
                    {
                        AppendLine(currentWriter, line);
                    }
                    finally
                    {
                        if (acquired)
                        {
                            crossProcessMutex!.ReleaseMutex();
                        }
                    }

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
            crossProcessMutex?.Dispose();
            crossProcessMutex = null;
        }
    }

    // 他プロセスが書いた分の後ろへ確実に追記するため、書く直前にファイル末尾へ移動する。
    // AutoFlush のため StreamWriter 内のバッファは常に空で、BaseStream を直接動かしても内容は失われない。
    internal static void AppendLine(StreamWriter target, string line)
    {
        target.BaseStream.Seek(0, SeekOrigin.End);
        target.WriteLine(line);
    }

    internal static StreamWriter OpenLogWriter(string logPath)
        => new(File.Open(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        {
            AutoFlush = true
        };

    // ミューテックスを使えない環境でもログ自体は止めない (末尾への移動だけでも上書きの範囲は狭まる)。
    private static bool TryEnterCrossProcessLock()
    {
        if (crossProcessMutexFailed)
        {
            return false;
        }

        try
        {
            crossProcessMutex ??= new Mutex(false, CrossProcessMutexName);
            return crossProcessMutex.WaitOne(CrossProcessWaitTimeout);
        }
        catch (AbandonedMutexException)
        {
            // 前の所有プロセスが解放せずに終了した。所有権はこのスレッドに移っている。
            return true;
        }
        catch
        {
            crossProcessMutexFailed = true;
            return false;
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
            writer = OpenLogWriter(logPath);
            return writer;
        }
        catch
        {
            writerFailed = true;
            return null;
        }
    }
}
