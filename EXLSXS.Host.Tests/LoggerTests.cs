using Xunit;

namespace EXLSXS.Host.Tests;

public class LoggerTests
{
    // 2 つのホストプロセスが同じログへ追記する状況を、同じファイルへの 2 つの書き込みハンドルで再現する。
    // 書き込み位置がハンドルごとに保持されるため、末尾へ移動せずに書くと相手の行を上書きして失う。
    [Fact]
    public void AppendLine_TwoHandlesOnSameLog_KeepsEveryLine()
    {
        var logPath = Path.Combine(Path.GetTempPath(), $"exlsxs-log-{Guid.NewGuid():N}.log");
        try
        {
            using (var first = Logger.OpenLogWriter(logPath))
            using (var second = Logger.OpenLogWriter(logPath))
            {
                Logger.AppendLine(first, "first-1");
                Logger.AppendLine(second, "second-1");
                Logger.AppendLine(first, "first-2");
                Logger.AppendLine(second, "second-2");
            }

            Assert.Equal(["first-1", "second-1", "first-2", "second-2"], File.ReadAllLines(logPath));
        }
        finally
        {
            File.Delete(logPath);
        }
    }
}
