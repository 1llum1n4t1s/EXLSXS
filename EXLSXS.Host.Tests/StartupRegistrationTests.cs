using Xunit;

namespace EXLSXS.Host.Tests;

public class StartupRegistrationTests
{
    [Theory]
    // Run キーの値は空白を含むパスでも実行ファイルと引数が分かれるよう、パスを引用符で囲む契約。
    [InlineData(@"C:\Program Files\EXLSXS\EXLSXS.Host.exe", "--update-check", "\"C:\\Program Files\\EXLSXS\\EXLSXS.Host.exe\" --update-check")]
    public void BuildCommand_QuotesExecutablePath(string exePath, string argument, string expected)
    {
        Assert.Equal(expected, StartupRegistration.BuildCommand(exePath, argument));
    }

    [Theory]
    [InlineData("", "--register")]
    [InlineData(@"C:\Apps\EXLSXS.Host.exe", "")]
    public void BuildCommand_RejectsEmptyParts(string exePath, string argument)
    {
        Assert.Throws<ArgumentException>(() => StartupRegistration.BuildCommand(exePath, argument));
    }
}
