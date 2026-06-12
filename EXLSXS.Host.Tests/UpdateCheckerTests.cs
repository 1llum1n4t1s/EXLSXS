using Xunit;

namespace EXLSXS.Host.Tests;

public class UpdateCheckerTests
{
    [Fact]
    public void ShouldUseGithubSource_ExplicitGithub_IsTrue()
    {
        // SourceKind を明示している場合は URL のホストに関わらずそれを優先する。
        var settings = new UpdateSettings { SourceKind = UpdateSourceKind.Github, Source = "https://exlsxs.nephilim.jp" };

        Assert.True(UpdateChecker.ShouldUseGithubSource(settings));
    }

    [Fact]
    public void ShouldUseGithubSource_ExplicitSimple_IsFalse()
    {
        var settings = new UpdateSettings { SourceKind = UpdateSourceKind.Simple, Source = "https://github.com/owner/repo" };

        Assert.False(UpdateChecker.ShouldUseGithubSource(settings));
    }

    [Theory]
    [InlineData("https://github.com/owner/repo", true)]
    [InlineData("https://exlsxs.nephilim.jp", false)]
    [InlineData("not a url", false)]
    public void ShouldUseGithubSource_Auto_DetectsGithubHost(string source, bool expected)
    {
        var settings = new UpdateSettings { SourceKind = UpdateSourceKind.Auto, Source = source };

        Assert.Equal(expected, UpdateChecker.ShouldUseGithubSource(settings));
    }
}
