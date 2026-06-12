using Xunit;

namespace EXLSXS.Host.Tests;

public class UpdateSettingsTests
{
    [Fact]
    public void Defaults_AreSensible()
    {
        var settings = new UpdateSettings();

        Assert.Equal("win", settings.Channel);
        Assert.Equal(UpdateSourceKind.Auto, settings.SourceKind);
        Assert.False(settings.IsConfigured);
        Assert.False(settings.HasPublisherTrustConfiguration);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData("https://exlsxs.nephilim.jp", true)]
    public void IsConfigured_TracksSource(string source, bool expected)
    {
        var settings = new UpdateSettings { Source = source };

        Assert.Equal(expected, settings.IsConfigured);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData("6285702C9AF1FCFE3D9FE815B7F7F625508130C0", true)]
    public void HasPublisherTrustConfiguration_TracksThumbprint(string thumbprint, bool expected)
    {
        var settings = new UpdateSettings { ExpectedPublisherThumbprint = thumbprint };

        Assert.Equal(expected, settings.HasPublisherTrustConfiguration);
    }
}
