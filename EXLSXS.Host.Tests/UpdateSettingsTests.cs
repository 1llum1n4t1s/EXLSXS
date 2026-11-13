using Xunit;

namespace EXLSXS.Host.Tests;

public class UpdateSettingsTests
{
    // 発行元の拇印が空白だけの設定を「信頼設定あり」と誤認すると、署名者照合なしで更新を受け入れうる。
    [Theory]
    [InlineData("   ", false)]
    [InlineData("6285702C9AF1FCFE3D9FE815B7F7F625508130C0", true)]
    public void HasPublisherTrustConfiguration_TracksThumbprint(string thumbprint, bool expected)
    {
        var settings = new UpdateSettings { ExpectedPublisherThumbprint = thumbprint };

        Assert.Equal(expected, settings.HasPublisherTrustConfiguration);
    }
}
