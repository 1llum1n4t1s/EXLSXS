using Xunit;

namespace EXLSXS.Host.Tests;

// 自動更新のセキュリティゲートを支える純粋関数を固定化する。
// ここが壊れると、署名者の照合やエントリ名の正規化が誤り、未署名/別署名パッケージの
// サイレント適用 (RCE 級) につながりうるため、境界値を含めて検証する。
public class UpdatePackageTrustVerifierTests
{
    [Theory]
    [InlineData("62:85:70:2C", "6285702C")]
    [InlineData("62 85 70 2c", "6285702C")]
    [InlineData("6285702c", "6285702C")]
    [InlineData("0x6285702C", "06285702C")] // 16 進数字のみ抽出するため 'x' は除去され '0' は残る
    [InlineData("", "")]
    public void NormalizeThumbprint_StripsSeparatorsAndUppercases(string input, string expected)
    {
        Assert.Equal(expected, UpdatePackageTrustVerifier.NormalizeThumbprint(input));
    }

    [Theory]
    [InlineData("lib/net48/EXLSXS.Host.dll", "/lib/net48/EXLSXS.Host.dll")]
    [InlineData("lib\\net48\\EXLSXS.Host.dll", "/lib/net48/EXLSXS.Host.dll")]
    [InlineData("/EXLSXS.dll", "/EXLSXS.dll")]
    [InlineData("EXLSXS.dll", "/EXLSXS.dll")]
    public void NormalizeEntryName_NormalizesSeparatorsAndLeadingSlash(string input, string expected)
    {
        Assert.Equal(expected, UpdatePackageTrustVerifier.NormalizeEntryName(input));
    }
}
