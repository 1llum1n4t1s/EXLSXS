using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
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

    [Theory]
    [InlineData("lib/app/EXLSXS.Host.exe", true)]
    [InlineData("lib/app/EXLSXS.Host.dll", true)]
    [InlineData("lib/app/vsto/EXLSXS.dll.deploy", true)]
    [InlineData("lib/app/appsettings.json", false)]
    [InlineData("lib/app/EXLSXS.vsto", false)]
    public void IsPortableExecutableEntry_DetectsExecutablePayloads(string input, bool expected)
    {
        Assert.Equal(expected, UpdatePackageTrustVerifier.IsPortableExecutableEntry(input));
    }

    [Fact]
    public void IsPackagedTrustConfigurationAllowed_CurrentTrustAnchor_IsAccepted()
    {
        var verified = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "BBBB" };

        var accepted = UpdatePackageTrustVerifier.IsPackagedTrustConfigurationAllowed(
            "AAAA", "Publisher", "AA AA", "Publisher", verified, out _);

        Assert.True(accepted);
    }

    [Fact]
    public void IsPackagedTrustConfigurationAllowed_VerifiedRotatedSigner_IsAccepted()
    {
        var verified = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "BBBB" };

        var accepted = UpdatePackageTrustVerifier.IsPackagedTrustConfigurationAllowed(
            "AAAA", "Publisher", "BB BB", "Publisher", verified, out _);

        Assert.True(accepted);
    }

    [Theory]
    [InlineData("CCCC", "Publisher")]
    [InlineData("BBBB", "Other Publisher")]
    public void IsPackagedTrustConfigurationAllowed_UnverifiedTrustChange_IsRejected(
        string packagedThumbprint,
        string packagedSubject)
    {
        var verified = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "BBBB" };

        var accepted = UpdatePackageTrustVerifier.IsPackagedTrustConfigurationAllowed(
            "AAAA", "Publisher", packagedThumbprint, packagedSubject, verified, out _);

        Assert.False(accepted);
    }

    [Fact]
    public void VerifyFileSigner_UnsignedFile_IsRejected()
    {
        var path = Path.Combine(Path.GetTempPath(), $"exlsxs-unsigned-{Guid.NewGuid():N}.exe");
        File.WriteAllText(path, "not a portable executable");
        try
        {
            var settings = new UpdateSettings
            {
                ExpectedPublisherThumbprint = "AAAA",
                ExpectedPublisherSubject = "Publisher"
            };

            Assert.False(UpdatePackageTrustVerifier.VerifyFileSigner(path, settings, "fixture", out var message));
            Assert.Contains("Authenticode", message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void VerifyEmbeddedSignature_WindowsSystemBinary_IsTrusted()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");

        Assert.True(
            AuthenticodeTrustVerifier.VerifyEmbeddedSignature(systemBinary, out var failureReason),
            failureReason);
    }

    [Fact]
    public void TryVerifyEmbeddedSignature_WindowsSystemBinary_ReturnsVerifiedSigner()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");

        Assert.True(
            AuthenticodeTrustVerifier.TryVerifyEmbeddedSignature(
                systemBinary,
                out var signerCertificate,
                out var failureReason),
            failureReason);
        using (signerCertificate)
        {
            Assert.NotNull(signerCertificate);
            Assert.False(string.IsNullOrWhiteSpace(signerCertificate.Thumbprint));
        }
    }

    [Fact]
    public void VerifyEmbeddedSignature_ModifiedSignedBinary_IsRejected()
    {
        var source = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var modified = Path.Combine(Path.GetTempPath(), $"exlsxs-tampered-{Guid.NewGuid():N}.dll");
        File.Copy(source, modified);
        try
        {
            using (var stream = new FileStream(modified, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                stream.Position = Math.Min(1024, stream.Length - 1);
                var original = stream.ReadByte();
                stream.Position--;
                stream.WriteByte((byte)(original ^ 0x01));
            }

            Assert.False(AuthenticodeTrustVerifier.VerifyEmbeddedSignature(modified, out _));
        }
        finally
        {
            File.Delete(modified);
        }
    }

    [Fact]
    public void VerifyPackageFile_AllRequiredAndExecutableEntriesTrusted_IsAccepted()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-trusted-{Guid.NewGuid():N}-full.nupkg");
        CreatePackage(package, systemBinary, settings, includeTamperedEntry: false);
        try
        {
            Assert.True(
                UpdatePackageTrustVerifier.VerifyPackageFile(
                    package,
                    settings,
                    deleteUntrustedPackage: false,
                    out var message),
                message);
        }
        finally
        {
            File.Delete(package);
        }
    }

    [Fact]
    public void VerifyPackageFile_TamperedExecutableEntry_IsRejectedAndDeleted()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-untrusted-{Guid.NewGuid():N}-full.nupkg");
        CreatePackage(package, systemBinary, settings, includeTamperedEntry: true);

        Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
            package,
            settings,
            deleteUntrustedPackage: true,
            out _));
        Assert.False(File.Exists(package));
    }

    [Fact]
    public void VerifyPackageFile_DuplicateTrustConfiguration_IsRejected()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-duplicate-settings-{Guid.NewGuid():N}-full.nupkg");
        CreatePackage(package, systemBinary, settings, includeTamperedEntry: false);
        using (var archive = ZipFile.Open(package, ZipArchiveMode.Update))
        {
            var duplicate = archive.CreateEntry("lib/app/appsettings.json");
            using var writer = new StreamWriter(duplicate.Open());
            writer.Write("{}");
        }

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package,
                settings,
                deleteUntrustedPackage: false,
                out _));
        }
        finally
        {
            File.Delete(package);
        }
    }

    [Theory]
    [InlineData("https://exlsxs.kagayoi.com", "Simple", "win", "", false, true)]
    [InlineData("https://exlsxs.nephilim.jp", "Simple", "win", "", false, true)]
    [InlineData("https://attacker.example", "Simple", "win", "", false, false)]
    [InlineData("https://exlsxs.kagayoi.com", "Github", "win", "", false, false)]
    [InlineData("https://exlsxs.kagayoi.com", "Simple", "beta", "", false, false)]
    [InlineData("https://exlsxs.kagayoi.com", "Simple", "win", "secret", false, false)]
    [InlineData("https://exlsxs.kagayoi.com", "Simple", "win", "", true, false)]
    public void IsPackagedUpdateConfigurationAllowed_EnforcesReleaseEndpointContract(
        string source,
        string sourceKind,
        string channel,
        string accessToken,
        bool prerelease,
        bool expected)
    {
        var accepted = UpdatePackageTrustVerifier.IsPackagedUpdateConfigurationAllowed(
            source, sourceKind, channel, accessToken, prerelease, out _);

        Assert.Equal(expected, accepted);
    }

    [Theory]
    [InlineData("EXLSXS-1.0.8-full.nupkg", true)]
    [InlineData("../EXLSXS-1.0.8-full.nupkg", false)]
    [InlineData("subdir/EXLSXS-1.0.8-full.nupkg", false)]
    [InlineData("EXLSXS-1.0.8-full.zip", false)]
    public void TryResolvePackagePath_StaysInsideLocatorDirectory(string fileName, bool expected)
    {
        var packagesDirectory = Path.Combine(Path.GetTempPath(), "custom-install", "packages");

        var accepted = UpdatePackageTrustVerifier.TryResolvePackagePath(
            packagesDirectory, fileName, out var packagePath, out _);

        Assert.Equal(expected, accepted);
        if (accepted)
        {
            Assert.Equal(Path.Combine(Path.GetFullPath(packagesDirectory), fileName), packagePath);
        }
    }

    [Fact]
    public void TryInstallVerifiedUpdater_TrustedPayload_ReplacesTarget()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var directory = Path.Combine(Path.GetTempPath(), $"exlsxs-updater-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var updaterPath = Path.Combine(directory, "Update.exe");
        File.WriteAllText(updaterPath, "old updater");
        try
        {
            Assert.True(
                UpdatePackageTrustVerifier.TryInstallVerifiedUpdater(
                    updaterPath,
                    File.ReadAllBytes(systemBinary),
                    settings,
                    out var message),
                message);
            Assert.True(AuthenticodeTrustVerifier.VerifyEmbeddedSignature(updaterPath, out message), message);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void UpdaterReadHandle_DeniesVelopackStyleOverwrite()
    {
        var updaterPath = Path.Combine(Path.GetTempPath(), $"exlsxs-updater-lock-{Guid.NewGuid():N}.exe");
        File.WriteAllText(updaterPath, "trusted updater");
        try
        {
            using var updaterReadLock = new FileStream(updaterPath, FileMode.Open, FileAccess.Read, FileShare.Read);

            Assert.Throws<IOException>(() => File.WriteAllBytes(updaterPath, [1, 2, 3]));
        }
        finally
        {
            File.Delete(updaterPath);
        }
    }

    private static UpdateSettings CreateSettingsForSignedFile(string filePath)
    {
#pragma warning disable SYSLIB0057
        using var certificate = new X509Certificate2(X509Certificate.CreateFromSignedFile(filePath));
#pragma warning restore SYSLIB0057
        return new UpdateSettings
        {
            Source = "https://exlsxs.kagayoi.com",
            SourceKind = UpdateSourceKind.Simple,
            Channel = "win",
            ExpectedPublisherThumbprint = certificate.Thumbprint,
            ExpectedPublisherSubject = certificate.GetNameInfo(X509NameType.SimpleName, forIssuer: false)
        };
    }

    private static void CreatePackage(
        string packagePath,
        string signedFilePath,
        UpdateSettings settings,
        bool includeTamperedEntry)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Create);
        foreach (var entryName in new[]
                 {
                     "lib/app/EXLSXS.Host.exe",
                     "lib/app/EXLSXS.Host.dll",
                     "lib/app/vsto/EXLSXS.dll",
                     "lib/app/Squirrel.exe"
                 })
        {
            var entry = archive.CreateEntry(entryName);
            using var destination = entry.Open();
            using var source = File.OpenRead(signedFilePath);
            source.CopyTo(destination);
        }

        var settingsEntry = archive.CreateEntry("lib/app/appsettings.json");
        using (var stream = settingsEntry.Open())
        {
            JsonSerializer.Serialize(stream, new
            {
                Update = new
                {
                    settings.Source,
                    SourceKind = settings.SourceKind.ToString(),
                    settings.Channel,
                    settings.AccessToken,
                    settings.Prerelease,
                    settings.ExpectedPublisherThumbprint,
                    settings.ExpectedPublisherSubject
                }
            });
        }

        if (!includeTamperedEntry)
        {
            return;
        }

        var bytes = File.ReadAllBytes(signedFilePath);
        bytes[Math.Min(1024, bytes.Length - 1)] ^= 0x01;
        var tamperedEntry = archive.CreateEntry("lib/app/tampered.dll");
        using var tamperedStream = tamperedEntry.Open();
        tamperedStream.Write(bytes);
    }
}
