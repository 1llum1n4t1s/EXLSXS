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
    public void TryVerifyEmbeddedSignature_ModifiedSignedBinary_IsRejected()
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

            Assert.False(AuthenticodeTrustVerifier.TryVerifyEmbeddedSignature(modified, out var signerCertificate, out _));
            Assert.Null(signerCertificate);
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

    // Win32 の正規化で別ファイルに化ける名前は、PE 判定 (拡張子の末尾一致) をすり抜けるため拒否する。
    [Theory]
    [InlineData("lib/app/coreclr.dll.", true)]
    [InlineData("lib/app/coreclr.dll ", true)]
    [InlineData("lib/app/coreclr.dll::$DATA", true)]
    [InlineData("lib/app./appsettings.json", true)]
    [InlineData("lib/app/../coreclr.dll", true)]
    [InlineData("/lib/app/x.bat", true)]
    [InlineData("lib/app//Users/Public/x.bat", true)]
    [InlineData("lib/app/coreclr.dll.__symlink", true)]
    [InlineData("", true)]
    [InlineData("lib/app/coreclr.dll", false)]
    [InlineData("lib/app/vsto/", false)]
    [InlineData("_rels/.rels", false)]
    public void HasAmbiguousWindowsEntryName_DetectsWin32Aliases(string entryName, bool expected)
    {
        Assert.Equal(expected, UpdatePackageTrustVerifier.HasAmbiguousWindowsEntryName(entryName));
    }

    [Theory]
    [InlineData("lib/app/coreclr.dll.")]
    [InlineData("lib/app/coreclr.dll::$DATA")]
    public void VerifyPackageFile_UnsignedPayloadBehindWin32Alias_IsRejectedAndDeleted(string aliasEntryName)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-alias-{Guid.NewGuid():N}-full.nupkg");
        CreatePackage(package, systemBinary, settings, includeTamperedEntry: false);
        using (var archive = ZipFile.Open(package, ZipArchiveMode.Update))
        {
            // 必須エントリはすべて正しく署名されたまま、別名で未署名のペイロードだけを足す。
            var alias = archive.CreateEntry(aliasEntryName);
            using var stream = alias.Open();
            stream.Write("not signed"u8);
        }

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package,
                settings,
                deleteUntrustedPackage: true,
                out var message));
            Assert.Contains("not allowed", message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(package));
        }
        finally
        {
            if (File.Exists(package))
            {
                File.Delete(package);
            }
        }
    }

    // .NET は Info-ZIP Unicode Path (0x7075) を無視して readme.txt (非 PE) と読むが、
    // Update.exe の zip クレートはこの値で coreclr.dll として展開する。署名検査前にパッケージごと拒否する。
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void VerifyPackageFile_UnicodePathExtraFieldAlias_IsRejectedAndDeleted(bool inCentralDirectory)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-unicode-path-{Guid.NewGuid():N}-full.nupkg");
        var aliasField = CreateUnicodePathExtraField("lib/app/readme.txt", "lib/app/coreclr.dll");
        RawZipWriter.Write(package,
        [
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/appsettings.json", CreatePackagedSettingsJson(settings)),
            new("lib/app/readme.txt", "not signed"u8.ToArray(),
                CentralExtra: inCentralDirectory ? aliasField : [],
                LocalExtra: inCentralDirectory ? [] : aliasField)
        ]);

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package,
                settings,
                deleteUntrustedPackage: true,
                out var message));
            Assert.Contains("Unicode Path", message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(package));
        }
        finally
        {
            if (File.Exists(package))
            {
                File.Delete(package);
            }
        }
    }

    // 末尾の中央ディレクトリは正規エントリだけ、非 PE エントリの中身に「未署名 coreclr.dll を指す別の zip」を埋め込む。
    // zip クレートは末尾の読み取りに失敗すると手前の EOCD へ戻るため、戻り先がある時点で拒否する。
    // 記載 offset をずらしても (zip クレートは前方を探して中央ヘッダーを使う)、size を偽っても
    // (zip クレートは位置決めに size を使わない) 戻り先として拒否されること。
    // 総数 0・ディスク上の数 1 以上の EOCD (zip クレートは中央ヘッダーを探さずに採用し、ディスク上の数だけ読む) も含める。
    [Theory]
    [InlineData(0u, null, false)]
    [InlineData(5u, null, false)]
    [InlineData(0u, 0u, false)]
    [InlineData(0u, uint.MaxValue - 1, false)]
    [InlineData(0u, null, true)]
    public void VerifyPackageFile_EmbeddedAlternateCentralDirectory_IsRejectedAndDeleted(uint offsetShift, uint? sizeOverride, bool zeroTotalEntries)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        const string carrierName = "lib/app/notes.bin";
        // 先頭エントリのデータはローカルヘッダー (30 バイト) + 名前の直後から始まる。
        var carrierDataOffset = (uint)(30 + System.Text.Encoding.UTF8.GetByteCount(carrierName));
        var hidden = RawZipWriter.Build(
            [new("lib/app/coreclr.dll", "not signed"u8.ToArray())],
            carrierDataOffset,
            eocdOffsetShift: offsetShift,
            eocdSizeOverride: sizeOverride,
            eocdTotalEntriesOverride: zeroTotalEntries ? (ushort)0 : null);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-alt-cd-{Guid.NewGuid():N}-full.nupkg");
        RawZipWriter.Write(package,
        [
            new(carrierName, hidden),
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/appsettings.json", CreatePackagedSettingsJson(settings))
        ]);

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package,
                settings,
                deleteUntrustedPackage: true,
                out var message));
            Assert.Contains("another end of central directory", message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(package));
        }
        finally
        {
            if (File.Exists(package))
            {
                File.Delete(package);
            }
        }
    }

    // 末尾の EOCD にコメントを付けたもの・Zip64 の合図 (size の上限値) を持つものは、
    // zip クレートと .NET で読み方が分かれうるため拒否する。.NET が開けない形なら、その時点で適用されない。
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void VerifyPackageFile_EndRecordWithCommentOrZip64Marker_IsRejected(bool withComment, bool zip64SizeMarker)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var bytes = RawZipWriter.Build(
        [
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/appsettings.json", CreatePackagedSettingsJson(settings))
        ], eocdSizeOverride: zip64SizeMarker ? uint.MaxValue : null);
        if (withComment)
        {
            // EOCD のコメント長 (末尾 2 バイト) を 4 にして、コメントを足す。
            bytes[^2] = 4;
            bytes = [.. bytes, .. "note"u8.ToArray()];
        }

        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-eocd-{Guid.NewGuid():N}-full.nupkg");
        File.WriteAllBytes(package, bytes);
        try
        {
            bool accepted;
            try
            {
                accepted = UpdatePackageTrustVerifier.VerifyPackageFile(package, settings, deleteUntrustedPackage: false, out _);
            }
            catch (InvalidDataException)
            {
                accepted = false;
            }

            Assert.False(accepted);
        }
        finally
        {
            File.Delete(package);
        }
    }

    // 同梱設定の形が不正 (ルートや Update がオブジェクトでない、JSON でない) でも例外にせず、
    // 信頼設定の検証失敗として拒否・削除される (呼び出し元は Untrusted として扱える)。
    [Theory]
    [InlineData("[]")]
    [InlineData("\"text\"")]
    [InlineData("{\"Update\":[]}")]
    [InlineData("{\"Update\":\"x\"}")]
    [InlineData("not json")]
    public void VerifyPackageFile_MalformedPackagedSettingsShape_IsRejectedWithoutThrowing(string settingsJson)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-badsettings-{Guid.NewGuid():N}-full.nupkg");
        RawZipWriter.Write(package,
        [
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/appsettings.json", System.Text.Encoding.UTF8.GetBytes(settingsJson))
        ]);

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package,
                settings,
                deleteUntrustedPackage: true,
                out var message));
            Assert.Contains("Packaged publisher trust configuration", message, StringComparison.Ordinal);
            Assert.False(File.Exists(package));
        }
        finally
        {
            if (File.Exists(package))
            {
                File.Delete(package);
            }
        }
    }

    // 必須エントリの欠落と、署名済みの updater が 0 件・2 件の組み合わせは、いずれも適用しない。
    [Theory]
    [InlineData("lib/app/EXLSXS.Host.exe")]
    [InlineData("lib/app/EXLSXS.Host.dll")]
    [InlineData("lib/app/vsto/EXLSXS.dll")]
    [InlineData("lib/app/Squirrel.exe")]
    public void VerifyPackageFile_MissingRequiredEntry_IsRejected(string missingEntry)
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var entries = new[]
            {
                "lib/app/EXLSXS.Host.exe",
                "lib/app/EXLSXS.Host.dll",
                "lib/app/vsto/EXLSXS.dll",
                "lib/app/Squirrel.exe"
            }
            .Where(name => name != missingEntry)
            .Select(name => new RawZipEntry(name, signed))
            .Append(new RawZipEntry("lib/app/appsettings.json", CreatePackagedSettingsJson(settings)))
            .ToList();
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-missing-{Guid.NewGuid():N}-full.nupkg");
        RawZipWriter.Write(package, entries);

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package, settings, deleteUntrustedPackage: false, out var message));
            Assert.Contains("was not found in the update package", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(package);
        }
    }

    [Fact]
    public void VerifyPackageFile_TwoSignedUpdaterEntries_IsRejected()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-two-updaters-{Guid.NewGuid():N}-full.nupkg");
        RawZipWriter.Write(package,
        [
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/other/Squirrel.exe", signed),
            new("lib/app/appsettings.json", CreatePackagedSettingsJson(settings))
        ]);

        try
        {
            Assert.False(UpdatePackageTrustVerifier.VerifyPackageFile(
                package, settings, deleteUntrustedPackage: false, out var message));
            Assert.Contains("exactly one", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(package);
        }
    }

    // 同じ書き方で拡張フィールドを付けなければ受理されることを確かめ、上のテストが
    // 自前 zip の構造不備ではなく 0x7075 そのもので落ちていることを保証する。
    [Fact]
    public void VerifyPackageFile_RawZipWithoutExtraFields_IsAccepted()
    {
        var systemBinary = Path.Combine(Environment.SystemDirectory, "kernel32.dll");
        var settings = CreateSettingsForSignedFile(systemBinary);
        var signed = File.ReadAllBytes(systemBinary);
        var package = Path.Combine(Path.GetTempPath(), $"exlsxs-raw-{Guid.NewGuid():N}-full.nupkg");
        RawZipWriter.Write(package,
        [
            new("lib/app/EXLSXS.Host.exe", signed),
            new("lib/app/EXLSXS.Host.dll", signed),
            new("lib/app/vsto/EXLSXS.dll", signed),
            new("lib/app/Squirrel.exe", signed),
            new("lib/app/appsettings.json", CreatePackagedSettingsJson(settings)),
            new("lib/app/readme.txt", "plain text"u8.ToArray())
        ]);

        try
        {
            Assert.True(
                UpdatePackageTrustVerifier.VerifyPackageFile(package, settings, deleteUntrustedPackage: false, out var message),
                message);
        }
        finally
        {
            File.Delete(package);
        }
    }

    [Theory]
    [InlineData("https://exlsxs.kagayoi.com", "Simple", "win", "", false, true)]
    [InlineData("https://exlsxs.kagayoi.com:443", "Simple", "win", "", false, true)]
    [InlineData("https://exlsxs.kagayoi.com:8443", "Simple", "win", "", false, false)]
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
            Assert.True(AuthenticodeTrustVerifier.TryVerifyEmbeddedSignature(updaterPath, out var signerCertificate, out message), message);
            signerCertificate?.Dispose();
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
            // 製品と同じヘルパーでロックを取る (FileShare を緩めるとこのテストが失敗する)。
            using var updaterReadLock = UpdateChecker.OpenUpdaterWriteDenyHandle(updaterPath);

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
            stream.Write(CreatePackagedSettingsJson(settings));
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

    private static byte[] CreatePackagedSettingsJson(UpdateSettings settings)
    {
        return JsonSerializer.SerializeToUtf8Bytes(new
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

    // Info-ZIP Unicode Path 拡張フィールド: version(1) + ヘッダー名の CRC-32 + UTF-8 の名前。
    // zip クレートは CRC がヘッダー名と一致するとき、この名前でヘッダー名を置き換える。
    private static byte[] CreateUnicodePathExtraField(string headerName, string aliasName)
    {
        var alias = System.Text.Encoding.UTF8.GetBytes(aliasName);
        var data = new byte[1 + 4 + alias.Length];
        data[0] = 1;
        BitConverter.TryWriteBytes(data.AsSpan(1), RawZipWriter.Crc32(System.Text.Encoding.UTF8.GetBytes(headerName)));
        alias.CopyTo(data, 5);

        var field = new byte[4 + data.Length];
        BitConverter.TryWriteBytes(field.AsSpan(0), (ushort)0x7075);
        BitConverter.TryWriteBytes(field.AsSpan(2), (ushort)data.Length);
        data.CopyTo(field, 4);
        return field;
    }

    private sealed record RawZipEntry(string Name, byte[] Data, byte[]? CentralExtra = null, byte[]? LocalExtra = null);

    // 拡張フィールドを自由に付けるため、無圧縮 (stored) の zip を手で組み立てる最小限の書き出し。
    private static class RawZipWriter
    {
        public static void Write(string path, IReadOnlyList<RawZipEntry> entries)
            => File.WriteAllBytes(path, Build(entries));

        // baseOffset: 組み立てた zip を別ファイルのこの位置へ埋め込む前提で、記録する位置をずらす。
        // eocdOffsetShift / eocdSizeOverride: EOCD に記載する中央ディレクトリの offset / size を細工する。
        public static byte[] Build(
            IReadOnlyList<RawZipEntry> entries,
            uint baseOffset = 0,
            uint eocdOffsetShift = 0,
            uint? eocdSizeOverride = null,
            ushort? eocdTotalEntriesOverride = null)
        {
            using var output = new MemoryStream();
            using var writer = new BinaryWriter(output);
            var offsets = new List<uint>();
            foreach (var entry in entries)
            {
                var name = System.Text.Encoding.UTF8.GetBytes(entry.Name);
                var localExtra = entry.LocalExtra ?? [];
                offsets.Add(baseOffset + (uint)output.Position);
                writer.Write(0x04034b50u);
                writer.Write((ushort)20);
                writer.Write((ushort)0x0800);
                writer.Write((ushort)0);
                writer.Write(0u);
                writer.Write(Crc32(entry.Data));
                writer.Write((uint)entry.Data.Length);
                writer.Write((uint)entry.Data.Length);
                writer.Write((ushort)name.Length);
                writer.Write((ushort)localExtra.Length);
                writer.Write(name);
                writer.Write(localExtra);
                writer.Write(entry.Data);
            }

            var centralStartInBuffer = (uint)output.Position;
            var centralStart = baseOffset + centralStartInBuffer;
            for (var i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                var name = System.Text.Encoding.UTF8.GetBytes(entry.Name);
                var centralExtra = entry.CentralExtra ?? [];
                writer.Write(0x02014b50u);
                writer.Write((ushort)20);
                writer.Write((ushort)20);
                writer.Write((ushort)0x0800);
                writer.Write((ushort)0);
                writer.Write(0u);
                writer.Write(Crc32(entry.Data));
                writer.Write((uint)entry.Data.Length);
                writer.Write((uint)entry.Data.Length);
                writer.Write((ushort)name.Length);
                writer.Write((ushort)centralExtra.Length);
                writer.Write((ushort)0);
                writer.Write((ushort)0);
                writer.Write((ushort)0);
                writer.Write(0u);
                writer.Write(offsets[i]);
                writer.Write(name);
                writer.Write(centralExtra);
            }

            var centralSize = (uint)output.Position - centralStartInBuffer;
            writer.Write(0x06054b50u);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)entries.Count);
            writer.Write(eocdTotalEntriesOverride ?? (ushort)entries.Count);
            writer.Write(eocdSizeOverride ?? centralSize);
            writer.Write(centralStart - eocdOffsetShift);
            writer.Write((ushort)0);
            writer.Flush();
            return output.ToArray();
        }

        public static uint Crc32(byte[] data)
        {
            var crc = 0xFFFFFFFFu;
            foreach (var value in data)
            {
                crc ^= value;
                for (var bit = 0; bit < 8; bit++)
                {
                    crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
                }
            }

            return ~crc;
        }
    }
}
