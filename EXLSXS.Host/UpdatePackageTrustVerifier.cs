using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Velopack;

namespace EXLSXS.Host;

internal static class UpdatePackageTrustVerifier
{
    private sealed record RequiredSignedEntry(string Label, Func<string, bool> Matches);

    private static readonly RequiredSignedEntry[] RequiredEntries =
    [
        new("EXLSXS host", name => name.EndsWith("/EXLSXS.Host.exe", StringComparison.OrdinalIgnoreCase)),
        new("EXLSXS host assembly", name => name.EndsWith("/EXLSXS.Host.dll", StringComparison.OrdinalIgnoreCase)),
        new("EXLSXS VSTO assembly", name => name.EndsWith("/EXLSXS.dll.deploy", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith("/EXLSXS.dll", StringComparison.OrdinalIgnoreCase))
    ];

    public static bool Verify(UpdateInfo updateInfo, UpdateSettings settings, out string message)
    {
        if (!settings.HasPublisherTrustConfiguration)
        {
            message = "Expected publisher thumbprint is not configured.";
            return false;
        }

        var packageFileName = updateInfo.TargetFullRelease.FileName;
        var packagePath = ResolvePackagePath(packageFileName);
        if (!File.Exists(packagePath))
        {
            message = $"Downloaded update package was not found: {packageFileName}";
            return false;
        }

        var expectedThumbprint = NormalizeThumbprint(settings.ExpectedPublisherThumbprint);
        var tempDir = Path.Combine(Path.GetTempPath(), "EXLSXS-update-verify", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            using var archive = ZipFile.OpenRead(packagePath);
            foreach (var requiredEntry in RequiredEntries)
            {
                var entry = archive.Entries.FirstOrDefault(candidate => requiredEntry.Matches(NormalizeEntryName(candidate.FullName)));
                if (entry == null)
                {
                    message = $"{requiredEntry.Label} was not found in the update package.";
                    return false;
                }

                var extractedPath = Path.Combine(tempDir, $"{Guid.NewGuid():N}-{Path.GetFileName(entry.FullName).Replace(".deploy", "", StringComparison.OrdinalIgnoreCase)}");
                entry.ExtractToFile(extractedPath);

                if (!VerifySignedFile(extractedPath, expectedThumbprint, settings.ExpectedPublisherSubject, requiredEntry.Label, out message))
                {
                    return false;
                }
            }
        }
        finally
        {
            TryDeleteDirectory(tempDir);
        }

        message = "Update publisher verified.";
        return true;
    }

    /// <summary>
    /// 単一ファイルの Authenticode 署名者が信頼設定 (thumbprint / subject) を満たすか検証する。
    /// 前提条件ブートストラッパー (vsto\setup.exe) を昇格起動する前のゲートに使う。
    /// 信頼設定が無い (HasPublisherTrustConfiguration=false) 開発ビルドでは false を返すので、
    /// 呼び出し側はその場合の扱い (起動可否) を明示的に判断すること。
    /// </summary>
    public static bool VerifyFileSigner(string filePath, UpdateSettings settings, string label, out string message)
    {
        if (!settings.HasPublisherTrustConfiguration)
        {
            message = "Expected publisher thumbprint is not configured.";
            return false;
        }

        if (!File.Exists(filePath))
        {
            message = $"{label} was not found: {filePath}";
            return false;
        }

        var expectedThumbprint = NormalizeThumbprint(settings.ExpectedPublisherThumbprint);
        return VerifySignedFile(filePath, expectedThumbprint, settings.ExpectedPublisherSubject, label, out message);
    }

    private static string ResolvePackagePath(string packageFileName)
    {
        foreach (var directory in GetPackageDirectories())
        {
            var candidate = Path.Combine(directory, packageFileName);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(GetPackageDirectories().First(), packageFileName);
    }

    private static IEnumerable<string> GetPackageDirectories()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (!string.IsNullOrWhiteSpace(localAppData))
        {
            yield return Path.Combine(localAppData, "EXLSXS", "packages");
        }

        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current != null)
        {
            yield return Path.Combine(current.FullName, "packages");
            current = current.Parent;
        }
    }

    private static bool VerifySignedFile(
        string filePath,
        string expectedThumbprint,
        string expectedSubject,
        string label,
        out string message)
    {
        X509Certificate2 certificate;
        try
        {
#pragma warning disable SYSLIB0057
            certificate = new X509Certificate2(X509Certificate.CreateFromSignedFile(filePath));
#pragma warning restore SYSLIB0057
        }
        catch (CryptographicException ex)
        {
            message = $"{label} is not Authenticode signed: {ex.Message}";
            return false;
        }

        using (certificate)
        {
            var actualThumbprint = NormalizeThumbprint(certificate.Thumbprint);
            var thumbprintMatches = string.Equals(actualThumbprint, expectedThumbprint, StringComparison.OrdinalIgnoreCase);

            if (!thumbprintMatches)
            {
                // 証明書 (Certum/SimplySign) は年次更新で thumbprint が変わる。固定 thumbprint だけに
                // 依存すると、証明書更新時に既存クライアントが新パッケージを一律 Untrusted 拒否し、
                // 自己回復不能で自動更新が恒久停止する単一障害点になる。
                // そこで thumbprint 不一致でも「subject が一致し、かつ証明書チェーンが信頼された CA に
                // チェーンする (= 失効していない正規発行証明書)」場合は受理する。subject 名は誰でも
                // 名乗れるため、チェーン検証 (信頼ルート + コード署名 EKU) で自己署名偽装を弾く。
                if (string.IsNullOrWhiteSpace(expectedSubject))
                {
                    message = $"{label} signer thumbprint mismatch. Expected {expectedThumbprint}, got {actualThumbprint}.";
                    return false;
                }

                if (certificate.Subject.IndexOf(expectedSubject, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    message = $"{label} signer thumbprint mismatch and subject mismatch. Expected subject containing '{expectedSubject}', got '{certificate.Subject}'.";
                    return false;
                }

                if (!IsCertificateChainTrusted(certificate, out var chainFailure))
                {
                    message = $"{label} signer thumbprint mismatch and certificate chain is not trusted: {chainFailure}";
                    return false;
                }

                Logger.Log($"{label} signer thumbprint rotated; accepted via subject + trusted chain. Subject='{certificate.Subject}', thumbprint={actualThumbprint}.", LogLevel.Warning);
                message = $"{label} signer verified via subject + trusted chain.";
                return true;
            }

            if (!string.IsNullOrWhiteSpace(expectedSubject)
                && certificate.Subject.IndexOf(expectedSubject, StringComparison.OrdinalIgnoreCase) < 0)
            {
                message = $"{label} signer subject mismatch. Expected subject containing '{expectedSubject}', got '{certificate.Subject}'.";
                return false;
            }
        }

        message = $"{label} signer verified.";
        return true;
    }

    /// <summary>
    /// 証明書が信頼された CA にチェーンし、コード署名用途であることを検証する。
    /// オフライン端末で失効確認ができない場合 (RevocationStatusUnknown / OfflineRevocation) は許容するが、
    /// UntrustedRoot・NotTimeValid・実際の失効などチェーン構築を妨げる状態が 1 つでもあれば信頼しない。
    /// </summary>
    private static bool IsCertificateChainTrusted(X509Certificate2 certificate, out string failureReason)
    {
        using var chain = new X509Chain();
        chain.ChainPolicy.RevocationMode = X509RevocationMode.Online;
        chain.ChainPolicy.RevocationFlag = X509RevocationFlag.ExcludeRoot;
        chain.ChainPolicy.UrlRetrievalTimeout = TimeSpan.FromSeconds(15);
        chain.ChainPolicy.ApplicationPolicy.Add(new Oid("1.3.6.1.5.5.7.3.3")); // Code Signing EKU

        chain.Build(certificate);

        const X509ChainStatusFlags acceptable = X509ChainStatusFlags.NoError
            | X509ChainStatusFlags.RevocationStatusUnknown
            | X509ChainStatusFlags.OfflineRevocation;

        var blocking = chain.ChainStatus
            .Where(status => (status.Status & ~acceptable) != X509ChainStatusFlags.NoError)
            .ToArray();

        if (blocking.Length == 0)
        {
            failureReason = "";
            return true;
        }

        failureReason = string.Join("; ", blocking.Select(status => $"{status.Status}: {status.StatusInformation.Trim()}"));
        return false;
    }

    internal static string NormalizeEntryName(string name)
    {
        return "/" + name.Replace('\\', '/').TrimStart('/');
    }

    internal static string NormalizeThumbprint(string thumbprint)
    {
        return new string(thumbprint.Where(Uri.IsHexDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static void TryDeleteDirectory(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
        catch (IOException ex)
        {
            Logger.LogException($"Temporary update verification directory could not be removed: {directory}", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            Logger.LogException($"Temporary update verification directory could not be removed: {directory}", ex);
        }
    }
}
