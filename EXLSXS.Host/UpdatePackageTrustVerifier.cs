using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using Velopack;

namespace EXLSXS.Host;

internal static class UpdatePackageTrustVerifier
{
    private const string MicrosoftPublisherOrganization = "Microsoft Corporation";
    private const string UpdaterEntrySuffix = "/Squirrel.exe";

    private static readonly HashSet<string> AllowedPackagedUpdateHosts =
    [
        "exlsxs.kagayoi.com",
        "exlsxs.nephilim.jp"
    ];

    internal sealed record VerifiedUpdatePackage(string PackagePath, byte[] UpdaterBytes);

    private sealed record RequiredSignedEntry(string Label, Func<string, bool> Matches, bool IsUpdater = false);

    private static readonly RequiredSignedEntry[] RequiredEntries =
    [
        new("EXLSXS host", name => name.EndsWith("/EXLSXS.Host.exe", StringComparison.OrdinalIgnoreCase)),
        new("EXLSXS host assembly", name => name.EndsWith("/EXLSXS.Host.dll", StringComparison.OrdinalIgnoreCase)),
        new("EXLSXS VSTO assembly", name => name.EndsWith("/EXLSXS.dll.deploy", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith("/EXLSXS.dll", StringComparison.OrdinalIgnoreCase)),
        new("Velopack updater", name => name.EndsWith(UpdaterEntrySuffix, StringComparison.OrdinalIgnoreCase), IsUpdater: true)
    ];

    public static bool Verify(
        UpdateInfo updateInfo,
        UpdateSettings settings,
        string packagesDirectory,
        out VerifiedUpdatePackage? verifiedPackage,
        out string message)
    {
        verifiedPackage = null;
        if (!settings.HasPublisherTrustConfiguration)
        {
            message = "Expected publisher thumbprint is not configured.";
            return false;
        }

        var packageFileName = updateInfo.TargetFullRelease.FileName;
        if (!TryResolvePackagePath(packagesDirectory, packageFileName, out var packagePath, out message))
        {
            return false;
        }

        if (!File.Exists(packagePath))
        {
            message = $"Downloaded update package was not found: {packageFileName}";
            return false;
        }

        if (!VerifyPackageFile(
                packagePath,
                settings,
                deleteUntrustedPackage: true,
                out var updaterBytes,
                out message))
        {
            return false;
        }

        verifiedPackage = new VerifiedUpdatePackage(packagePath, updaterBytes!);
        return true;
    }

    internal static bool VerifyPackageFile(
        string packagePath,
        UpdateSettings settings,
        bool deleteUntrustedPackage,
        out string message)
    {
        return VerifyPackageFile(packagePath, settings, deleteUntrustedPackage, out _, out message);
    }

    private static bool VerifyPackageFile(
        string packagePath,
        UpdateSettings settings,
        bool deleteUntrustedPackage,
        out byte[]? verifiedUpdaterBytes,
        out string message)
    {
        verifiedUpdaterBytes = null;
        if (!settings.HasPublisherTrustConfiguration)
        {
            message = "Expected publisher thumbprint is not configured.";
            return false;
        }

        if (!File.Exists(packagePath))
        {
            message = $"Update package was not found: {packagePath}";
            return false;
        }

        var expectedThumbprint = NormalizeThumbprint(settings.ExpectedPublisherThumbprint);
        var verifiedPublisherThumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var matchedRequiredEntries = new HashSet<RequiredSignedEntry>();
        var updaterEntryCount = 0;
        var tempDir = Path.Combine(Path.GetTempPath(), "EXLSXS-update-verify", Guid.NewGuid().ToString("N"));
        var packageTrusted = false;

        try
        {
            Directory.CreateDirectory(tempDir);
            using var archive = ZipFile.OpenRead(packagePath);

            foreach (var entry in archive.Entries)
            {
                var normalizedName = NormalizeEntryName(entry.FullName);
                if (!IsPortableExecutableEntry(normalizedName))
                {
                    continue;
                }

                var requiredEntry = RequiredEntries.FirstOrDefault(candidate => candidate.Matches(normalizedName));
                var label = requiredEntry?.Label ?? $"Package entry '{normalizedName}'";
                var extractedPath = Path.Combine(
                    tempDir,
                    $"{Guid.NewGuid():N}-{Path.GetFileName(entry.FullName).Replace(".deploy", "", StringComparison.OrdinalIgnoreCase)}");

                entry.ExtractToFile(extractedPath);
                try
                {
                    if (requiredEntry != null)
                    {
                        if (!VerifySignedFile(
                                extractedPath,
                                expectedThumbprint,
                                settings.ExpectedPublisherSubject,
                                label,
                                out var actualThumbprint,
                                out message))
                        {
                            return false;
                        }

                        matchedRequiredEntries.Add(requiredEntry);
                        verifiedPublisherThumbprints.Add(actualThumbprint);
                        if (requiredEntry.IsUpdater)
                        {
                            updaterEntryCount++;
                            verifiedUpdaterBytes = File.ReadAllBytes(extractedPath);
                        }
                    }
                    else if (!VerifyAllowedPackageFile(
                                 extractedPath,
                                 expectedThumbprint,
                                 settings.ExpectedPublisherSubject,
                                 label,
                                 out message))
                    {
                        return false;
                    }
                }
                finally
                {
                    TryDeleteFile(extractedPath, "Temporary update verification file");
                }
            }

            var missingEntry = RequiredEntries.FirstOrDefault(entry => !matchedRequiredEntries.Contains(entry));
            if (missingEntry != null)
            {
                message = $"{missingEntry.Label} was not found in the update package.";
                return false;
            }

            if (updaterEntryCount != 1 || verifiedUpdaterBytes == null)
            {
                message = $"The update package must contain exactly one publisher-signed Squirrel.exe entry; found {updaterEntryCount}.";
                return false;
            }

            if (!VerifyPackagedTrustConfiguration(archive, settings, verifiedPublisherThumbprints, out message))
            {
                return false;
            }

            packageTrusted = true;
            message = "Update package publisher and executable contents verified.";
            return true;
        }
        finally
        {
            TryDeleteDirectory(tempDir);
            if (!packageTrusted && deleteUntrustedPackage)
            {
                TryDeleteFile(packagePath, "Untrusted update package");
            }
        }
    }

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
        return VerifySignedFile(filePath, expectedThumbprint, settings.ExpectedPublisherSubject, label, out _, out message);
    }

    private static bool VerifyAllowedPackageFile(
        string filePath,
        string expectedThumbprint,
        string expectedSubject,
        string label,
        out string message)
    {
        if (!TryGetVerifiedSignerCertificate(filePath, label, out var certificate, out message))
        {
            return false;
        }

        using (certificate)
        {
            if (CertificateMatchesExpectedPublisher(certificate, expectedThumbprint, expectedSubject, out _)
                || IsMicrosoftPlatformCertificate(certificate))
            {
                message = $"{label} signer verified.";
                return true;
            }

            message = $"{label} is signed by an unexpected publisher: '{certificate.Subject}'.";
            return false;
        }
    }

    private static bool VerifySignedFile(
        string filePath,
        string expectedThumbprint,
        string expectedSubject,
        string label,
        out string actualThumbprint,
        out string message)
    {
        actualThumbprint = "";
        if (!TryGetVerifiedSignerCertificate(filePath, label, out var certificate, out message))
        {
            return false;
        }

        using (certificate)
        {
            actualThumbprint = NormalizeThumbprint(certificate.Thumbprint);
            if (!CertificateMatchesExpectedPublisher(certificate, expectedThumbprint, expectedSubject, out var rotated))
            {
                message = $"{label} signer does not match the configured publisher. "
                    + $"Expected thumbprint {expectedThumbprint} or subject '{expectedSubject}', got '{certificate.Subject}' ({actualThumbprint}).";
                return false;
            }

            if (rotated)
            {
                Logger.Log(
                    $"{label} signer thumbprint rotated; accepted after Authenticode trust and exact subject verification. "
                    + $"Subject='{certificate.Subject}', thumbprint={actualThumbprint}.",
                    LogLevel.Warning);
            }
        }

        message = $"{label} signer verified.";
        return true;
    }

    private static bool TryGetVerifiedSignerCertificate(
        string filePath,
        string label,
        out X509Certificate2 certificate,
        out string message)
    {
        certificate = null!;
        if (!AuthenticodeTrustVerifier.TryVerifyEmbeddedSignature(
                filePath,
                out var verifiedSignerCertificate,
                out var trustFailure))
        {
            message = $"{label} Authenticode signature is not trusted: {trustFailure}";
            return false;
        }

        certificate = verifiedSignerCertificate!;
        message = "";
        return true;
    }

    private static bool CertificateMatchesExpectedPublisher(
        X509Certificate2 certificate,
        string expectedThumbprint,
        string expectedSubject,
        out bool thumbprintRotated)
    {
        var actualThumbprint = NormalizeThumbprint(certificate.Thumbprint);
        var thumbprintMatches = string.Equals(actualThumbprint, expectedThumbprint, StringComparison.OrdinalIgnoreCase);
        var actualSubject = certificate.GetNameInfo(X509NameType.SimpleName, forIssuer: false);
        var subjectMatches = string.IsNullOrWhiteSpace(expectedSubject)
            || string.Equals(actualSubject, expectedSubject.Trim(), StringComparison.OrdinalIgnoreCase);

        thumbprintRotated = !thumbprintMatches && subjectMatches && !string.IsNullOrWhiteSpace(expectedSubject);
        return subjectMatches && (thumbprintMatches || thumbprintRotated);
    }

    private static bool VerifyPackagedTrustConfiguration(
        ZipArchive archive,
        UpdateSettings currentSettings,
        IReadOnlySet<string> verifiedPublisherThumbprints,
        out string message)
    {
        var settingsEntries = archive.Entries.Where(entry =>
            string.Equals(
                NormalizeEntryName(entry.FullName),
                "/lib/app/appsettings.json",
                StringComparison.OrdinalIgnoreCase)).ToArray();

        if (settingsEntries.Length != 1)
        {
            message = $"The update package must contain exactly one lib/app/appsettings.json entry; found {settingsEntries.Length}.";
            return false;
        }

        try
        {
            using var stream = settingsEntries[0].Open();
            using var document = JsonDocument.Parse(stream);
            if (!document.RootElement.TryGetProperty("Update", out var update)
                || !update.TryGetProperty("Source", out var sourceElement)
                || sourceElement.ValueKind != JsonValueKind.String
                || !update.TryGetProperty("SourceKind", out var sourceKindElement)
                || sourceKindElement.ValueKind != JsonValueKind.String
                || !update.TryGetProperty("Channel", out var channelElement)
                || channelElement.ValueKind != JsonValueKind.String
                || !update.TryGetProperty("AccessToken", out var accessTokenElement)
                || accessTokenElement.ValueKind != JsonValueKind.String
                || !update.TryGetProperty("Prerelease", out var prereleaseElement)
                || prereleaseElement.ValueKind is not (JsonValueKind.True or JsonValueKind.False)
                || !update.TryGetProperty("ExpectedPublisherThumbprint", out var thumbprintElement)
                || thumbprintElement.ValueKind != JsonValueKind.String
                || !update.TryGetProperty("ExpectedPublisherSubject", out var subjectElement)
                || subjectElement.ValueKind != JsonValueKind.String)
            {
                message = "Packaged publisher trust configuration is incomplete.";
                return false;
            }

            if (!IsPackagedUpdateConfigurationAllowed(
                    sourceElement.GetString() ?? "",
                    sourceKindElement.GetString() ?? "",
                    channelElement.GetString() ?? "",
                    accessTokenElement.GetString() ?? "",
                    prereleaseElement.GetBoolean(),
                    out message))
            {
                return false;
            }

            return IsPackagedTrustConfigurationAllowed(
                currentSettings.ExpectedPublisherThumbprint,
                currentSettings.ExpectedPublisherSubject,
                thumbprintElement.GetString() ?? "",
                subjectElement.GetString() ?? "",
                verifiedPublisherThumbprints,
                out message);
        }
        catch (JsonException ex)
        {
            message = $"Packaged publisher trust configuration is invalid JSON: {ex.Message}";
            return false;
        }
    }

    internal static bool IsPackagedUpdateConfigurationAllowed(
        string source,
        string sourceKind,
        string channel,
        string accessToken,
        bool prerelease,
        out string message)
    {
        if (!Uri.TryCreate(source, UriKind.Absolute, out var sourceUri)
            || !string.Equals(sourceUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
            || !AllowedPackagedUpdateHosts.Contains(sourceUri.Host)
            || !string.IsNullOrEmpty(sourceUri.UserInfo)
            || !string.IsNullOrEmpty(sourceUri.Query)
            || !string.IsNullOrEmpty(sourceUri.Fragment)
            || sourceUri.AbsolutePath != "/")
        {
            message = $"Packaged update source is not an allowed EXLSXS HTTPS endpoint: '{source}'.";
            return false;
        }

        if (!string.Equals(sourceKind, UpdateSourceKind.Simple.ToString(), StringComparison.OrdinalIgnoreCase))
        {
            message = $"Packaged update source kind must be Simple; found '{sourceKind}'.";
            return false;
        }

        if (!string.Equals(channel, "win", StringComparison.OrdinalIgnoreCase))
        {
            message = $"Packaged update channel must be 'win'; found '{channel}'.";
            return false;
        }

        if (!string.IsNullOrEmpty(accessToken))
        {
            message = "Packaged update access token must be empty.";
            return false;
        }

        if (prerelease)
        {
            message = "Packaged prerelease updates are not allowed.";
            return false;
        }

        message = "Packaged update endpoint configuration verified.";
        return true;
    }

    internal static bool IsPackagedTrustConfigurationAllowed(
        string currentThumbprint,
        string currentSubject,
        string packagedThumbprint,
        string packagedSubject,
        IReadOnlySet<string> verifiedPublisherThumbprints,
        out string message)
    {
        var normalizedCurrent = NormalizeThumbprint(currentThumbprint);
        var normalizedPackaged = NormalizeThumbprint(packagedThumbprint);
        if (string.IsNullOrWhiteSpace(normalizedPackaged))
        {
            message = "Packaged publisher thumbprint is empty.";
            return false;
        }

        if (!string.Equals(currentSubject.Trim(), packagedSubject.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            message = $"Packaged publisher subject changed from '{currentSubject}' to '{packagedSubject}'.";
            return false;
        }

        if (!string.Equals(normalizedCurrent, normalizedPackaged, StringComparison.OrdinalIgnoreCase)
            && !verifiedPublisherThumbprints.Contains(normalizedPackaged))
        {
            message = "Packaged publisher thumbprint does not match the current trust anchor or a verified core signer.";
            return false;
        }

        message = "Packaged publisher trust configuration verified.";
        return true;
    }

    internal static bool IsPortableExecutableEntry(string entryName)
    {
        var normalized = NormalizeEntryName(entryName);
        return normalized.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            || normalized.EndsWith(".dll", StringComparison.OrdinalIgnoreCase)
            || normalized.EndsWith(".exe.deploy", StringComparison.OrdinalIgnoreCase)
            || normalized.EndsWith(".dll.deploy", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMicrosoftPlatformCertificate(X509Certificate2 certificate)
    {
        var decodedSubject = certificate.SubjectName.Decode(X500DistinguishedNameFlags.UseNewLines);
        return decodedSubject
            .Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Any(part => string.Equals(part, $"O={MicrosoftPublisherOrganization}", StringComparison.OrdinalIgnoreCase));
    }

    internal static bool TryResolvePackagePath(
        string packagesDirectory,
        string packageFileName,
        out string packagePath,
        out string message)
    {
        packagePath = "";
        if (string.IsNullOrWhiteSpace(packagesDirectory)
            || string.IsNullOrWhiteSpace(packageFileName)
            || Path.IsPathRooted(packageFileName)
            || packageFileName.IndexOfAny([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar]) >= 0
            || packageFileName.Contains("..", StringComparison.Ordinal)
            || !string.Equals(Path.GetExtension(packageFileName), ".nupkg", StringComparison.OrdinalIgnoreCase))
        {
            message = $"Update package file name is unsafe: '{packageFileName}'.";
            return false;
        }

        try
        {
            var fullPackagesDirectory = Path.GetFullPath(packagesDirectory);
            packagePath = Path.GetFullPath(Path.Combine(fullPackagesDirectory, packageFileName));
            if (!string.Equals(Path.GetDirectoryName(packagePath), fullPackagesDirectory, StringComparison.OrdinalIgnoreCase))
            {
                message = $"Update package escaped the Velopack packages directory: '{packageFileName}'.";
                packagePath = "";
                return false;
            }
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            message = $"Update package path is invalid: {ex.Message}";
            packagePath = "";
            return false;
        }

        message = "Update package path resolved inside the Velopack packages directory.";
        return true;
    }

    internal static bool TryInstallVerifiedUpdater(
        string updaterPath,
        byte[] updaterBytes,
        UpdateSettings settings,
        out string message)
    {
        if (updaterBytes.Length == 0)
        {
            message = "Verified updater payload is empty.";
            return false;
        }

        var updaterDirectory = Path.GetDirectoryName(updaterPath);
        if (string.IsNullOrWhiteSpace(updaterDirectory) || !Directory.Exists(updaterDirectory))
        {
            message = $"Velopack updater directory was not found: {updaterDirectory}";
            return false;
        }

        var temporaryPath = Path.Combine(updaterDirectory, $".{Path.GetFileName(updaterPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllBytes(temporaryPath, updaterBytes);
            if (!VerifyFileSigner(temporaryPath, settings, "Velopack updater", out message))
            {
                return false;
            }

            File.Move(temporaryPath, updaterPath, overwrite: true);
            message = "Verified Velopack updater installed.";
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            message = $"Verified Velopack updater could not be installed: {ex.Message}";
            return false;
        }
        finally
        {
            TryDeleteFile(temporaryPath, "Temporary verified updater");
        }
    }

    internal static string NormalizeEntryName(string name)
    {
        return "/" + name.Replace('\\', '/').TrimStart('/');
    }

    internal static string NormalizeThumbprint(string thumbprint)
    {
        return new string(thumbprint.Where(Uri.IsHexDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static void TryDeleteFile(string filePath, string label)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch (IOException ex)
        {
            Logger.LogException($"{label} could not be removed: {filePath}", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            Logger.LogException($"{label} could not be removed: {filePath}", ex);
        }
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
