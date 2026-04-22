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
            if (!string.Equals(actualThumbprint, expectedThumbprint, StringComparison.OrdinalIgnoreCase))
            {
                message = $"{label} signer thumbprint mismatch. Expected {expectedThumbprint}, got {actualThumbprint}.";
                return false;
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

    private static string NormalizeEntryName(string name)
    {
        return "/" + name.Replace('\\', '/').TrimStart('/');
    }

    private static string NormalizeThumbprint(string thumbprint)
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
