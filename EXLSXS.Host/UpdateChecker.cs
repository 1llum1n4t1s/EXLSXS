using Velopack;
using Velopack.Sources;

namespace EXLSXS.Host;

internal static class UpdateChecker
{
    private const int CheckTimeoutMs = 10000;
    private static readonly TimeSpan DownloadTimeout = TimeSpan.FromMinutes(10);

    public enum UpdateResult
    {
        NoUpdate,
        Downloaded,
        Untrusted,
        NotInstalled,
        NotConfigured,
        Error
    }

    public sealed record CheckResult(UpdateResult Result, UpdateInfo? Info, UpdateManager? Manager, string Message);

    private sealed class LocatedUpdateManager : UpdateManager
    {
        public LocatedUpdateManager(string source, UpdateOptions options)
            : base(source, options)
        {
        }

        public LocatedUpdateManager(IUpdateSource source, UpdateOptions options)
            : base(source, options)
        {
        }

        public string PackagesDirectory => Locator.PackagesDir
            ?? throw new InvalidOperationException("Velopack packages directory is not available.");

        public string UpdaterPath => Locator.UpdateExePath
            ?? throw new InvalidOperationException("Velopack updater path is not available.");
    }

    public static async Task<CheckResult> CheckAndDownloadAsync(
        IProgress<string>? statusProgress = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var settings = UpdateSettings.Load();
            if (!settings.IsConfigured)
            {
                Logger.Log("Update source is not configured. Skipping update check.");
                return new CheckResult(UpdateResult.NotConfigured, null, null, "Update source is not configured.");
            }

            var updateManager = CreateUpdateManager(settings);
            if (!updateManager.IsInstalled)
            {
                Logger.Log("Application is not installed by Velopack. Skipping update check.");
                return new CheckResult(UpdateResult.NotInstalled, null, null, "Application is not installed by Velopack.");
            }

            Logger.Log($"Checking updates from '{settings.Source}' on channel '{settings.Channel}'.");
            statusProgress?.Report("Checking for updates.");

            UpdateInfo? updateInfo;
            try
            {
                updateInfo = await updateManager.CheckForUpdatesAsync()
                    .WaitAsync(TimeSpan.FromMilliseconds(CheckTimeoutMs), cancellationToken);
            }
            catch (TimeoutException)
            {
                Logger.Log("Update check timed out.", LogLevel.Warning);
                return new CheckResult(UpdateResult.Error, null, null, "Update check timed out.");
            }

            if (updateInfo == null)
            {
                Logger.Log("No update is available.");
                return new CheckResult(UpdateResult.NoUpdate, null, null, "No update is available.");
            }

            Logger.Log("Update found. Downloading update.");
            statusProgress?.Report("Downloading update.");

            using var downloadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            downloadCts.CancelAfter(DownloadTimeout);
            UpdatePackageTrustVerifier.VerifiedUpdatePackage? verifiedPackage;
            var updaterPath = updateManager.UpdaterPath;
            using (new FileStream(updaterPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // Velopack 1.2.0 は DownloadUpdatesAsync の finally で、パッケージ内 Squirrel.exe を
                // 検証前に Update.exe へ展開する。書き込み共有を拒否したハンドルを保持し、
                // 未信頼パッケージが updater を先に差し替える経路を閉じる。
                await updateManager.DownloadUpdatesAsync(updateInfo, null, downloadCts.Token);

                Logger.Log("Update download completed.");
                statusProgress?.Report("Verifying update publisher.");
                if (!UpdatePackageTrustVerifier.Verify(
                        updateInfo,
                        settings,
                        updateManager.PackagesDirectory,
                        out verifiedPackage,
                        out var trustMessage))
                {
                    Logger.Log($"Update publisher verification failed: {trustMessage}", LogLevel.Warning);
                    return new CheckResult(UpdateResult.Untrusted, null, null, trustMessage);
                }
            }

            if (!UpdatePackageTrustVerifier.TryInstallVerifiedUpdater(
                    updaterPath,
                    verifiedPackage!.UpdaterBytes,
                    settings,
                    out var updaterMessage))
            {
                Logger.Log($"Verified updater installation failed: {updaterMessage}", LogLevel.Warning);
                return new CheckResult(UpdateResult.Error, null, null, updaterMessage);
            }

            return new CheckResult(UpdateResult.Downloaded, updateInfo, updateManager, "Update downloaded.");
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            Logger.Log("Update download timed out.", LogLevel.Warning);
            return new CheckResult(UpdateResult.Error, null, null, "Update download timed out.");
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            Logger.LogException("Update check failed.", ex);
            return new CheckResult(UpdateResult.Error, null, null, "Update check failed.");
        }
    }

    private static LocatedUpdateManager CreateUpdateManager(UpdateSettings settings)
    {
        var options = new UpdateOptions();
        if (!string.IsNullOrWhiteSpace(settings.Channel))
        {
            options.ExplicitChannel = settings.Channel;
        }

        if (ShouldUseGithubSource(settings))
        {
            var source = new GithubSource(settings.Source, settings.AccessToken, settings.Prerelease);
            return new LocatedUpdateManager(source, options);
        }

        return new LocatedUpdateManager(settings.Source, options);
    }

    internal static bool ShouldUseGithubSource(UpdateSettings settings)
    {
        if (settings.SourceKind == UpdateSourceKind.Github)
        {
            return true;
        }

        if (settings.SourceKind == UpdateSourceKind.Simple)
        {
            return false;
        }

        return Uri.TryCreate(settings.Source, UriKind.Absolute, out var uri)
            && string.Equals(uri.Host, "github.com", StringComparison.OrdinalIgnoreCase);
    }
}
