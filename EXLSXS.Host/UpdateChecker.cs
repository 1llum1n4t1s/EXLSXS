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

            var checkTask = updateManager.CheckForUpdatesAsync();
            var timeoutTask = Task.Delay(CheckTimeoutMs, cancellationToken);
            var completedTask = await Task.WhenAny(checkTask, timeoutTask);
            if (completedTask == timeoutTask)
            {
                Logger.Log("Update check timed out.", LogLevel.Warning);
                return new CheckResult(UpdateResult.Error, null, null, "Update check timed out.");
            }

            var updateInfo = await checkTask;
            if (updateInfo == null)
            {
                Logger.Log("No update is available.");
                return new CheckResult(UpdateResult.NoUpdate, null, null, "No update is available.");
            }

            Logger.Log("Update found. Downloading update.");
            statusProgress?.Report("Downloading update.");

            using var downloadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            downloadCts.CancelAfter(DownloadTimeout);
            await updateManager.DownloadUpdatesAsync(updateInfo, null, downloadCts.Token);

            Logger.Log("Update download completed.");
            statusProgress?.Report("Verifying update publisher.");
            if (!UpdatePackageTrustVerifier.Verify(updateInfo, settings, out var trustMessage))
            {
                Logger.Log($"Update publisher verification failed: {trustMessage}", LogLevel.Warning);
                return new CheckResult(UpdateResult.Untrusted, null, null, trustMessage);
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

    private static UpdateManager CreateUpdateManager(UpdateSettings settings)
    {
        var options = new UpdateOptions();
        if (!string.IsNullOrWhiteSpace(settings.Channel))
        {
            options.ExplicitChannel = settings.Channel;
        }

        if (ShouldUseGithubSource(settings))
        {
            var source = new GithubSource(settings.Source, settings.AccessToken, settings.Prerelease);
            return new UpdateManager(source, options);
        }

        return new UpdateManager(settings.Source, options);
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
