using Velopack;

namespace EXLSXS.Host;

internal static class Program
{
    private const string UpdateCheckArg = "--update-check";

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            VelopackApp.Build()
                .OnAfterInstallFastCallback(_ => InstalledAppMaintenance.Register(AddInRegistrationMode.ForceEnabled))
                .OnAfterUpdateFastCallback(_ => InstalledAppMaintenance.Register(AddInRegistrationMode.PreserveLoadBehavior))
                .OnBeforeUninstallFastCallback(_ => InstalledAppMaintenance.Unregister())
                .Run();

            if (HasArg(args, UpdateCheckArg))
            {
                RunSilentUpdateCheck();
                return 0;
            }

            if (HasArg(args, "--register"))
            {
                InstalledAppMaintenance.Register(AddInRegistrationMode.ForceEnabled);
                return 0;
            }

            if (HasArg(args, "--unregister"))
            {
                InstalledAppMaintenance.Unregister();
                return 0;
            }

            InstalledAppMaintenance.Register(AddInRegistrationMode.PreserveLoadBehavior);
            Logger.Log("Registration check completed.");
            return 0;
        }
        catch (Exception ex)
        {
            Logger.LogException("Host failed.", ex);
            return 1;
        }
        finally
        {
            Logger.Dispose();
        }
    }

    private static bool HasArg(IEnumerable<string> args, string value)
    {
        return args.Any(arg => string.Equals(arg, value, StringComparison.OrdinalIgnoreCase));
    }

    private static void RunSilentUpdateCheck()
    {
        try
        {
            Logger.Log("Silent update check started.");
            InstalledAppMaintenance.Register(AddInRegistrationMode.PreserveLoadBehavior);

            var result = UpdateChecker.CheckAndDownloadAsync().GetAwaiter().GetResult();
            if (result.Result == UpdateChecker.UpdateResult.Downloaded && result.Info != null && result.Manager != null)
            {
                Logger.Log("Update downloaded. Applying update and exiting.");
                result.Manager.ApplyUpdatesAndExit(result.Info);
                return;
            }

            Logger.Log($"Silent update check completed: {result.Result}. {result.Message}");
        }
        catch (Exception ex)
        {
            Logger.LogException("Silent update check failed.", ex);
        }
    }
}
