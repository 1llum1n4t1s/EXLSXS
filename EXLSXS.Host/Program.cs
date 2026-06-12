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
                .OnAfterInstallFastCallback(_ => SafeRegister(AddInRegistrationMode.ForceEnabled, allowPrerequisiteInstall: true))
                .OnAfterUpdateFastCallback(_ => SafeRegister(AddInRegistrationMode.PreserveLoadBehavior))
                .OnBeforeUninstallFastCallback(_ => InstalledAppMaintenance.Unregister())
                .Run();

            if (HasArg(args, UpdateCheckArg))
            {
                RunSilentUpdateCheck();
                return 0;
            }

            if (HasArg(args, "--register"))
            {
                InstalledAppMaintenance.Register(AddInRegistrationMode.ForceEnabled, allowPrerequisiteInstall: true);
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

    private static void SafeRegister(AddInRegistrationMode mode, bool allowPrerequisiteInstall = false)
    {
        try
        {
            InstalledAppMaintenance.Register(mode, allowPrerequisiteInstall);
        }
        catch (Exception ex)
        {
            // インストール/更新の fast callback 内で登録が失敗しても Velopack のインストール状態は壊さない。
            // 次回ログオン時の --update-check で Register が走り、現在の BaseDirectory の manifest パスで
            // 自己修復される (vstolocal 絶対パスと Velopack のフォルダ入れ替えの不整合に対する保険)。
            Logger.LogException($"Add-in registration failed during Velopack callback ({mode}).", ex);
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
