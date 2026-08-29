using System.Runtime.InteropServices;
using Velopack;

namespace EXLSXS.Host;

internal static class Program
{
    private const string UpdateCheckArg = "--update-check";
    private const string AppUserModelId = "velopack.EXLSXS";

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            TrySetCurrentProcessAppUserModelId();

            VelopackApp.Build()
                .SetAutoApplyOnStartup(false)
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

    private static void TrySetCurrentProcessAppUserModelId()
    {
        try { _ = SetCurrentProcessExplicitAppUserModelID(AppUserModelId); }
        catch { /* シェル連携の失敗だけでホスト処理を止めない */ }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SetCurrentProcessExplicitAppUserModelID(string appId);

    private static void SafeRegister(AddInRegistrationMode mode, bool allowPrerequisiteInstall = false)
    {
        try
        {
            InstalledAppMaintenance.Register(mode, allowPrerequisiteInstall);
        }
        catch (Exception ex)
        {
            // 登録失敗だけで Velopack のインストールや更新処理を止めない。
            // --update-check では更新を先に届け、更新後 callback で現在の BaseDirectory を再登録できるようにする。
            Logger.LogException($"Add-in registration failed; processing will continue ({mode}).", ex);
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
            SafeRegister(AddInRegistrationMode.PreserveLoadBehavior);

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
