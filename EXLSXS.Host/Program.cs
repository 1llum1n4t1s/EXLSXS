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

            var handledFirstRun = false;
            VelopackApp.Build()
                .SetAutoApplyOnStartup(false)
                // FastCallback は短時間で戻る必要があるため、前提インストーラーをここでは待たない。
                // first run が起動しない／再起動が必要な場合に備えて、先に一度だけ再登録を予約する。
                .OnAfterInstallFastCallback(_ => SafeScheduleRegistrationRetry())
                .OnFirstRun(_ =>
                {
                    handledFirstRun = true;
                    SafeRegister(AddInRegistrationMode.ForceEnabled, allowPrerequisiteInstall: true);
                })
                .OnAfterUpdateFastCallback(_ => SafeRegister(AddInRegistrationMode.PreserveLoadBehavior))
                .OnBeforeUninstallFastCallback(_ => InstalledAppMaintenance.Unregister())
                .Run();

            if (handledFirstRun)
            {
                return 0;
            }

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

    private static void SafeScheduleRegistrationRetry()
    {
        try
        {
            StartupRegistration.RegisterRegistrationRetry();
        }
        catch (Exception ex)
        {
            // 回復経路の登録失敗だけで Velopack のインストール自体を止めない。
            Logger.LogException("Registration retry could not be scheduled; installation will continue.", ex);
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
            // first run の失敗時は OnAfterInstallFastCallback が登録した RunOnce を残す。
            // update callback では更新を先に届け、次回起動で現在の BaseDirectory を再登録できるようにする。
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
