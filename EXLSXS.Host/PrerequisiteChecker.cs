using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.Versioning;

namespace EXLSXS.Host;

[SupportedOSPlatform("windows")]
internal static class PrerequisiteChecker
{
    private const int NetFramework481Release = 533320;
    private const string NetFrameworkFullKey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full";
    // VSTO ランタイムは導入経路により "v4" (Office / VS 同梱) か "v4R" (再頒布パッケージ単体) のいずれかに登録される。
    // どちらか一方でも Install=1 なら導入済みとみなす。片方しか見ないと、入っていても「未導入」と誤判定する。
    private static readonly string[] VstoRuntimeKeys =
    [
        @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
        @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
    ];

    public static void EnsureReady(bool allowPrerequisiteInstall = false)
    {
        var missing = GetMissingPrerequisites();
        if (missing.Count == 0)
        {
            return;
        }

        var bootstrapperPath = ResolveBootstrapperPath();

        // インストール文脈でのみ、同梱の前提インストーラー (vsto\setup.exe) を実走させて自動修復する。
        // サイレント更新 (Windows 起動時) などでは false のまま例外を投げ、UAC を不意に出さない。
        if (allowPrerequisiteInstall && File.Exists(bootstrapperPath))
        {
            Logger.Log($"Prerequisites missing ({string.Join(", ", missing)}). Running bundled bootstrapper: {bootstrapperPath}");
            RunBootstrapper(bootstrapperPath);

            missing = GetMissingPrerequisites();
            if (missing.Count == 0)
            {
                Logger.Log("Prerequisites satisfied after running the bundled bootstrapper.");
                return;
            }

            Logger.Log($"Prerequisites still missing after running the bootstrapper: {string.Join(", ", missing)}. A reboot may be required.", LogLevel.Warning);
        }

        var hint = File.Exists(bootstrapperPath)
            ? $" Run the bundled bootstrapper first: {bootstrapperPath}"
            : " Install the missing prerequisites before registering the add-in.";

        throw new InvalidOperationException($"Required prerequisites are missing: {string.Join(", ", missing)}.{hint}");
    }

    private static List<string> GetMissingPrerequisites()
    {
        var missing = new List<string>();
        if (!IsNetFramework481OrLaterInstalled())
        {
            missing.Add(".NET Framework 4.8.1");
        }

        if (!IsVstoRuntimeInstalled())
        {
            missing.Add("Visual Studio Tools for Office Runtime");
        }

        return missing;
    }

    private static void RunBootstrapper(string bootstrapperPath)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = bootstrapperPath,
                // UseShellExecute は setup.exe 内の前提インストーラーが UAC で昇格できるようにするため必須。
                UseShellExecute = true,
                WorkingDirectory = Path.GetDirectoryName(bootstrapperPath) ?? AppContext.BaseDirectory
            };

            using var process = Process.Start(startInfo);
            if (process == null)
            {
                Logger.Log("Bootstrapper process could not be started.", LogLevel.Warning);
                return;
            }

            process.WaitForExit();
            Logger.Log($"Bootstrapper exited with code {process.ExitCode}.", LogLevel.Debug);
        }
        catch (Exception ex)
        {
            // UAC キャンセル等で失敗しても、呼び出し側が前提を再チェックして最終判断する。
            Logger.LogException("Running the bundled bootstrapper failed.", ex);
        }
    }

    private static bool IsNetFramework481OrLaterInstalled()
    {
        using var key = Registry.LocalMachine.OpenSubKey(NetFrameworkFullKey);
        return TryReadInt(key?.GetValue("Release"), out var release) && release >= NetFramework481Release;
    }

    private static bool IsVstoRuntimeInstalled()
    {
        return IsVstoRuntimeInstalled(RegistryView.Registry64)
            || IsVstoRuntimeInstalled(RegistryView.Registry32);
    }

    private static bool IsVstoRuntimeInstalled(RegistryView view)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            foreach (var subKey in VstoRuntimeKeys)
            {
                using var key = baseKey.OpenSubKey(subKey);
                if (TryReadInt(key?.GetValue("Install"), out var installed) && installed == 1)
                {
                    return true;
                }
            }

            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static bool TryReadInt(object? value, out int result)
    {
        switch (value)
        {
            case int intValue:
                result = intValue;
                return true;
            case string stringValue when int.TryParse(stringValue, out var parsed):
                result = parsed;
                return true;
            default:
                result = 0;
                return false;
        }
    }

    private static string ResolveBootstrapperPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "vsto", "setup.exe");
    }
}
