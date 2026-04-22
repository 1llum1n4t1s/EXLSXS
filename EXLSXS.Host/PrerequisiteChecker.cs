using Microsoft.Win32;
using System.Runtime.Versioning;

namespace EXLSXS.Host;

[SupportedOSPlatform("windows")]
internal static class PrerequisiteChecker
{
    private const int NetFramework481Release = 533320;
    private const string NetFrameworkFullKey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full";
    private const string VstoRuntimeKey = @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R";

    public static void EnsureReady()
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

        if (missing.Count == 0)
        {
            return;
        }

        var bootstrapperPath = ResolveBootstrapperPath();
        var hint = File.Exists(bootstrapperPath)
            ? $" Run the bundled bootstrapper first: {bootstrapperPath}"
            : " Install the missing prerequisites before registering the add-in.";

        throw new InvalidOperationException($"Required prerequisites are missing: {string.Join(", ", missing)}.{hint}");
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
            using var key = baseKey.OpenSubKey(VstoRuntimeKey);
            return TryReadInt(key?.GetValue("Install"), out var installed) && installed == 1;
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
