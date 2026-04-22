using Microsoft.Win32;
using System.Runtime.Versioning;

namespace EXLSXS.Host;

[SupportedOSPlatform("windows")]
internal static class StartupRegistration
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string EntryName = "EXLSXS";

    public static void Register()
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exePath))
        {
            throw new InvalidOperationException("Startup registration failed because process path is empty.");
        }

        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath);
        if (key == null)
        {
            throw new InvalidOperationException("Startup registration key could not be opened.");
        }

        var value = $"\"{exePath}\" --update-check";
        key.SetValue(EntryName, value, RegistryValueKind.String);
        Logger.Log($"Startup registration updated: {value}", LogLevel.Debug);
    }

    public static void Unregister()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
            if (key?.GetValue(EntryName) != null)
            {
                key.DeleteValue(EntryName);
                Logger.Log("Startup registration removed.", LogLevel.Debug);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException("Startup registration removal failed.", ex);
        }
    }
}
