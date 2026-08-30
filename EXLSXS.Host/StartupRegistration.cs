using Microsoft.Win32;
using System.Runtime.Versioning;

namespace EXLSXS.Host;

[SupportedOSPlatform("windows")]
internal static class StartupRegistration
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string RunOnceKeyPath = @"Software\Microsoft\Windows\CurrentVersion\RunOnce";
    private const string EntryName = "EXLSXS";
    private const string RegistrationRetryEntryName = "EXLSXS Registration Retry";

    public static void Register()
    {
        var exePath = GetProcessPath();

        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath);
        if (key == null)
        {
            throw new InvalidOperationException("Startup registration key could not be opened.");
        }

        var value = BuildCommand(exePath, "--update-check");
        key.SetValue(EntryName, value, RegistryValueKind.String);
        Logger.Log($"Startup registration updated: {value}", LogLevel.Debug);
    }

    public static void RegisterRegistrationRetry()
    {
        var exePath = GetProcessPath();

        using var key = Registry.CurrentUser.CreateSubKey(RunOnceKeyPath);
        if (key == null)
        {
            throw new InvalidOperationException("Registration retry key could not be opened.");
        }

        var value = BuildCommand(exePath, "--register");
        key.SetValue(RegistrationRetryEntryName, value, RegistryValueKind.String);
        Logger.Log($"One-time registration retry scheduled: {value}", LogLevel.Debug);
    }

    public static void UnregisterRegistrationRetry()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunOnceKeyPath, writable: true);
            if (key?.GetValue(RegistrationRetryEntryName) != null)
            {
                key.DeleteValue(RegistrationRetryEntryName);
                Logger.Log("One-time registration retry removed.", LogLevel.Debug);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException("Registration retry removal failed.", ex);
        }
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

        UnregisterRegistrationRetry();
    }

    internal static string BuildCommand(string exePath, string argument)
    {
        if (string.IsNullOrWhiteSpace(exePath))
        {
            throw new ArgumentException("Executable path must not be empty.", nameof(exePath));
        }

        if (string.IsNullOrWhiteSpace(argument))
        {
            throw new ArgumentException("Argument must not be empty.", nameof(argument));
        }

        return $"\"{exePath}\" {argument}";
    }

    private static string GetProcessPath()
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exePath))
        {
            throw new InvalidOperationException("Startup registration failed because process path is empty.");
        }

        return exePath;
    }
}
