using Microsoft.Win32;
using System.Runtime.Versioning;

namespace EXLSXS.Host;

[SupportedOSPlatform("windows")]
internal static class VstoRegistration
{
    private const string AddInName = "EXLSXS";
    private const string FriendlyName = "EXLSXS";
    private const string Description = "Excel workbook finishing add-in.";

    private static readonly string[] AddInRegistryPaths =
    [
        @"Software\Microsoft\Office\Excel\Addins\" + AddInName,
        @"Software\WOW6432Node\Microsoft\Office\Excel\Addins\" + AddInName
    ];

    public static void Register(AddInRegistrationMode mode)
    {
        var manifestPath = ResolveManifestPath();
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("VSTO manifest was not found.", manifestPath);
        }

        var manifest = new Uri(manifestPath).AbsoluteUri + "|vstolocal";
        foreach (var path in AddInRegistryPaths)
        {
            RegisterAt(path, manifest, mode);
        }
    }

    public static void Unregister()
    {
        foreach (var path in AddInRegistryPaths)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
                Logger.Log($"VSTO registration removed: HKCU\\{path}", LogLevel.Debug);
            }
            catch (Exception ex)
            {
                Logger.LogException($"VSTO registration removal failed: HKCU\\{path}", ex);
            }
        }
    }

    private static void RegisterAt(string registryPath, string manifest, AddInRegistrationMode mode)
    {
        using var key = Registry.CurrentUser.CreateSubKey(registryPath);
        if (key == null)
        {
            throw new InvalidOperationException($"VSTO registration key could not be created: HKCU\\{registryPath}");
        }

        key.SetValue("FriendlyName", FriendlyName, RegistryValueKind.String);
        key.SetValue("Description", Description, RegistryValueKind.String);
        if (mode == AddInRegistrationMode.ForceEnabled || key.GetValue("LoadBehavior") == null)
        {
            key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
        }
        else
        {
            Logger.Log($"VSTO LoadBehavior preserved: HKCU\\{registryPath} = {key.GetValue("LoadBehavior")}", LogLevel.Debug);
        }

        key.SetValue("Manifest", manifest, RegistryValueKind.String);
        Logger.Log($"VSTO registration updated: HKCU\\{registryPath} -> {manifest}", LogLevel.Debug);
    }

    private static string ResolveManifestPath()
    {
        var baseDir = AppContext.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(baseDir, "vsto", "EXLSXS.vsto"),
            Path.Combine(baseDir, "EXLSXS.vsto")
        };

        return candidates.FirstOrDefault(File.Exists) ?? candidates[0];
    }
}
