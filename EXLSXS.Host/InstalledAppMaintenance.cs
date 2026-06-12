namespace EXLSXS.Host;

internal static class InstalledAppMaintenance
{
    public static void Register(AddInRegistrationMode mode, bool allowPrerequisiteInstall = false)
    {
        PrerequisiteChecker.EnsureReady(allowPrerequisiteInstall);
        VstoRegistration.Register(mode);
        StartupRegistration.Register();
    }

    public static void Unregister()
    {
        StartupRegistration.Unregister();
        VstoRegistration.Unregister();
    }
}
