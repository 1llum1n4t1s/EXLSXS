namespace EXLSXS.Host;

internal static class InstalledAppMaintenance
{
    public static void Register(AddInRegistrationMode mode, bool allowPrerequisiteInstall = false)
    {
        PrerequisiteChecker.EnsureReady(allowPrerequisiteInstall);
        VstoRegistration.Register(mode);
        StartupRegistration.Register();
        StartupRegistration.UnregisterRegistrationRetry();
    }

    public static void Unregister()
    {
        StartupRegistration.Unregister();
        VstoRegistration.Unregister();
    }
}
