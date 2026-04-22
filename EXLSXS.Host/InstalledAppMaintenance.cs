namespace EXLSXS.Host;

internal static class InstalledAppMaintenance
{
    public static void Register(AddInRegistrationMode mode)
    {
        PrerequisiteChecker.EnsureReady();
        VstoRegistration.Register(mode);
        StartupRegistration.Register();
    }

    public static void Unregister()
    {
        StartupRegistration.Unregister();
        VstoRegistration.Unregister();
    }
}
