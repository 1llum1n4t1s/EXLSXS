using System.Text.Json;

namespace EXLSXS.Host;

internal enum UpdateSourceKind
{
    Auto,
    Github,
    Simple
}

internal sealed class UpdateSettings
{
    public string Source { get; set; } = "";
    public UpdateSourceKind SourceKind { get; set; } = UpdateSourceKind.Auto;
    public string Channel { get; set; } = "win";
    public string AccessToken { get; set; } = "";
    public bool Prerelease { get; set; }
    public string ExpectedPublisherThumbprint { get; set; } = "";
    public string ExpectedPublisherSubject { get; set; } = "";

    public bool IsConfigured => !string.IsNullOrWhiteSpace(Source);
    public bool HasPublisherTrustConfiguration => !string.IsNullOrWhiteSpace(ExpectedPublisherThumbprint);

    public static UpdateSettings Load()
    {
        var settings = new UpdateSettings();
        var settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");

        if (File.Exists(settingsPath))
        {
            try
            {
                using var document = JsonDocument.Parse(File.ReadAllText(settingsPath));
                if (document.RootElement.TryGetProperty("Update", out var update))
                {
                    settings.Source = ReadString(update, "Source", settings.Source);
                    settings.SourceKind = ReadEnum(update, "SourceKind", settings.SourceKind);
                    settings.Channel = ReadString(update, "Channel", settings.Channel);
                    settings.AccessToken = ReadString(update, "AccessToken", settings.AccessToken);
                    settings.Prerelease = ReadBool(update, "Prerelease", settings.Prerelease);
                    settings.ExpectedPublisherThumbprint = ReadString(update, "ExpectedPublisherThumbprint", settings.ExpectedPublisherThumbprint);
                    settings.ExpectedPublisherSubject = ReadString(update, "ExpectedPublisherSubject", settings.ExpectedPublisherSubject);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("Update settings could not be read.", ex);
            }
        }

        settings.Source = ReadEnvironment("EXLSXS_UPDATE_SOURCE", settings.Source);
        settings.SourceKind = ReadEnvironmentEnum("EXLSXS_UPDATE_SOURCE_KIND", settings.SourceKind);
        settings.Channel = ReadEnvironment("EXLSXS_UPDATE_CHANNEL", settings.Channel);
        settings.AccessToken = ReadEnvironment("EXLSXS_UPDATE_TOKEN", settings.AccessToken);
        settings.Prerelease = ReadEnvironmentBool("EXLSXS_UPDATE_PRERELEASE", settings.Prerelease);
        settings.ExpectedPublisherThumbprint = ReadEnvironment("EXLSXS_EXPECTED_PUBLISHER_THUMBPRINT", settings.ExpectedPublisherThumbprint);
        settings.ExpectedPublisherSubject = ReadEnvironment("EXLSXS_EXPECTED_PUBLISHER_SUBJECT", settings.ExpectedPublisherSubject);

        return settings;
    }

    private static string ReadString(JsonElement element, string name, string fallback)
    {
        return element.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? fallback
            : fallback;
    }

    private static bool ReadBool(JsonElement element, string name, bool fallback)
    {
        return element.TryGetProperty(name, out var property) && property.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? property.GetBoolean()
            : fallback;
    }

    private static UpdateSourceKind ReadEnum(JsonElement element, string name, UpdateSourceKind fallback)
    {
        var value = ReadString(element, name, fallback.ToString());
        return Enum.TryParse(value, ignoreCase: true, out UpdateSourceKind parsed) ? parsed : fallback;
    }

    private static string ReadEnvironment(string name, string fallback)
    {
        var value = Environment.GetEnvironmentVariable(name);
        return string.IsNullOrWhiteSpace(value) ? fallback : value;
    }

    private static bool ReadEnvironmentBool(string name, bool fallback)
    {
        var value = Environment.GetEnvironmentVariable(name);
        return bool.TryParse(value, out var parsed) ? parsed : fallback;
    }

    private static UpdateSourceKind ReadEnvironmentEnum(string name, UpdateSourceKind fallback)
    {
        var value = Environment.GetEnvironmentVariable(name);
        return Enum.TryParse(value, ignoreCase: true, out UpdateSourceKind parsed) ? parsed : fallback;
    }
}
