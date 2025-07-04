using System.Diagnostics;
using System.IO;
using System.Text.Json;
namespace EXLSXS;

/// <summary>
/// アプリケーション設定クラス
/// </summary>
public class Settings
{
    public bool CreateBackup { get; set; } = true;
    public bool OptimizeFile { get; set; } = true;
    public bool ImprovePerformance { get; set; } = true;
    public bool RemoveUnusedStyles { get; set; } = true;
    public bool CompressImages { get; set; } = false;
    public string RepairLevel { get; set; } = "Safe"; // Safe, Standard, Aggressive
    public bool ShowLogAfterProcess { get; set; } = true;
    public string LastSelectedFolder { get; set; } = "";
    public int MaxCellsToProcess { get; set; } = 100000;
    public bool EnableDetailedLogging { get; set; } = true;

    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EXLSXS_Tool",
        "settings.json"
    );

    /// <summary>
    /// 設定を読み込み
    /// </summary>
    public static Settings Load()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                return JsonSerializer.Deserialize<Settings>(json) ?? new Settings();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"設定読み込みエラー: {ex.Message}");
        }

        return new Settings();
    }

    /// <summary>
    /// 設定を保存
    /// </summary>
    public void Save()
    {
        try
        {
            var directory = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"設定保存エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// デフォルト設定にリセット
    /// </summary>
    public void ResetToDefault()
    {
        CreateBackup = true;
        OptimizeFile = true;
        ImprovePerformance = true;
        RemoveUnusedStyles = true;
        CompressImages = false;
        RepairLevel = "Safe";
        ShowLogAfterProcess = true;
        LastSelectedFolder = "";
        MaxCellsToProcess = 100000;
        EnableDetailedLogging = true;
    }
}