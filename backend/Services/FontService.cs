using System.Drawing.Text;

namespace EXLSXS.API.Services;

public interface IFontService
{
    IEnumerable<string> GetInstalledFonts();
}

public class FontService : IFontService
{
    public IEnumerable<string> GetInstalledFonts()
    {
        // Windows環境でのみ動作する想定
        if (!OperatingSystem.IsWindows())
        {
            return new[] { "Arial", "Calibri", "Times New Roman" }; // フォールバック
        }

        using var fonts = new InstalledFontCollection();
        return fonts.Families
            .Where(f => f.IsStyleAvailable(System.Drawing.FontStyle.Regular)) 
            .Select(f => f.Name)
            .OrderBy(n => n)
            .ToList();
    }
}
