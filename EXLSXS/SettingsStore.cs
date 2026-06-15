using Microsoft.Win32;

namespace EXLSXS
{
	// 「表示モード」「倍率」でユーザーが選んだ値を HKCU レジストリに永続化する。
	// VSTO は EXCEL.EXE にロードされ、かつ AssemblyVersion がリリースごとに変わるため、
	// バージョン依存のパスに保存される user.config (Properties.Settings) ではなく、
	// バージョン非依存の HKCU\Software\EXLSXS に保存する (リリース更新後も設定を引き継ぐ)。
	// レジストリ操作は失敗しても致命的でないため、読取・書込とも例外を握りつぶし既定動作へフォールバックする。
	internal static class SettingsStore
	{
		private const string RegistryKeyPath = @"Software\EXLSXS";
		private const string WindowViewValueName = "WindowView";
		private const string WindowZoomValueName = "WindowZoom";

		// 保存済みの表示モード (XlWindowView の数値)。未保存・読取失敗時は null。
		public static int? ReadWindowView()
		{
			return ReadInt(WindowViewValueName);
		}

		// 保存済みの倍率 (パーセント値)。未保存・読取失敗時は null。
		public static int? ReadWindowZoom()
		{
			return ReadInt(WindowZoomValueName);
		}

		public static void WriteWindowView(int value)
		{
			WriteInt(WindowViewValueName, value);
		}

		public static void WriteWindowZoom(int value)
		{
			WriteInt(WindowZoomValueName, value);
		}

		private static int? ReadInt(string valueName)
		{
			try
			{
				using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
				{
					if (key?.GetValue(valueName) is int value)
					{
						return value;
					}
				}
			}
			catch
			{
			}

			return null;
		}

		private static void WriteInt(string valueName, int value)
		{
			try
			{
				using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
				{
					key?.SetValue(valueName, value, RegistryValueKind.DWord);
				}
			}
			catch
			{
			}
		}
	}
}
