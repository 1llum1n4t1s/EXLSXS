using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace EXLSXS
{
	public partial class MyRibbon
	{
		public bool AdjustFont
		{
			get { return AdjustFontBox.Checked; }
		}

		public string AdjustFontName
		{
			get { return FontBox.SelectedItem?.Label ?? string.Empty; }
		}

		public Excel.XlWindowView AdjustView
		{
			get { return (Excel.XlWindowView)WindowViewBox.SelectedItem.Tag; }
		}

		public int AdjustZoom
		{
			get { return (int)WindowZoomBox.SelectedItem.Tag; }
		}

		public bool AdjustGrid
		{
			get { return AdjustGridBox.Checked; }
		}

		public double AdjustGridSizeMm
		{
			get { return (double)GridSizeBox.SelectedItem.Tag; }
		}

		public bool AdjustNumberFormat
		{
			get { return AdjustNumberFormatBox.Checked; }
		}

		public string AdjustNumberFormatCode
		{
			get { return (string)NumberFormatBox.SelectedItem.Tag; }
		}

		private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			try
			{
				LocalBuildLabel.Visible = IsLocalBuild();
			}
			catch
			{
			}

			// フォント一覧の取得だけ Excel COM (StandardFont) に依存し失敗しうるため、
			// COM 非依存の他の設定項目を巻き込まないよう独立した try/catch にする。
			try
			{
				SetupFontBox();
				AdjustFontBox.Checked = false;
				FontBox.Enabled = false;
			}
			catch
			{
				AdjustFontBox.Checked = false;
				AdjustFontBox.Enabled = false;
				FontBox.Enabled = false;
			}

			try
			{
				SetupWindowViewBox();
				SetupWindowZoomBox();
			}
			catch
			{
			}

			try
			{
				SetupGridSizeBox();
				SetupNumberFormatBox();
				AdjustGridBox.Checked = false;
				GridSizeBox.Enabled = false;
				AdjustNumberFormatBox.Checked = false;
				NumberFormatBox.Enabled = false;
			}
			catch
			{
				AdjustGridBox.Checked = false;
				AdjustGridBox.Enabled = false;
				GridSizeBox.Enabled = false;
				AdjustNumberFormatBox.Checked = false;
				AdjustNumberFormatBox.Enabled = false;
				NumberFormatBox.Enabled = false;
			}
		}

		// 実行中の EXLSXS.dll が Velopack 正規インストール先 (%LocalAppData%\EXLSXS\current\vsto\)
		// から読み込まれているかを判定する。一致しなければ bin\Debug/Release 等からのローカルビルドとみなす。
		//
		// ⚠ Assembly.Location は使わない: vstolocal 登録でも strong-name 付きアセンブリは Fusion が
		// シャドウコピーし、Location は %LocalAppData%\assembly\dl3\... を返す (本番/ローカル双方が dl3 経由に
		// なるため Location では区別できず、本番でも「ローカル」と誤判定する)。CodeBase はシャドウコピー前の
		// 元のロード元 URL を返すため、こちらで本番インストール先配下かどうかを判定する。
		// 判定不能時は「ローカル実行」側に倒す (本番で誤ってマークが出るより、開発中にマークが出ない方が困る)。
		private static bool IsLocalBuild()
		{
			try
			{
				string productionRoot = Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
					"EXLSXS", "current", "vsto");

				Assembly assembly = Assembly.GetExecutingAssembly();
				string originPath = null;
				try
				{
					string codeBase = assembly.CodeBase;
					if (!string.IsNullOrEmpty(codeBase))
					{
						originPath = new Uri(codeBase).LocalPath;
					}
				}
				catch
				{
				}

				// CodeBase が取れない環境では Location にフォールバックする
				// (dl3 を返すと誤判定しうるが、CodeBase が無い方が例外的なので最後の手段)。
				if (string.IsNullOrEmpty(originPath))
				{
					originPath = assembly.Location;
				}

				return string.IsNullOrEmpty(originPath)
					|| !originPath.StartsWith(productionRoot, StringComparison.OrdinalIgnoreCase);
			}
			catch
			{
				return true;
			}
		}

		public void RefreshStatus(bool isWorkbookOpened)
		{
			Finish.Enabled = isWorkbookOpened;
		}

		private void Finish_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.DoFinish();
		}

		private void SetupFontBox()
		{
			// StandardFont の取得は COM 往復で失敗しうるため、重いプレビュー画像生成ループの前に
			// 済ませておく (失敗時に無駄な画像生成をせず即座に catch へ抜けられる)。
			string standardFont = Globals.ThisAddIn.Application.StandardFont;

			// Excel のフォント一覧のように、各フォント名をそのフォント自身で描画したプレビュー画像を付ける。
			// (画像生成でリボン Load が数百 ms 伸び、生成した Bitmap は UI 生存中常駐するが、
			//  「選択するフォントの見た目が分かる」ことを優先する。)
			foreach (FontFamily fontFamily in new InstalledFontCollection().Families)
			{
				if (fontFamily.IsStyleAvailable(FontStyle.Regular)
					&& fontFamily.IsStyleAvailable(FontStyle.Bold)
					&& fontFamily.IsStyleAvailable(FontStyle.Italic)
					&& fontFamily.IsStyleAvailable(FontStyle.Strikeout)
					&& fontFamily.IsStyleAvailable(FontStyle.Underline))
				{
					Bitmap image = CreateFontPreviewImage(fontFamily);
					RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
					item.Label = fontFamily.Name;
					item.Image = image;
					FontBox.Items.Add(item);
				}

				fontFamily.Dispose();
			}

			RibbonDropDownItem selectedFont = FontBox.Items.FirstOrDefault(item => item.Label == standardFont);
			if (selectedFont != null)
			{
				FontBox.SelectedItem = selectedFont;
				return;
			}

			if (FontBox.Items.Count > 0)
			{
				FontBox.SelectedItemIndex = 0;
			}
		}

		// フォント名を、そのフォント自身で描画したプレビュー画像にしてドロップダウン一覧に表示する
		// (Excel のフォント一覧と同様)。閉じた FontBox は item.Label(フォント名テキスト)を表示するため、
		// この画像の高さは閉じた箱の高さ・リボンの行間には影響しない (一覧の項目高さだけを決める)。
		// TextRenderer.DrawText (GDI・不透明描画) を使う。GDI+ の DrawString を透明背景に描くと
		// リボン上で文字が出ず空画像になるため、ここは GDI 描画を維持する。
		private static Bitmap CreateFontPreviewImage(FontFamily fontFamily)
		{
			Bitmap image = new Bitmap(200, 20);
			using (Graphics graphics = Graphics.FromImage(image))
			using (Font font = new Font(fontFamily, 16f, GraphicsUnit.Pixel))
			{
				graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
				TextRenderer.DrawText(graphics, fontFamily.Name, font, Point.Empty, SystemColors.MenuText);
			}

			return image;
		}

		// (ラベル, 値) の一覧から RibbonDropDownItem を生成して box に追加する。
		// savedValue (前回ユーザーが選んだ値) があり、対応する項目があればそれを選択し、
		// 無ければ defaultLabel の項目を選択する。valueToInt は Tag の値をレジストリ保存値 (int) に対応付ける。
		// WindowViewBox / WindowZoomBox が同じ骨格なので共通化する。
		private static void PopulateDropDown<T>(RibbonDropDown box, (string Label, T Value)[] items, string defaultLabel, int? savedValue = null, Func<T, int> valueToInt = null)
		{
			foreach (var (label, value) in items)
			{
				RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				item.Label = label;
				item.Tag = value;
				box.Items.Add(item);
			}

			if (savedValue.HasValue)
			{
				RibbonDropDownItem savedItem = box.Items.FirstOrDefault(item => valueToInt((T)item.Tag) == savedValue.Value);
				if (savedItem != null)
				{
					box.SelectedItem = savedItem;
					return;
				}
			}

			box.SelectedItem = box.Items.First(item => item.Label == defaultLabel);
		}

		private void SetupWindowViewBox()
		{
			PopulateDropDown(WindowViewBox, new (string Label, Excel.XlWindowView Value)[]
			{
				("標準", Excel.XlWindowView.xlNormalView),
				("ページ レイアウト", Excel.XlWindowView.xlPageLayoutView),
				("改ページ プレビュー", Excel.XlWindowView.xlPageBreakPreview)
			}, "標準", SettingsStore.ReadWindowView(), value => (int)value);
		}

		private void SetupWindowZoomBox()
		{
			PopulateDropDown(WindowZoomBox, new (string Label, int Value)[]
			{
				("50%", 50),
				("55%", 55),
				("60%", 60),
				("65%", 65),
				("70%", 70),
				("75%", 75),
				("80%", 80),
				("85%", 85),
				("90%", 90),
				("95%", 95),
				("100%", 100),
				("105%", 105),
				("110%", 110),
				("115%", 115),
				("120%", 120),
				("125%", 125),
				("130%", 130),
				("135%", 135),
				("140%", 140),
				("145%", 145),
				("150%", 150),
				("200%", 200)
			}, "100%", SettingsStore.ReadWindowZoom(), value => value);
		}

		private void SetupGridSizeBox()
		{
			PopulateDropDown(GridSizeBox, new (string Label, double Value)[]
			{
				("2mm", 2.0),
				("3mm", 3.0),
				("4mm", 4.0),
				("5mm", 5.0),
				("6mm", 6.0),
				("8mm", 8.0),
				("10mm", 10.0),
				("15mm", 15.0),
				("20mm", 20.0)
			}, "5mm");
		}

		// 標準の「セルの書式設定」ダイアログの数値の書式ドロップダウンと同じ項目構成
		// (標準/数値/通貨/会計/日付(短い/長い)/時刻/パーセンテージ/分数/指数/文字列)。
		private void SetupNumberFormatBox()
		{
			PopulateDropDown(NumberFormatBox, new (string Label, string Value)[]
			{
				("標準", "General"),
				("数値", "0.00"),
				("通貨", "\"¥\"#,##0.00_);(\"¥\"#,##0.00)"),
				("会計", "_(\"¥\"* #,##0.00_);_(\"¥\"* (#,##0.00);_(\"¥\"* \"-\"??_);_(@_)"),
				("日付の短い形式", "yyyy/m/d"),
				("日付の長い形式", "yyyy\"年\"m\"月\"d\"日\""),
				("時刻", "h:mm:ss"),
				("パーセンテージ", "0.00%"),
				("分数", "# ?/?"),
				("指数", "0.00E+00"),
				("文字列", "@")
			}, "標準");
		}

		// 表示モードをユーザーが変更したら選択値を保存する。
		// (コードで SelectedItem を代入する復元時には SelectionChanged は発火しないため、
		//  このハンドラはユーザー操作時のみ呼ばれる。)
		private void WindowViewBox_SelectionChanged(object sender, RibbonControlEventArgs e)
		{
			if (WindowViewBox.SelectedItem?.Tag is Excel.XlWindowView view)
			{
				SettingsStore.WriteWindowView((int)view);
			}
		}

		// 倍率をユーザーが変更したら選択値を保存する。
		private void WindowZoomBox_SelectionChanged(object sender, RibbonControlEventArgs e)
		{
			if (WindowZoomBox.SelectedItem?.Tag is int zoom)
			{
				SettingsStore.WriteWindowZoom(zoom);
			}
		}

		private void AdjustFontBox_Click(object sender, RibbonControlEventArgs e)
		{
			FontBox.Enabled = AdjustFontBox.Checked;
		}

		private void AdjustGridBox_Click(object sender, RibbonControlEventArgs e)
		{
			GridSizeBox.Enabled = AdjustGridBox.Checked;
		}

		private void AdjustNumberFormatBox_Click(object sender, RibbonControlEventArgs e)
		{
			NumberFormatBox.Enabled = AdjustNumberFormatBox.Checked;
		}

		private void GroupEXLSXS_DialogLauncherClick(object sender, RibbonControlEventArgs e)
		{
			try
			{
				new AboutBox().ShowDialog();
			}
			catch
			{
			}
		}
	}
}
