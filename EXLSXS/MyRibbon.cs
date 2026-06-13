using System;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
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

		private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			try
			{
				SetupFontBox();
				SetupWindowViewBox();
				SetupWindowZoomBox();
				AdjustFontBox.Checked = false;
				FontBox.Enabled = false;
			}
			catch
			{
				AdjustFontBox.Checked = false;
				AdjustFontBox.Enabled = false;
				FontBox.Enabled = false;
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

			// StandardFont の取得は COM 往復なので、述語内で毎回評価せずループ前に 1 回退避する。
			string standardFont = Globals.ThisAddIn.Application.StandardFont;
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

		// (ラベル, 値) の一覧から RibbonDropDownItem を生成して box に追加し、defaultLabel の項目を選択する。
		// WindowViewBox / WindowZoomBox が同じ骨格なので共通化する。
		private static void PopulateDropDown<T>(RibbonDropDown box, (string Label, T Value)[] items, string defaultLabel)
		{
			foreach (var (label, value) in items)
			{
				RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				item.Label = label;
				item.Tag = value;
				box.Items.Add(item);
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
			}, "標準");
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
			}, "100%");
		}

		private void AdjustFontBox_Click(object sender, RibbonControlEventArgs e)
		{
			FontBox.Enabled = AdjustFontBox.Checked;
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
