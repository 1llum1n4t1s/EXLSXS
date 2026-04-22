using System;
using System.Collections.Generic;
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
			get { return FontBox.SelectedItem.Label; }
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

			RibbonDropDownItem selectedFont = FontBox.Items.FirstOrDefault(item => item.Label == Globals.ThisAddIn.Application.StandardFont);
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

		private void SetupWindowViewBox()
		{
			List<Tuple<string, Excel.XlWindowView>> views = new List<Tuple<string, Excel.XlWindowView>>
			{
				new Tuple<string, Excel.XlWindowView>("標準", Excel.XlWindowView.xlNormalView),
				new Tuple<string, Excel.XlWindowView>("ページ レイアウト", Excel.XlWindowView.xlPageLayoutView),
				new Tuple<string, Excel.XlWindowView>("改ページ プレビュー", Excel.XlWindowView.xlPageBreakPreview)
			};

			foreach (Tuple<string, Excel.XlWindowView> view in views)
			{
				RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				item.Label = view.Item1;
				item.Tag = view.Item2;
				WindowViewBox.Items.Add(item);
			}

			WindowViewBox.SelectedItem = WindowViewBox.Items.First(item => item.Label == "標準");
		}

		private void SetupWindowZoomBox()
		{
			List<Tuple<string, int>> zoomLevels = new List<Tuple<string, int>>
			{
				new Tuple<string, int>("50%", 50),
				new Tuple<string, int>("55%", 55),
				new Tuple<string, int>("60%", 60),
				new Tuple<string, int>("65%", 65),
				new Tuple<string, int>("70%", 70),
				new Tuple<string, int>("75%", 75),
				new Tuple<string, int>("80%", 80),
				new Tuple<string, int>("85%", 85),
				new Tuple<string, int>("90%", 90),
				new Tuple<string, int>("95%", 95),
				new Tuple<string, int>("100%", 100),
				new Tuple<string, int>("105%", 105),
				new Tuple<string, int>("110%", 110),
				new Tuple<string, int>("115%", 115),
				new Tuple<string, int>("120%", 120),
				new Tuple<string, int>("125%", 125),
				new Tuple<string, int>("130%", 130),
				new Tuple<string, int>("135%", 135),
				new Tuple<string, int>("140%", 140),
				new Tuple<string, int>("145%", 145),
				new Tuple<string, int>("150%", 150),
				new Tuple<string, int>("200%", 200)
			};

			foreach (Tuple<string, int> zoomLevel in zoomLevels)
			{
				RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				item.Label = zoomLevel.Item1;
				item.Tag = zoomLevel.Item2;
				WindowZoomBox.Items.Add(item);
			}

			WindowZoomBox.SelectedItem = WindowZoomBox.Items.First(item => item.Label == "100%");
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
