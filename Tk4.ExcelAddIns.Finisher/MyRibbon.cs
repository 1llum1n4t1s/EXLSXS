using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace Tk4.ExcelAddIns.Finisher
{
	// Token: 0x02000003 RID: 3
	public class MyRibbon : RibbonBase
	{
		// Token: 0x06000006 RID: 6 RVA: 0x000026F0 File Offset: 0x000008F0
		public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
		{
			this.InitializeComponent();
		}

		// Token: 0x06000007 RID: 7 RVA: 0x00002708 File Offset: 0x00000908
		protected override void Dispose(bool disposing)
		{
			if (disposing && this.components != null)
			{
				this.components.Dispose();
			}
			base.Dispose(disposing);
		}

		// Token: 0x06000008 RID: 8 RVA: 0x00002728 File Offset: 0x00000928
		private void InitializeComponent()
		{
			RibbonDialogLauncher ribbonDialogLauncher = base.Factory.CreateRibbonDialogLauncher();
			this.TabFinisher = base.Factory.CreateRibbonTab();
			this.GroupFinisher = base.Factory.CreateRibbonGroup();
			this.AdjustFontBox = base.Factory.CreateRibbonCheckBox();
			this.FontBox = base.Factory.CreateRibbonDropDown();
			this.WindowZoomBox = base.Factory.CreateRibbonDropDown();
			this.WindowViewBox = base.Factory.CreateRibbonDropDown();
			this.separator1 = base.Factory.CreateRibbonSeparator();
			this.Finish = base.Factory.CreateRibbonButton();
			this.TabFinisher.SuspendLayout();
			this.GroupFinisher.SuspendLayout();
			this.TabFinisher.ControlId.ControlIdType = RibbonControlIdType.Office;
			this.TabFinisher.ControlId.OfficeId = "TabHome";
			this.TabFinisher.Groups.Add(this.GroupFinisher);
			this.TabFinisher.Label = "TabHome";
			this.TabFinisher.Name = "TabFinisher";
			ribbonDialogLauncher.SuperTip = "ドキュメント仕上げ アドインの情報ボックスを表示します。";
			this.GroupFinisher.DialogLauncher = ribbonDialogLauncher;
			this.GroupFinisher.Items.Add(this.Finish);
			this.GroupFinisher.Items.Add(this.AdjustFontBox);
			this.GroupFinisher.Items.Add(this.FontBox);
			this.GroupFinisher.Items.Add(this.WindowViewBox);
			this.GroupFinisher.Items.Add(this.separator1);
			this.GroupFinisher.Items.Add(this.WindowZoomBox);
			this.GroupFinisher.Label = "仕上げ";
			this.GroupFinisher.Name = "GroupFinisher";
			this.GroupFinisher.DialogLauncherClick += this.GroupFinisher_DialogLauncherClick;
			this.AdjustFontBox.Label = "フォントを揃える";
			this.AdjustFontBox.Name = "AdjustFontBox";
			this.AdjustFontBox.ScreenTip = "フォントを揃える";
			this.AdjustFontBox.SuperTip = "ブックのすべてのセルのフォントを下で選択したフォントに揃えます。";
			this.AdjustFontBox.Click += this.AdjustFontBox_Click;
			this.FontBox.Label = "フォント";
			this.FontBox.Name = "FontBox";
			this.FontBox.ShowItemLabel = false;
			this.FontBox.ShowLabel = false;
			this.FontBox.SizeString = "1234567890123";
			this.WindowZoomBox.Label = "倍率";
			this.WindowZoomBox.Name = "WindowZoomBox";
			this.WindowZoomBox.ShowLabel = false;
			this.WindowZoomBox.SizeString = "100%";
			this.WindowViewBox.Label = "表示モード";
			this.WindowViewBox.Name = "WindowViewBox";
			this.WindowViewBox.ShowLabel = false;
			this.WindowViewBox.SizeString = "改ページ プレビュー";
			this.separator1.Name = "separator1";
			this.Finish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Finish.Label = "倍率と選択位置を揃える";
			this.Finish.Name = "Finish";
			this.Finish.OfficeImageId = "CreateQueryFromWizard";
			this.Finish.ScreenTip = "倍率と選択位置を揃える";
			this.Finish.ShowImage = true;
			this.Finish.SuperTip = "すべてのシートを「標準ビュー」「倍率100%」「A1セル選択」にしてから先頭のシートをアクティブにします。";
			this.Finish.Click += this.Finish_Click;
			base.Name = "MyRibbon";
			base.RibbonType = "Microsoft.Excel.Workbook";
			base.Tabs.Add(this.TabFinisher);
			base.Load += this.MyRibbon_Load;
			this.TabFinisher.ResumeLayout(false);
			this.TabFinisher.PerformLayout();
			this.GroupFinisher.ResumeLayout(false);
			this.GroupFinisher.PerformLayout();
		}

		// Token: 0x17000001 RID: 1
		// (get) Token: 0x06000009 RID: 9 RVA: 0x00002B1A File Offset: 0x00000D1A
		public bool AdjustFont
		{
			get
			{
				return this.AdjustFontBox.Checked;
			}
		}

		// Token: 0x17000002 RID: 2
		// (get) Token: 0x0600000A RID: 10 RVA: 0x00002B27 File Offset: 0x00000D27
		public string AdjustFontName
		{
			get
			{
				return this.FontBox.SelectedItem.Label;
			}
		}

		// Token: 0x17000003 RID: 3
		// (get) Token: 0x0600000B RID: 11 RVA: 0x00002B39 File Offset: 0x00000D39
		public XlWindowView AdjustView
		{
			get
			{
				return (XlWindowView)this.WindowViewBox.SelectedItem.Tag;
			}
		}

		// Token: 0x17000004 RID: 4
		// (get) Token: 0x0600000C RID: 12 RVA: 0x00002B50 File Offset: 0x00000D50
		public int AdjustZoom
		{
			get
			{
				return (int)this.WindowZoomBox.SelectedItem.Tag;
			}
		}

		// Token: 0x0600000D RID: 13 RVA: 0x00002B68 File Offset: 0x00000D68
		private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			try
			{
				this.SetupFontBox();
				this.SetupWindowViewBox();
				this.SetupWindowZoomBox();
				this.AdjustFontBox.Checked = false;
				this.FontBox.Enabled = false;
			}
			catch
			{
				this.AdjustFontBox.Checked = false;
				this.AdjustFontBox.Enabled = false;
				this.FontBox.Enabled = false;
			}
		}

		// Token: 0x0600000E RID: 14 RVA: 0x00002BD8 File Offset: 0x00000DD8
		public void RefreshStatus(bool isWorkbookOpened)
		{
			if (isWorkbookOpened)
			{
				this.Finish.Enabled = true;
				return;
			}
			this.Finish.Enabled = false;
		}

		// Token: 0x0600000F RID: 15 RVA: 0x00002BF6 File Offset: 0x00000DF6
		private void Finish_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.DoFinish();
		}

		// Token: 0x06000010 RID: 16 RVA: 0x00002C20 File Offset: 0x00000E20
		private void SetupFontBox()
		{
			foreach (FontFamily fontFamily in new InstalledFontCollection().Families)
			{
				if (fontFamily.IsStyleAvailable(FontStyle.Regular) && fontFamily.IsStyleAvailable(FontStyle.Bold) && fontFamily.IsStyleAvailable(FontStyle.Italic) && fontFamily.IsStyleAvailable(FontStyle.Strikeout) && fontFamily.IsStyleAvailable(FontStyle.Underline))
				{
					Bitmap image = new Bitmap(200, 20);
					using (Graphics graphics = Graphics.FromImage(image))
					{
						using (System.Drawing.Font font = new System.Drawing.Font(fontFamily, 16f, GraphicsUnit.Pixel))
						{
							graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
							TextRenderer.DrawText(graphics, fontFamily.Name, font, Point.Empty, SystemColors.MenuText);
						}
					}
					RibbonDropDownItem ribbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
					ribbonDropDownItem.Label = fontFamily.Name;
					ribbonDropDownItem.Image = image;
					this.FontBox.Items.Add(ribbonDropDownItem);
				}
				fontFamily.Dispose();
			}
			RibbonDropDownItem ribbonDropDownItem2 = this.FontBox.Items.FirstOrDefault((RibbonDropDownItem item) => item.Label == Globals.ThisAddIn.Application.StandardFont);
			if (ribbonDropDownItem2 != null)
			{
				this.FontBox.SelectedItem = ribbonDropDownItem2;
				return;
			}
			this.FontBox.SelectedItemIndex = 0;
		}

		// Token: 0x06000011 RID: 17 RVA: 0x00002DA8 File Offset: 0x00000FA8
		private void SetupWindowViewBox()
		{
			List<Tuple<string, XlWindowView>> list = new List<Tuple<string, XlWindowView>>
			{
				new Tuple<string, XlWindowView>("標準", XlWindowView.xlNormalView),
				new Tuple<string, XlWindowView>("ページ レイアウト", XlWindowView.xlPageLayoutView),
				new Tuple<string, XlWindowView>("改ページ プレビュー", XlWindowView.xlPageBreakPreview)
			};
			foreach (Tuple<string, XlWindowView> tuple in list)
			{
				RibbonDropDownItem ribbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				ribbonDropDownItem.Label = tuple.Item1;
				ribbonDropDownItem.Tag = tuple.Item2;
				this.WindowViewBox.Items.Add(ribbonDropDownItem);
			}
			this.WindowViewBox.SelectedItem = this.WindowViewBox.Items.First((RibbonDropDownItem item) => item.Label == "標準");
		}

		// Token: 0x06000012 RID: 18 RVA: 0x00002EB8 File Offset: 0x000010B8
		private void SetupWindowZoomBox()
		{
			List<Tuple<string, int>> list = new List<Tuple<string, int>>
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
			foreach (Tuple<string, int> tuple in list)
			{
				RibbonDropDownItem ribbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
				ribbonDropDownItem.Label = tuple.Item1;
				ribbonDropDownItem.Tag = tuple.Item2;
				this.WindowZoomBox.Items.Add(ribbonDropDownItem);
			}
			this.WindowZoomBox.SelectedItem = this.WindowZoomBox.Items.First((RibbonDropDownItem item) => item.Label == "100%");
		}

		// Token: 0x06000013 RID: 19 RVA: 0x0000311C File Offset: 0x0000131C
		private void AdjustFontBox_Click(object sender, RibbonControlEventArgs e)
		{
			this.FontBox.Enabled = this.AdjustFontBox.Checked;
		}

		// Token: 0x06000014 RID: 20 RVA: 0x00003134 File Offset: 0x00001334
		private void GroupFinisher_DialogLauncherClick(object sender, RibbonControlEventArgs e)
		{
			try
			{
				new AboutBox().ShowDialog();
			}
			catch
			{
			}
		}

		// Token: 0x0400000A RID: 10
		private IContainer components;

		// Token: 0x0400000B RID: 11
		internal RibbonTab TabFinisher;

		// Token: 0x0400000C RID: 12
		internal RibbonGroup GroupFinisher;

		// Token: 0x0400000D RID: 13
		internal RibbonButton Finish;

		// Token: 0x0400000E RID: 14
		private RibbonDropDown FontBox;

		// Token: 0x0400000F RID: 15
		private RibbonCheckBox AdjustFontBox;

		// Token: 0x04000010 RID: 16
		private RibbonDropDown WindowZoomBox;

		// Token: 0x04000011 RID: 17
		private RibbonDropDown WindowViewBox;

		// Token: 0x04000012 RID: 18
		internal RibbonSeparator separator1;
	}
}
