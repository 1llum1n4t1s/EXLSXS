namespace EXLSXS
{
	partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public MyRibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}

			base.Dispose(disposing);
		}

		#region コンポーネント デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncher = this.Factory.CreateRibbonDialogLauncher();
			this.TabEXLSXS = this.Factory.CreateRibbonTab();
			this.GroupEXLSXS = this.Factory.CreateRibbonGroup();
			this.Finish = this.Factory.CreateRibbonButton();
			this.AdjustFontBox = this.Factory.CreateRibbonCheckBox();
			this.FontBox = this.Factory.CreateRibbonDropDown();
			this.WindowViewBox = this.Factory.CreateRibbonDropDown();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.WindowZoomBox = this.Factory.CreateRibbonDropDown();
			this.TabEXLSXS.SuspendLayout();
			this.GroupEXLSXS.SuspendLayout();
			this.SuspendLayout();
			// 
			// TabEXLSXS
			// 
			this.TabEXLSXS.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.TabEXLSXS.ControlId.OfficeId = "TabHome";
			this.TabEXLSXS.Groups.Add(this.GroupEXLSXS);
			this.TabEXLSXS.Label = "TabHome";
			this.TabEXLSXS.Name = "TabEXLSXS";
			// 
			// GroupEXLSXS
			// 
			ribbonDialogLauncher.SuperTip = "ドキュメント仕上げ アドインの情報ボックスを表示します。";
			this.GroupEXLSXS.DialogLauncher = ribbonDialogLauncher;
			this.GroupEXLSXS.Items.Add(this.Finish);
			this.GroupEXLSXS.Items.Add(this.AdjustFontBox);
			this.GroupEXLSXS.Items.Add(this.FontBox);
			this.GroupEXLSXS.Items.Add(this.WindowViewBox);
			this.GroupEXLSXS.Items.Add(this.separator1);
			this.GroupEXLSXS.Items.Add(this.WindowZoomBox);
			this.GroupEXLSXS.Label = "仕上げ";
			this.GroupEXLSXS.Name = "GroupEXLSXS";
			this.GroupEXLSXS.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GroupEXLSXS_DialogLauncherClick);
			// 
			// Finish
			// 
			this.Finish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Finish.Label = "倍率と選択位置を揃える";
			this.Finish.Name = "Finish";
			this.Finish.OfficeImageId = "CreateQueryFromWizard";
			this.Finish.ScreenTip = "倍率と選択位置を揃える";
			this.Finish.ShowImage = true;
			this.Finish.SuperTip = "すべてのシートを「標準ビュー」「倍率100%」「A1セル選択」にしてから先頭のシートをアクティブにします。";
			this.Finish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Finish_Click);
			// 
			// AdjustFontBox
			// 
			this.AdjustFontBox.Label = "フォントを揃える";
			this.AdjustFontBox.Name = "AdjustFontBox";
			this.AdjustFontBox.ScreenTip = "フォントを揃える";
			this.AdjustFontBox.SuperTip = "ブックのすべてのセルのフォントを下で選択したフォントに揃えます。";
			this.AdjustFontBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AdjustFontBox_Click);
			// 
			// FontBox
			// 
			this.FontBox.Label = "フォント";
			this.FontBox.Name = "FontBox";
			this.FontBox.ShowItemLabel = false;
			this.FontBox.ShowLabel = false;
			this.FontBox.SizeString = "1234567890123";
			// 
			// WindowViewBox
			// 
			this.WindowViewBox.Label = "表示モード";
			this.WindowViewBox.Name = "WindowViewBox";
			this.WindowViewBox.ShowLabel = false;
			this.WindowViewBox.SizeString = "改ページ プレビュー";
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// WindowZoomBox
			// 
			this.WindowZoomBox.Label = "倍率";
			this.WindowZoomBox.Name = "WindowZoomBox";
			this.WindowZoomBox.ShowLabel = false;
			this.WindowZoomBox.SizeString = "100%";
			// 
			// MyRibbon
			// 
			this.Name = "MyRibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.TabEXLSXS);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
			this.TabEXLSXS.ResumeLayout(false);
			this.TabEXLSXS.PerformLayout();
			this.GroupEXLSXS.ResumeLayout(false);
			this.GroupEXLSXS.PerformLayout();
			this.ResumeLayout(false);
		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab TabEXLSXS;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupEXLSXS;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Finish;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox AdjustFontBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown FontBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown WindowViewBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown WindowZoomBox;
	}

	partial class ThisRibbonCollection
	{
		internal MyRibbon MyRibbon
		{
			get { return this.GetRibbon<MyRibbon>(); }
		}
	}
}
