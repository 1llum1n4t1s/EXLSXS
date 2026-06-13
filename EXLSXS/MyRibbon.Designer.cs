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
			this.FontRowBox = this.Factory.CreateRibbonBox();
			this.StackBox = this.Factory.CreateRibbonBox();
			this.TabEXLSXS.SuspendLayout();
			this.GroupEXLSXS.SuspendLayout();
			this.FontRowBox.SuspendLayout();
			this.StackBox.SuspendLayout();
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
			this.GroupEXLSXS.Items.Add(this.separator1);
			this.GroupEXLSXS.Items.Add(this.StackBox);
			this.GroupEXLSXS.Label = "仕上げ";
			this.GroupEXLSXS.Name = "GroupEXLSXS";
			this.GroupEXLSXS.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GroupEXLSXS_DialogLauncherClick);
			// 
			// Finish
			// 
			this.Finish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Finish.Image = global::EXLSXS.Properties.Resources.RibbonFinish_32x32;
			this.Finish.Label = "倍率と選択位置を揃える";
			this.Finish.Name = "Finish";
			this.Finish.ScreenTip = "倍率と選択位置を揃える";
			this.Finish.ShowImage = true;
			this.Finish.SuperTip = "選択中の表示モード・倍率（「フォントを揃える」が ON のときはフォントも）をすべてのシートに適用し、各シート A1 を選択してから先頭のシートをアクティブにします。";
			this.Finish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Finish_Click);
			// 
			// AdjustFontBox
			// 
			this.AdjustFontBox.Label = "フォントを揃える";
			this.AdjustFontBox.Name = "AdjustFontBox";
			this.AdjustFontBox.ScreenTip = "フォントを揃える";
			this.AdjustFontBox.SuperTip = "ブックのすべてのセルのフォントを、右で選択したフォントに揃えます。";
			this.AdjustFontBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AdjustFontBox_Click);
			// 
			// FontBox
			// 
			this.FontBox.Label = "フォント";
			this.FontBox.Name = "FontBox";
			this.FontBox.ScreenTip = "揃えるフォント";
			this.FontBox.ShowItemImage = true;
			this.FontBox.ShowItemLabel = false;
			this.FontBox.ShowLabel = false;
			this.FontBox.SizeString = "改ページ プレビュー";
			this.FontBox.SuperTip = "「フォントを揃える」が ON のとき、すべてのシートのフォントをこの名前に揃えます。";
			// 
			// WindowViewBox
			// 
			this.WindowViewBox.Label = "表示モード";
			this.WindowViewBox.Name = "WindowViewBox";
			this.WindowViewBox.ScreenTip = "表示モード";
			this.WindowViewBox.ShowLabel = true;
			this.WindowViewBox.SizeString = "改ページ プレビュー";
			this.WindowViewBox.SuperTip = "すべてのシートをこの表示モード（標準 / ページ レイアウト / 改ページ プレビュー）に揃えます。";
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// WindowZoomBox
			// 
			this.WindowZoomBox.Label = "倍率";
			this.WindowZoomBox.Name = "WindowZoomBox";
			this.WindowZoomBox.ScreenTip = "倍率";
			this.WindowZoomBox.ShowLabel = true;
			this.WindowZoomBox.SizeString = "改ページ プレビュー";
			this.WindowZoomBox.SuperTip = "すべてのシートの表示倍率をこの値（50%〜200%）に揃えます。";
			//
			// FontRowBox
			//
			this.FontRowBox.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Horizontal;
			this.FontRowBox.Items.Add(this.AdjustFontBox);
			this.FontRowBox.Items.Add(this.FontBox);
			this.FontRowBox.Name = "FontRowBox";
			//
			// StackBox
			//
			this.StackBox.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
			this.StackBox.Items.Add(this.WindowViewBox);
			this.StackBox.Items.Add(this.WindowZoomBox);
			this.StackBox.Items.Add(this.FontRowBox);
			this.StackBox.Name = "StackBox";
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
			this.FontRowBox.ResumeLayout(false);
			this.FontRowBox.PerformLayout();
			this.StackBox.ResumeLayout(false);
			this.StackBox.PerformLayout();
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
		internal Microsoft.Office.Tools.Ribbon.RibbonBox FontRowBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonBox StackBox;
	}

	partial class ThisRibbonCollection
	{
		internal MyRibbon MyRibbon
		{
			get { return this.GetRibbon<MyRibbon>(); }
		}
	}
}
