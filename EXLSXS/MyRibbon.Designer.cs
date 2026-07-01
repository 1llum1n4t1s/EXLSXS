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
			this.separator2 = this.Factory.CreateRibbonSeparator();
			this.AdjustGridBox = this.Factory.CreateRibbonCheckBox();
			this.GridSizeBox = this.Factory.CreateRibbonDropDown();
			this.AdjustNumberFormatBox = this.Factory.CreateRibbonCheckBox();
			this.NumberFormatBox = this.Factory.CreateRibbonDropDown();
			this.GridRowBox = this.Factory.CreateRibbonBox();
			this.NumberFormatRowBox = this.Factory.CreateRibbonBox();
			this.FormatStackBox = this.Factory.CreateRibbonBox();
			this.LocalBuildLabel = this.Factory.CreateRibbonLabel();
			this.TabEXLSXS.SuspendLayout();
			this.GroupEXLSXS.SuspendLayout();
			this.FontRowBox.SuspendLayout();
			this.StackBox.SuspendLayout();
			this.GridRowBox.SuspendLayout();
			this.NumberFormatRowBox.SuspendLayout();
			this.FormatStackBox.SuspendLayout();
			this.SuspendLayout();
			//
			// TabEXLSXS
			//
			this.TabEXLSXS.Groups.Add(this.GroupEXLSXS);
			this.TabEXLSXS.Label = "EXLSXS";
			this.TabEXLSXS.Name = "TabEXLSXS";
			// 
			// GroupEXLSXS
			// 
			ribbonDialogLauncher.SuperTip = "ドキュメント仕上げ アドインの情報ボックスを表示します。";
			this.GroupEXLSXS.DialogLauncher = ribbonDialogLauncher;
			this.GroupEXLSXS.Items.Add(this.Finish);
			this.GroupEXLSXS.Items.Add(this.separator1);
			this.GroupEXLSXS.Items.Add(this.StackBox);
			this.GroupEXLSXS.Items.Add(this.separator2);
			this.GroupEXLSXS.Items.Add(this.FormatStackBox);
			this.GroupEXLSXS.Label = "仕上げ";
			this.GroupEXLSXS.Name = "GroupEXLSXS";
			this.GroupEXLSXS.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GroupEXLSXS_DialogLauncherClick);
			// 
			// Finish
			// 
			this.Finish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Finish.Image = global::EXLSXS.Properties.Resources.RibbonFinish_32x32;
			this.Finish.Label = "全シートに適用";
			this.Finish.Name = "Finish";
			this.Finish.ScreenTip = "全シートに適用";
			this.Finish.ShowImage = true;
			this.Finish.SuperTip = "選択中の表示モード・倍率をすべてのシートに適用し、各シート A1 を選択してから先頭のシートをアクティブにします。「フォントを揃える」「方眼紙にする」「表示形式を揃える」が ON のときは、それぞれの設定もすべてのシート・すべてのセルに適用します。";
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
			this.WindowViewBox.SuperTip = "すべてのシートをこの表示モード(標準 / ページ レイアウト / 改ページ プレビュー)に揃えます。";
			this.WindowViewBox.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WindowViewBox_SelectionChanged);
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
			this.WindowZoomBox.SuperTip = "すべてのシートの表示倍率をこの値(50%〜200%)に揃えます。";
			this.WindowZoomBox.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WindowZoomBox_SelectionChanged);
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
			// separator2
			//
			this.separator2.Name = "separator2";
			//
			// AdjustGridBox
			//
			this.AdjustGridBox.Label = "方眼紙にする";
			this.AdjustGridBox.Name = "AdjustGridBox";
			this.AdjustGridBox.ScreenTip = "方眼紙にする";
			this.AdjustGridBox.SuperTip = "ブックのすべてのシートのすべてのセルの行の高さ・列の幅を、右で選択したサイズの正方形に揃えます。";
			this.AdjustGridBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AdjustGridBox_Click);
			//
			// GridSizeBox
			//
			this.GridSizeBox.Label = "マスのサイズ";
			this.GridSizeBox.Name = "GridSizeBox";
			this.GridSizeBox.ScreenTip = "マスのサイズ";
			this.GridSizeBox.ShowLabel = false;
			this.GridSizeBox.SizeString = "改ページ プレビュー";
			this.GridSizeBox.SuperTip = "「方眼紙にする」が ON のとき、すべてのシートのセルをこのサイズの正方形に揃えます。";
			//
			// AdjustNumberFormatBox
			//
			this.AdjustNumberFormatBox.Label = "表示形式を揃える";
			this.AdjustNumberFormatBox.Name = "AdjustNumberFormatBox";
			this.AdjustNumberFormatBox.ScreenTip = "表示形式を揃える";
			this.AdjustNumberFormatBox.SuperTip = "ブックのすべてのシートのすべてのセルの表示形式を、右で選択した形式に揃えます。";
			this.AdjustNumberFormatBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AdjustNumberFormatBox_Click);
			//
			// NumberFormatBox
			//
			this.NumberFormatBox.Label = "表示形式";
			this.NumberFormatBox.Name = "NumberFormatBox";
			this.NumberFormatBox.ScreenTip = "表示形式";
			this.NumberFormatBox.ShowLabel = false;
			this.NumberFormatBox.SizeString = "改ページ プレビュー";
			this.NumberFormatBox.SuperTip = "「表示形式を揃える」が ON のとき、すべてのシートのセルの表示形式をこの形式に揃えます。";
			//
			// GridRowBox
			//
			this.GridRowBox.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Horizontal;
			this.GridRowBox.Items.Add(this.AdjustGridBox);
			this.GridRowBox.Items.Add(this.GridSizeBox);
			this.GridRowBox.Name = "GridRowBox";
			//
			// NumberFormatRowBox
			//
			this.NumberFormatRowBox.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Horizontal;
			this.NumberFormatRowBox.Items.Add(this.AdjustNumberFormatBox);
			this.NumberFormatRowBox.Items.Add(this.NumberFormatBox);
			this.NumberFormatRowBox.Name = "NumberFormatRowBox";
			//
			// FormatStackBox
			//
			this.FormatStackBox.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
			this.FormatStackBox.Items.Add(this.GridRowBox);
			this.FormatStackBox.Items.Add(this.NumberFormatRowBox);
			this.FormatStackBox.Items.Add(this.LocalBuildLabel);
			this.FormatStackBox.Name = "FormatStackBox";
			//
			// LocalBuildLabel
			//
			this.LocalBuildLabel.Label = "⚠ ローカル実行版";
			this.LocalBuildLabel.Name = "LocalBuildLabel";
			this.LocalBuildLabel.ScreenTip = "ローカル実行版";
			this.LocalBuildLabel.SuperTip = "この EXLSXS はローカルでビルドされたバージョンです。Cloudflare R2 で配信されている正式なインストール版ではありません。";
			this.LocalBuildLabel.Visible = false;
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
			this.GridRowBox.ResumeLayout(false);
			this.GridRowBox.PerformLayout();
			this.NumberFormatRowBox.ResumeLayout(false);
			this.NumberFormatRowBox.PerformLayout();
			this.FormatStackBox.ResumeLayout(false);
			this.FormatStackBox.PerformLayout();
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
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox AdjustGridBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown GridSizeBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonBox GridRowBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox AdjustNumberFormatBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown NumberFormatBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonBox NumberFormatRowBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonBox FormatStackBox;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel LocalBuildLabel;
	}

	partial class ThisRibbonCollection
	{
		internal MyRibbon MyRibbon
		{
			get { return this.GetRibbon<MyRibbon>(); }
		}
	}
}
