using System;
using System.Diagnostics;
using System.Windows;

namespace Tk4.ExcelAddIns.Finisher
{
    /// <summary>
    /// About ダイアログ - WPF 実装 (.NET 10)
    /// </summary>
    public partial class AboutBox : Window
    {
        public AboutBox()
        {
            InitializeComponent();
            SetupLabels();
        }

        /// <summary>ラベルをセットアップ</summary>
        private void SetupLabels()
        {
            if (Globals.ThisAddIn != null)
            {
                this.Title = $"{Globals.ThisAddIn.AssemblyTitle} について";
                ProductNameLabel.Text = Globals.ThisAddIn.AssemblyProduct ?? "Excel Finisher Add-in";
                VersionLabel.Text = $"Version {Globals.ThisAddIn.PublishVersion}";
                CopyrightLabel.Text = $"{Globals.ThisAddIn.AssemblyCopyright} All rights reserved.";
            }
            else
            {
                this.Title = "About";
                ProductNameLabel.Text = "Excel Finisher Add-in";
                VersionLabel.Text = "Version 1.0.0.0";
                CopyrightLabel.Text = "Copyright © 2024. All rights reserved.";
            }
        }

        /// <summary>OK ボタンクリック</summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>ホームページリンククリック</summary>
        private void HomePageLink_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var hyperlink = sender as System.Windows.Documents.Hyperlink;
                if (hyperlink?.Inlines.FirstInline is System.Windows.Documents.Run run)
                {
                    Process.Start(new ProcessStartInfo(run.Text) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HomePageLink_Click Error: {ex}");
            }
        }
    }
}
