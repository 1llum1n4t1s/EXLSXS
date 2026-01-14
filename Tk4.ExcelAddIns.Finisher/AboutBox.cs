using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Tk4.ExcelAddIns.Finisher
{
    /// <summary>
    /// About ダイアログ (.NET 10 対応)
    /// </summary>
    public class AboutBox : Form
    {
        private TableLayoutPanel tableLayoutPanel;
        private Label productNameLabel;
        private Label versionLabel;
        private Label copyrightLabel;
        private LinkLabel homePageLink;
        private Button okButton;

        public AboutBox()
        {
            InitializeComponent();
            SetupLabels();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 250);
            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // TableLayoutPanel
            tableLayoutPanel = new TableLayoutPanel();
            tableLayoutPanel.ColumnCount = 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.RowCount = 5;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            tableLayoutPanel.Padding = new Padding(10);

            // Labels
            productNameLabel = new Label { AutoSize = true, Dock = DockStyle.Top, Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold) };
            versionLabel = new Label { AutoSize = true, Dock = DockStyle.Top };
            copyrightLabel = new Label { AutoSize = true, Dock = DockStyle.Top };

            tableLayoutPanel.Controls.Add(productNameLabel, 0, 0);
            tableLayoutPanel.Controls.Add(versionLabel, 0, 1);
            tableLayoutPanel.Controls.Add(copyrightLabel, 0, 2);

            // Home Page Link
            homePageLink = new LinkLabel { AutoSize = true, Text = "https://example.com", Dock = DockStyle.Top };
            homePageLink.LinkClicked += HomePageLink_LinkClicked;
            tableLayoutPanel.Controls.Add(homePageLink, 0, 3);

            // OK Button
            okButton = new Button { DialogResult = DialogResult.OK, Text = "OK" };
            okButton.Click += OKButton_Click;
            tableLayoutPanel.Controls.Add(okButton, 0, 4);
            okButton.Dock = DockStyle.Right;
            okButton.Width = 80;

            this.Controls.Add(tableLayoutPanel);
            this.ResumeLayout(false);
        }

        private void SetupLabels()
        {
            if (Globals.ThisAddIn != null)
            {
                this.Text = $"{Globals.ThisAddIn.AssemblyTitle} について";
                productNameLabel.Text = Globals.ThisAddIn.AssemblyProduct ?? "Excel Finisher Add-in";
                versionLabel.Text = $"Version {Globals.ThisAddIn.PublishVersion}";
                copyrightLabel.Text = $"{Globals.ThisAddIn.AssemblyCopyright} All rights reserved.";
            }
            else
            {
                this.Text = "About";
                productNameLabel.Text = "Excel Finisher Add-in";
                versionLabel.Text = "Version 1.0.0.0";
                copyrightLabel.Text = "Copyright © 2024";
            }
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void HomePageLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(homePageLink.Text) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HomePageLink_LinkClicked Error: {ex}");
            }
        }
    }
}
