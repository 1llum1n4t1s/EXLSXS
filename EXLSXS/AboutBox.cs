using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using EXLSXS.Properties;

namespace EXLSXS
{
	public partial class AboutBox : Form
	{
		public AboutBox()
		{
			this.InitializeComponent();
			this.Text = $"{Globals.ThisAddIn.AssemblyTitle}のバージョン情報";
			this.ProductNameLabel.Text = Globals.ThisAddIn.AssemblyProduct;
			this.VersionLabel.Text = $"Version {Globals.ThisAddIn.PublishVersion}";
			this.CopyrightLabel.Text = $"{Globals.ThisAddIn.AssemblyCopyright} All rights reserved.";
		}

		private void OKButton_Click(object sender, EventArgs e)
		{
			base.Close();
		}

		private void HomePageLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			try
			{
				Process.Start(this.HomePageLink.Text);
			}
			catch
			{
			}
		}
	}
}
