using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using EXLSXS.Properties;

namespace EXLSXS
{
	// Token: 0x02000002 RID: 2
	public partial class AboutBox : Form
	{
		// Token: 0x06000001 RID: 1 RVA: 0x000020D0 File Offset: 0x000002D0
		public AboutBox()
		{
			this.InitializeComponent();
			this.Text = string.Format("{0}のバージョン情報", Globals.ThisAddIn.AssemblyTitle);
			this.ProductNameLabel.Text = Globals.ThisAddIn.AssemblyProduct;
			this.VersionLabel.Text = string.Format("Version {0}", Globals.ThisAddIn.PublishVersion);
			this.CopyrightLabel.Text = string.Format("{0} All rights reserved.", Globals.ThisAddIn.AssemblyCopyright);
		}

		// Token: 0x06000002 RID: 2 RVA: 0x00002156 File Offset: 0x00000356
		private void OKButton_Click(object sender, EventArgs e)
		{
			base.Close();
		}

		// Token: 0x06000003 RID: 3 RVA: 0x00002160 File Offset: 0x00000360
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
