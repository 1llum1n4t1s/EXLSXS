namespace Tk4.ExcelAddIns.Finisher
{
	// Token: 0x02000002 RID: 2
	public partial class AboutBox : global::System.Windows.Forms.Form
	{
		// Token: 0x06000004 RID: 4 RVA: 0x00002194 File Offset: 0x00000394
		protected override void Dispose(bool disposing)
		{
			if (disposing && this.components != null)
			{
				this.components.Dispose();
			}
			base.Dispose(disposing);
		}

		// Token: 0x06000005 RID: 5 RVA: 0x000021B4 File Offset: 0x000003B4
		private void InitializeComponent()
		{
			this.OKButton = new global::System.Windows.Forms.Button();
			this.pictureBox1 = new global::System.Windows.Forms.PictureBox();
			this.ProductNameLabel = new global::System.Windows.Forms.Label();
			this.VersionLabel = new global::System.Windows.Forms.Label();
			this.CopyrightLabel = new global::System.Windows.Forms.Label();
			this.label1 = new global::System.Windows.Forms.Label();
			this.HomePageLink = new global::System.Windows.Forms.LinkLabel();
			this.label5 = new global::System.Windows.Forms.Label();
			((global::System.ComponentModel.ISupportInitialize)this.pictureBox1).BeginInit();
			base.SuspendLayout();
			this.OKButton.Anchor = (global::System.Windows.Forms.AnchorStyles.Bottom | global::System.Windows.Forms.AnchorStyles.Right);
			this.OKButton.AutoSize = true;
			this.OKButton.Location = new global::System.Drawing.Point(353, 197);
			this.OKButton.Margin = new global::System.Windows.Forms.Padding(3, 4, 3, 4);
			this.OKButton.Name = "OKButton";
			this.OKButton.Size = new global::System.Drawing.Size(75, 28);
			this.OKButton.TabIndex = 0;
			this.OKButton.Text = "OK";
			this.OKButton.UseVisualStyleBackColor = true;
			this.OKButton.Click += new global::System.EventHandler(this.OKButton_Click);
			this.pictureBox1.Image = global::Tk4.ExcelAddIns.Finisher.Properties.Resources.ExcelAddIn_64x64;
			this.pictureBox1.Location = new global::System.Drawing.Point(12, 12);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new global::System.Drawing.Size(64, 64);
			this.pictureBox1.TabIndex = 1;
			this.pictureBox1.TabStop = false;
			this.ProductNameLabel.AutoSize = true;
			this.ProductNameLabel.Location = new global::System.Drawing.Point(82, 18);
			this.ProductNameLabel.Name = "ProductNameLabel";
			this.ProductNameLabel.Size = new global::System.Drawing.Size(87, 18);
			this.ProductNameLabel.TabIndex = 2;
			this.ProductNameLabel.Text = "ProductName";
			this.VersionLabel.AutoSize = true;
			this.VersionLabel.Location = new global::System.Drawing.Point(82, 38);
			this.VersionLabel.Name = "VersionLabel";
			this.VersionLabel.Size = new global::System.Drawing.Size(50, 18);
			this.VersionLabel.TabIndex = 3;
			this.VersionLabel.Text = "Version";
			this.CopyrightLabel.AutoSize = true;
			this.CopyrightLabel.Location = new global::System.Drawing.Point(82, 58);
			this.CopyrightLabel.Name = "CopyrightLabel";
			this.CopyrightLabel.Size = new global::System.Drawing.Size(64, 18);
			this.CopyrightLabel.TabIndex = 4;
			this.CopyrightLabel.Text = "Copyright";
			this.label1.AutoSize = true;
			this.label1.Location = new global::System.Drawing.Point(12, 89);
			this.label1.Name = "label1";
			this.label1.Size = new global::System.Drawing.Size(121, 18);
			this.label1.TabIndex = 5;
			this.label1.Text = "製品のホームページ:";
			this.HomePageLink.AutoSize = true;
			this.HomePageLink.Location = new global::System.Drawing.Point(12, 110);
			this.HomePageLink.Name = "HomePageLink";
			this.HomePageLink.Size = new global::System.Drawing.Size(271, 18);
			this.HomePageLink.TabIndex = 6;
			this.HomePageLink.TabStop = true;
			this.HomePageLink.Text = "http://www.tk4.co.jp/product/excel_finisher/";
			this.HomePageLink.LinkClicked += new global::System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.HomePageLink_LinkClicked);
			this.label5.Anchor = (global::System.Windows.Forms.AnchorStyles.Bottom | global::System.Windows.Forms.AnchorStyles.Left | global::System.Windows.Forms.AnchorStyles.Right);
			this.label5.Location = new global::System.Drawing.Point(12, 156);
			this.label5.Name = "label5";
			this.label5.Size = new global::System.Drawing.Size(416, 38);
			this.label5.TabIndex = 7;
			this.label5.Text = "本ソフトに関連して直接又は間接に被ったいかなる損害についても、TKForce Inc. は一切責任を負いません。";
			base.AcceptButton = this.OKButton;
			base.AutoScaleDimensions = new global::System.Drawing.SizeF(7f, 18f);
			base.AutoScaleMode = global::System.Windows.Forms.AutoScaleMode.Font;
			base.ClientSize = new global::System.Drawing.Size(440, 238);
			base.Controls.Add(this.label5);
			base.Controls.Add(this.HomePageLink);
			base.Controls.Add(this.label1);
			base.Controls.Add(this.CopyrightLabel);
			base.Controls.Add(this.VersionLabel);
			base.Controls.Add(this.ProductNameLabel);
			base.Controls.Add(this.pictureBox1);
			base.Controls.Add(this.OKButton);
			this.Font = new global::System.Drawing.Font("メイリオ", 9f, global::System.Drawing.FontStyle.Regular, global::System.Drawing.GraphicsUnit.Point, 128);
			base.FormBorderStyle = global::System.Windows.Forms.FormBorderStyle.FixedDialog;
			base.Margin = new global::System.Windows.Forms.Padding(3, 4, 3, 4);
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.Name = "AboutBox";
			base.ShowIcon = false;
			base.ShowInTaskbar = false;
			base.SizeGripStyle = global::System.Windows.Forms.SizeGripStyle.Hide;
			base.StartPosition = global::System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "AboutBox";
			((global::System.ComponentModel.ISupportInitialize)this.pictureBox1).EndInit();
			base.ResumeLayout(false);
			base.PerformLayout();
		}

		// Token: 0x04000001 RID: 1
		private global::System.ComponentModel.IContainer components;

		// Token: 0x04000002 RID: 2
		private global::System.Windows.Forms.Button OKButton;

		// Token: 0x04000003 RID: 3
		private global::System.Windows.Forms.PictureBox pictureBox1;

		// Token: 0x04000004 RID: 4
		private global::System.Windows.Forms.Label ProductNameLabel;

		// Token: 0x04000005 RID: 5
		private global::System.Windows.Forms.Label VersionLabel;

		// Token: 0x04000006 RID: 6
		private global::System.Windows.Forms.Label CopyrightLabel;

		// Token: 0x04000007 RID: 7
		private global::System.Windows.Forms.Label label1;

		// Token: 0x04000008 RID: 8
		private global::System.Windows.Forms.LinkLabel HomePageLink;

		// Token: 0x04000009 RID: 9
		private global::System.Windows.Forms.Label label5;
	}
}
