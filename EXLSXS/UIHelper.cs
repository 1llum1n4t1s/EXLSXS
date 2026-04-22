using System;
using System.Windows.Forms;

namespace EXLSXS
{
	// Token: 0x02000009 RID: 9
	public static class UIHelper
	{
		// Token: 0x06000048 RID: 72 RVA: 0x00003780 File Offset: 0x00001980
		public static bool ShowConfirmDialog(bool showWarningIcon, string message, params string[] args)
		{
			DialogResult dialogResult = MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.YesNo, showWarningIcon ? MessageBoxIcon.Exclamation : MessageBoxIcon.Question);
			return dialogResult == DialogResult.Yes;
		}

		// Token: 0x06000049 RID: 73 RVA: 0x000037B2 File Offset: 0x000019B2
		public static void ShowInfomationDialog(string message, params string[] args)
		{
			MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}

		// Token: 0x0600004A RID: 74 RVA: 0x000037CE File Offset: 0x000019CE
		public static void ShowWarningDialog(string message, params string[] args)
		{
			MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		// Token: 0x0600004B RID: 75 RVA: 0x000037EA File Offset: 0x000019EA
		public static void ShowErrorDialog(Exception e)
		{
			MessageBox.Show(e.Message, Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}
}
