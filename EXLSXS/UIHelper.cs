using System;
using System.Windows.Forms;

namespace EXLSXS
{
	public static class UIHelper
	{
		public static bool ShowConfirmDialog(bool showWarningIcon, string message, params string[] args)
		{
			DialogResult dialogResult = MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.YesNo, showWarningIcon ? MessageBoxIcon.Exclamation : MessageBoxIcon.Question);
			return dialogResult == DialogResult.Yes;
		}

		public static void ShowInfomationDialog(string message, params string[] args)
		{
			MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}

		public static void ShowWarningDialog(string message, params string[] args)
		{
			MessageBox.Show(string.Format(message, args), Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		public static void ShowErrorDialog(Exception e)
		{
			MessageBox.Show(e.Message, Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}
}
