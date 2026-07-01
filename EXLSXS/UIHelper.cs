using System;
using System.Windows.Forms;

namespace EXLSXS
{
	public static class UIHelper
	{
		public static void ShowErrorDialog(Exception e)
		{
			MessageBox.Show(e.Message, Globals.ThisAddIn.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}
}
