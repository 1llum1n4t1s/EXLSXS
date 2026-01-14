using System;
using System.Windows.Forms;

namespace Tk4.ExcelAddIns.Finisher
{
    /// <summary>
    /// UI ダイアログ ヘルパー クラス
    /// </summary>
    public static class UIHelper
    {
        private static string GetTitle()
        {
            return Globals.ThisAddIn?.AssemblyTitle ?? "Excel Finisher Add-in";
        }

        /// <summary>確認ダイアログを表示</summary>
        public static bool ShowConfirmDialog(bool showWarningIcon, string message, params string[] args)
        {
            try
            {
                string formattedMessage = string.Format(message, args);
                MessageBoxIcon icon = showWarningIcon ? MessageBoxIcon.Exclamation : MessageBoxIcon.Question;
                DialogResult result = MessageBox.Show(formattedMessage, GetTitle(), MessageBoxButtons.YesNo, icon);
                return result == DialogResult.Yes;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowConfirmDialog Error: {ex}");
                return false;
            }
        }

        /// <summary>情報ダイアログを表示</summary>
        public static void ShowInformationDialog(string message, params string[] args)
        {
            try
            {
                string formattedMessage = string.Format(message, args);
                MessageBox.Show(formattedMessage, GetTitle(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowInformationDialog Error: {ex}");
            }
        }

        /// <summary>警告ダイアログを表示</summary>
        public static void ShowWarningDialog(string message, params string[] args)
        {
            try
            {
                string formattedMessage = string.Format(message, args);
                MessageBox.Show(formattedMessage, GetTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowWarningDialog Error: {ex}");
            }
        }

        /// <summary>エラー ダイアログを表示</summary>
        public static void ShowErrorDialog(Exception e)
        {
            try
            {
                string message = e?.Message ?? "Unknown error occurred";
                MessageBox.Show(message, GetTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowErrorDialog Error: {ex}");
            }
        }
    }
}
