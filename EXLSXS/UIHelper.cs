using System;
using System.Windows;

namespace EXLSXS
{
    /// <summary>
    /// UI ダイアログ ヘルパー クラス - WPF 実装
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
                MessageBoxImage icon = showWarningIcon ? MessageBoxImage.Exclamation : MessageBoxImage.Question;
                MessageBoxResult result = MessageBox.Show(
                    formattedMessage,
                    GetTitle(),
                    MessageBoxButton.YesNo,
                    icon);
                return result == MessageBoxResult.Yes;
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
                MessageBox.Show(
                    formattedMessage,
                    GetTitle(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
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
                MessageBox.Show(
                    formattedMessage,
                    GetTitle(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Exclamation);
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
                MessageBox.Show(
                    message,
                    GetTitle(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowErrorDialog Error: {ex}");
            }
        }
    }
}
