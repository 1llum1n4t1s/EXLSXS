using System;
using System.Diagnostics;

namespace Tk4.ExcelAddIns.Finisher
{
    /// <summary>
    /// グローバル設定クラス (.NET 10 COM 相互運用対応)
    /// </summary>
    internal sealed class Globals
    {
        private Globals()
        {
        }

        /// <summary>現在のアドイン インスタンス</summary>
        internal static ThisAddIn ThisAddIn { get; set; }
    }
}
