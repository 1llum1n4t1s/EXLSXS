namespace Tk4.ExcelAddIns.Finisher
{
    /// <summary>
    /// Excel VSTO アドインのエントリポイント。
    /// </summary>
    public class ThisAddIn
    {
        /// <summary>
        /// 起動時に呼び出される想定のイベントハンドラ。
        /// </summary>
        public void ThisAddIn_Startup()
        {
        }

        /// <summary>
        /// 終了時に呼び出される想定のイベントハンドラ。
        /// </summary>
        public void ThisAddIn_Shutdown()
        {
        }

        /// <summary>
        /// Ribbon ファクトリ取得の想定メソッド。
        /// </summary>
        public object GetRibbonFactory()
        {
            return null;
        }

        /// <summary>
        /// リボンコレクション取得の想定プロパティ。
        /// </summary>
        public object Ribbons
        {
            get { return null; }
        }
    }
}
