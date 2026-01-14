using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace EXLSXS
{
    /// <summary>
    /// Excel COM アドイン - .NET 10 実装
    /// </summary>
    [ComVisible(true)]
    [Guid("87654321-4321-4321-4321-210987654321")]
    public sealed class ThisAddIn : ExcelAddInBase
    {
        private readonly Dictionary<Workbook, AppEventListener> _workbookEventListeners =
            new(new ReferenceEqualityComparer<Workbook>());
        private string _assemblyTitle;
        private string _assemblyProduct;
        private string _publishVersion;
        private string _assemblyCopyright;

        public ThisAddIn()
        {
            Globals.ThisAddIn = this;
            SetupAddInInfoProperty();
        }

        /// <summary>アセンブリのメタデータを取得</summary>
        private void SetupAddInInfoProperty()
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            var attributes = executingAssembly.GetCustomAttributes(false);

            foreach (var attr in attributes)
            {
                if (attr is AssemblyTitleAttribute titleAttr)
                    _assemblyTitle = titleAttr.Title;
                else if (attr is AssemblyProductAttribute productAttr)
                    _assemblyProduct = productAttr.Product;
                else if (attr is AssemblyCopyrightAttribute copyrightAttr)
                    _assemblyCopyright = copyrightAttr.Copyright;
            }

            _publishVersion = "1.0.0.0"; // .NET 10 では ApplicationDeployment は使用不可
        }

        /// <summary>公開プロパティ</summary>
        public string AssemblyTitle => _assemblyTitle ?? string.Empty;
        public string AssemblyProduct => _assemblyProduct ?? string.Empty;
        public string PublishVersion => _publishVersion ?? "Unknown";
        public string AssemblyCopyright => _assemblyCopyright ?? string.Empty;

        /// <summary>OnConnected: アドイン初期化</summary>
        protected override void OnConnected()
        {
            try
            {
                // Excel イベントのリッスンを開始
                SetupExcelEventHandlers();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnConnected Error: {ex}");
            }
        }

        /// <summary>Excel アプリケーションのイベントをセットアップ</summary>
        private void SetupExcelEventHandlers()
        {
            if (ExcelApplication == null)
                return;

            try
            {
                // ワークブックのイベントをリッスン
                ExcelApplication.WorkbookOpen += ExcelApplication_WorkbookOpen;
                ExcelApplication.WorkbookActivate += ExcelApplication_WorkbookActivate;
                ExcelApplication.WorkbookDeactivate += ExcelApplication_WorkbookDeactivate;

                // 現在開いているワークブックにイベントリスナーを登録
                foreach (Workbook workbook in ExcelApplication.Workbooks)
                {
                    RegisterWorkbookEvents(workbook);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SetupExcelEventHandlers Error: {ex}");
            }
        }

        /// <summary>ワークブックのイベントをリッスン</summary>
        private void RegisterWorkbookEvents(Workbook workbook)
        {
            if (workbook == null)
                return;

            try
            {
                if (_workbookEventListeners.ContainsKey(workbook))
                    return;

                var listener = new AppEventListener(workbook, RemoveWorkbookEvents);
                _workbookEventListeners.Add(workbook, listener);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RegisterWorkbookEvents Error: {ex}");
            }
        }

        private void RemoveWorkbookEvents(Workbook workbook)
        {
            if (workbook == null)
                return;

            try
            {
                _workbookEventListeners.Remove(workbook);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RemoveWorkbookEvents Error: {ex}");
            }
        }

        /// <summary>ワークブック開く時のイベント</summary>
        private void ExcelApplication_WorkbookOpen(Workbook Wb)
        {
            RegisterWorkbookEvents(Wb);
        }

        /// <summary>ワークブックがアクティブになったときのイベント</summary>
        private void ExcelApplication_WorkbookActivate(Workbook Wb)
        {
            // リボン UI更新通知（オプション）
        }

        /// <summary>ワークブックが非アクティブになったときのイベント</summary>
        private void ExcelApplication_WorkbookDeactivate(Workbook Wb)
        {
            // リボン UI更新通知（オプション）
        }

        /// <summary>OnDisconnected: クリーンアップ</summary>
        protected override void OnDisconnected()
        {
            try
            {
                if (ExcelApplication != null)
                {
                    ExcelApplication.WorkbookOpen -= ExcelApplication_WorkbookOpen;
                    ExcelApplication.WorkbookActivate -= ExcelApplication_WorkbookActivate;
                    ExcelApplication.WorkbookDeactivate -= ExcelApplication_WorkbookDeactivate;
                }

                // イベントリスナーのクリーンアップ
                _workbookEventListeners?.Clear();
                Globals.ThisAddIn = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnDisconnected Error: {ex}");
            }
        }

        /// <summary>ワークシート仕上げ処理を実行</summary>
        public void DoFinish()
        {
            if (ExcelApplication?.ActiveWorkbook == null)
                return;

            try
            {
                Workbook activeWorkbook = ExcelApplication.ActiveWorkbook;
                int count = activeWorkbook.Worksheets.Count;

                for (int i = count; i > 0; i--)
                {
                    try
                    {
                        object obj = activeWorkbook.Worksheets[i];
                        if (obj is Worksheet worksheet)
                        {
                            worksheet.Activate();

                            // ビューモードを標準に設定
                            if (ExcelApplication.ActiveWindow != null)
                            {
                                ExcelApplication.ActiveWindow.View = XlWindowView.xlNormalView;
                                ExcelApplication.ActiveWindow.Zoom = 100;
                            }

                            // フォント設定（オプション）
                            // worksheet.Cells.Font.Name = "Calibri";

                            // A1 セルを選択
                            worksheet.Range["A1"].Select();

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"DoFinish worksheet {i} Error: {ex}");
                    }
                }

                // 最初のシートをアクティブにする
                if (count > 0)
                {
                    try
                    {
                        object firstSheet = activeWorkbook.Worksheets[1];
                        if (firstSheet is Worksheet ws)
                        {
                            ws.Activate();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"DoFinish activate first sheet Error: {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DoFinish Error: {ex}");
            }
        }
    }

    /// <summary>
    /// ワークブック イベント リスナー
    /// </summary>
    internal class AppEventListener
    {
        private readonly Workbook _workbook;
        private readonly Action<Workbook> _onWorkbookClosed;

        public AppEventListener(Workbook workbook, Action<Workbook> onWorkbookClosed)
        {
            _workbook = workbook;
            _onWorkbookClosed = onWorkbookClosed;
            try
            {
                // ワークブック イベントの登録
                _workbook.BeforeSave += Workbook_BeforeSave;
                _workbook.BeforeClose += Workbook_BeforeClose;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppEventListener init Error: {ex}");
            }
        }

        private void Workbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            // 保存前の処理（オプション）
        }

        private void Workbook_BeforeClose(ref bool Cancel)
        {
            // 閉じる前のクリーンアップ
            try
            {
                _workbook.BeforeSave -= Workbook_BeforeSave;
                _workbook.BeforeClose -= Workbook_BeforeClose;
                _onWorkbookClosed?.Invoke(_workbook);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Workbook_BeforeClose Error: {ex}");
            }
        }
    }

    internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T>
        where T : class
    {
        public bool Equals(T x, T y) => ReferenceEquals(x, y);

        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
