using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;

namespace EXLSXS
{
    /// <summary>
    /// Excel DNA アドイン - .NET 10 実装
    /// </summary>
    public class ThisAddIn : IExcelAddIn
    {
        private ExcelApplication _excelApplication;
        private Dictionary<Workbook, AppEventListener> _workbookEventListeners =
            new(new ReferenceEqualityComparer<Workbook>());
        private string _assemblyTitle;
        private string _assemblyProduct;
        private string _publishVersion;
        private string _assemblyCopyright;

        /// <summary>アドイン初期化</summary>
        public void AutoOpen()
        {
            try
            {
#pragma warning disable CA1416
                _excelApplication = ExcelDnaUtil.Application as ExcelApplication;
#pragma warning restore CA1416
                Globals.ThisAddIn = this;
                SetupAddInInfoProperty();
                SetupExcelEventHandlers();
                System.Diagnostics.Debug.WriteLine("EXLSXS AddIn loaded successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoOpen Error: {ex}");
            }
        }

        /// <summary>アドイン終了処理</summary>
        public void AutoClose()
        {
            try
            {
                if (_excelApplication != null)
                {
                    _excelApplication.WorkbookOpen -= ExcelApplication_WorkbookOpen;
                    _excelApplication.WorkbookActivate -= ExcelApplication_WorkbookActivate;
                    _excelApplication.WorkbookDeactivate -= ExcelApplication_WorkbookDeactivate;
                }

                _workbookEventListeners?.Clear();
                Globals.ThisAddIn = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoClose Error: {ex}");
            }
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

            _publishVersion = "1.0.0.0";
        }

        /// <summary>公開プロパティ</summary>
        public string AssemblyTitle => _assemblyTitle ?? string.Empty;
        public string AssemblyProduct => _assemblyProduct ?? string.Empty;
        public string PublishVersion => _publishVersion ?? "Unknown";
        public string AssemblyCopyright => _assemblyCopyright ?? string.Empty;

        /// <summary>Excel アプリケーションのイベントをセットアップ</summary>
        private void SetupExcelEventHandlers()
        {
            if (_excelApplication == null)
                return;

            try
            {
                _excelApplication.WorkbookOpen += ExcelApplication_WorkbookOpen;
                _excelApplication.WorkbookActivate += ExcelApplication_WorkbookActivate;
                _excelApplication.WorkbookDeactivate += ExcelApplication_WorkbookDeactivate;

                foreach (Workbook workbook in _excelApplication.Workbooks)
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
        }

        /// <summary>ワークブックが非アクティブになったときのイベント</summary>
        private void ExcelApplication_WorkbookDeactivate(Workbook Wb)
        {
        }

        /// <summary>ワークシート仕上げ処理を実行</summary>
        public void DoFinish()
        {
            if (_excelApplication?.ActiveWorkbook == null)
                return;

            try
            {
                Workbook activeWorkbook = _excelApplication.ActiveWorkbook;
                int count = activeWorkbook.Worksheets.Count;

                for (int i = count; i > 0; i--)
                {
                    try
                    {
                        object obj = activeWorkbook.Worksheets[i];
                        if (obj is Worksheet worksheet)
                        {
                            worksheet.Activate();

                            if (_excelApplication.ActiveWindow != null)
                            {
                                _excelApplication.ActiveWindow.View = XlWindowView.xlNormalView;
                                _excelApplication.ActiveWindow.Zoom = 100;
                            }

                            worksheet.Range["A1"].Select();

#pragma warning disable CA1416
                            Marshal.ReleaseComObject(worksheet);
#pragma warning restore CA1416
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"DoFinish worksheet {i} Error: {ex}");
                    }
                }

                if (count > 0)
                {
                    try
                    {
                        object firstSheet = activeWorkbook.Worksheets[1];
                        if (firstSheet is Worksheet ws)
                        {
                            ws.Activate();
#pragma warning disable CA1416
                            Marshal.ReleaseComObject(ws);
#pragma warning restore CA1416
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
        }

        private void Workbook_BeforeClose(ref bool Cancel)
        {
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

    /// <summary>
    /// リボンボタンコールバック - ExcelDNA
    /// </summary>
    public class ExcelAddInButtons
    {
        /// <summary>仕上げボタンクリック処理</summary>
        public void ExcelAddInButtonCallback(object control)
        {
            try
            {
                Globals.ThisAddIn?.DoFinish();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelAddInButtonCallback Error: {ex}");
            }
        }
    }

    /// <summary>
    /// グローバル設定クラス
    /// </summary>
    internal sealed class Globals
    {
        private Globals()
        {
        }

        internal static ThisAddIn ThisAddIn { get; set; }
    }

    internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T>
        where T : class
    {
        public bool Equals(T x, T y) => ReferenceEquals(x, y);

        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
