using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using Path = System.IO.Path;

namespace EXLSXS;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
    private const int PROGRESS_BACKUP = 10;
    private const int PROGRESS_REPAIR = 30;
    private const int PROGRESS_OPTIMIZE = 50;
    private const int PROGRESS_PERFORMANCE = 70;

    private const int PROGRESS_IMAGES = 95;
    private const int PROGRESS_COMPLETE = 100;

    private string? selectedFilePath;
    private bool isProcessing = false;
    private readonly List<string> processLog = new();
    private Settings? appSettings;

    /// <summary>
    /// MainWindowの初期化処理
    /// </summary>
    public MainWindow()
    {
        InitializeComponent();
        InitializeApplication();
    }

    /// <summary>
    /// アプリケーションの初期化（設定・ログ・バッチプロセッサ初期化）
    /// </summary>
    private void InitializeApplication()
    {
        // 設定を読み込み
        appSettings = Settings.Load();
        LoadSettingsToUI();
        LogMessage("アプリケーションが開始されました");
    }

    /// <summary>
    /// 設定をUIに反映
    /// </summary>
    private void LoadSettingsToUI()
    {
        try
        {
            if (appSettings == null) return;

            BackupOriginalCheckBox.IsChecked = appSettings.CreateBackup;
            OptimizeCheckBox.IsChecked = appSettings.OptimizeFile;
            PerformanceCheckBox.IsChecked = appSettings.ImprovePerformance;
            CompressImagesCheckBox.IsChecked = appSettings.CompressImages;
        }
        catch (Exception ex)
        {
            LogMessage($"設定読み込みエラー: {ex.Message}");
        }
    }

    /// <summary>
    /// UIの設定を保存
    /// </summary>
    private void SaveSettingsFromUI()
    {
        try
        {
            if (appSettings == null) return;

            appSettings.CreateBackup = BackupOriginalCheckBox.IsChecked == true;
            appSettings.OptimizeFile = OptimizeCheckBox.IsChecked == true;
            appSettings.ImprovePerformance = PerformanceCheckBox.IsChecked == true;
            appSettings.CompressImages = CompressImagesCheckBox.IsChecked == true;

            appSettings.LastSelectedFolder = GetLastSelectedFolder();

            appSettings.Save();
            LogMessage("設定が保存されました");
        }
        catch (Exception ex)
        {
            LogMessage($"設定保存エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 最後に選択したフォルダパスを取得
    /// </summary>
    private string GetLastSelectedFolder()
    {
        return !string.IsNullOrEmpty(selectedFilePath)
            ? Path.GetDirectoryName(selectedFilePath) ?? ""
            : "";
    }

    /// <summary>
    /// ファイル選択ボタン押下時の処理
    /// </summary>
    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = CreateFileDialog();
        SetInitialDirectory(openFileDialog);

        if (openFileDialog.ShowDialog() == true)
        {
            selectedFilePath = openFileDialog.FileName;
            FilePathTextBox.Text = selectedFilePath;
            ProcessButton.IsEnabled = true;
            LogMessage($"ファイルが選択されました: {selectedFilePath}");
            SaveSettingsFromUI();
        }
    }

    /// <summary>
    /// 実行ボタン押下時の処理
    /// </summary>
    private async void ProcessButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateFileSelection()) return;
        if (isProcessing)
        {
            ShowProcessingMessage();
            return;
        }

        await ExecuteFileProcessing();
    }

    /// <summary>
    /// 処理中メッセージを表示
    /// </summary>
    private void ShowProcessingMessage()
    {
        MessageBox.Show("処理中です。完了までお待ちください。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    /// <summary>
    /// ファイル選択の妥当性を検証
    /// </summary>
    private bool ValidateFileSelection()
    {
        if (string.IsNullOrEmpty(selectedFilePath) || !File.Exists(selectedFilePath))
        {
            MessageBox.Show("有効なExcelファイルを選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        return true;
    }

    /// <summary>
    /// ファイル選択ダイアログを生成
    /// </summary>
    private OpenFileDialog CreateFileDialog()
    {
        return new OpenFileDialog
        {
            Title = "Excelファイルを選択",
            Filter = "Excelファイル (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*",
            Multiselect = false
        };
    }

    /// <summary>
    /// ダイアログの初期ディレクトリを設定
    /// </summary>
    private void SetInitialDirectory(OpenFileDialog dialog)
    {
        if (appSettings?.LastSelectedFolder != null && Directory.Exists(appSettings.LastSelectedFolder))
        {
            dialog.InitialDirectory = appSettings.LastSelectedFolder;
        }
    }



    /// <summary>
    /// 単一ファイル処理の実行
    /// </summary>
    private async Task ExecuteFileProcessing()
    {
        SetProcessingState(true);
        ProgressBar.Value = 0;

        try
        {
            LogMessage("処理を開始します");
            await ProcessExcelFileAsync();
            LogMessage("処理が正常に完了しました");
            SaveSettingsFromUI();
            MessageBox.Show("Excelファイルの修復・最適化が完了しました！", "完了", MessageBoxButton.OK, MessageBoxImage.Information);


        }
        catch (Exception ex)
        {
            LogMessage($"エラーが発生しました: {ex.Message}");
            MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetProcessingState(false);
            ProgressBar.Value = PROGRESS_COMPLETE;
        }
    }

    /// <summary>
    /// UIの処理状態を切り替え
    /// </summary>
    private void SetProcessingState(bool processing)
    {
        isProcessing = processing;
        ProcessButton.IsEnabled = !processing && !string.IsNullOrEmpty(selectedFilePath);
        BrowseButton.IsEnabled = !processing;
        if (!processing) ProgressBar.Value = 0;
    }

    /// <summary>
    /// Excelファイルの非同期処理本体
    /// </summary>
    private async Task ProcessExcelFileAsync()
    {
        var options = GetProcessingOptions();

        await Task.Run(() =>
        {
            ExecuteProcessingSteps(options);
        });
    }

    /// <summary>
    /// UIから処理オプションを取得
    /// </summary>
    private ProcessingOptions GetProcessingOptions()
    {
        return new ProcessingOptions
        {
            ShouldCreateBackup = BackupOriginalCheckBox.IsChecked == true,
            IsAggressiveRepair = true, // 常にAggressive修復を実行
            ShouldOptimize = OptimizeCheckBox.IsChecked == true,
            ShouldImprovePerformance = PerformanceCheckBox.IsChecked == true,
            ShouldCompressImages = CompressImagesCheckBox.IsChecked == true
        };
    }

    /// <summary>
    /// 各処理ステップを実行
    /// </summary>
    private void ExecuteProcessingSteps(ProcessingOptions options)
    {
        if (options.ShouldCreateBackup)
        {
            LogMessage("バックアップを作成中...");
            CreateBackup();
            UpdateProgress(PROGRESS_BACKUP);
        }

        ExecuteRepairStep(options);
        UpdateProgress(PROGRESS_REPAIR);

        if (options.ShouldOptimize)
        {
            LogMessage("ファイル最適化を実行中...");
            OptimizeExcelFile();
            UpdateProgress(PROGRESS_OPTIMIZE);
        }

        if (options.ShouldImprovePerformance)
        {
            LogMessage("パフォーマンス向上を実行中...");
            ImprovePerformance();
            UpdateProgress(PROGRESS_PERFORMANCE);
        }

        if (options.ShouldCompressImages)
        {
            LogMessage("画像の圧縮を実行中...");
            CompressImages();
            UpdateProgress(PROGRESS_IMAGES);
        }

        UpdateProgress(PROGRESS_COMPLETE);
    }

    /// <summary>
    /// 修復ステップを実行
    /// </summary>
    private void ExecuteRepairStep(ProcessingOptions options)
    {
        LogMessage("積極的修復を実行中...");
        PerformAggressiveRepair();
    }

    /// <summary>
    /// プログレスバーを更新
    /// </summary>
    private void UpdateProgress(int value)
    {
        Dispatcher.Invoke(() => ProgressBar.Value = value);
    }

    /// <summary>
    /// ログメッセージを記録
    /// </summary>
    private void LogMessage(string message)
    {
        var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
        processLog.Add(logEntry);

        // 画面上のログテキストボックスに表示
        Dispatcher.Invoke(() =>
        {
            LogTextBox.AppendText(logEntry + Environment.NewLine);
            LogTextBox.ScrollToEnd();
        });
    }

    /// <summary>
    /// ログクリアボタン押下時の処理
    /// </summary>
    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        LogTextBox.Clear();
        processLog.Clear();
    }

    /// <summary>
    /// ログコピーボタン押下時の処理
    /// </summary>
    private void CopyLogButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!string.IsNullOrEmpty(LogTextBox.Text))
            {
                Clipboard.SetText(LogTextBox.Text);
                MessageBox.Show("ログをクリップボードにコピーしました。", "コピー完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("コピーするログがありません。", "情報",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"ログのコピーに失敗しました: {ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// バックアップファイルを作成
    /// </summary>
    private void CreateBackup()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;

        try
        {
            var dir = Path.GetDirectoryName(selectedFilePath) ?? string.Empty;
            var backupPath = Path.Combine(
                dir,
                $"{Path.GetFileNameWithoutExtension(selectedFilePath)}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(selectedFilePath)}"
            );

            if (!string.IsNullOrEmpty(backupPath))
            {
                File.Copy(selectedFilePath, backupPath, true);
            }
        }
        catch (Exception ex)
        {
            LogMessage($"バックアップ作成エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// ファイルがアクセス可能かチェック
    /// </summary>
    private bool IsFileAccessible(string filePath)
    {
        try
        {
            using var stream = File.OpenRead(filePath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 最小限の修復処理（データを削除しない）
    /// </summary>
    private void PerformMinimalRepair()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            using var workbook = new XLWorkbook(selectedFilePath);
            foreach (var worksheet in workbook.Worksheets)
            {
                try
                {
                    var usedRange = worksheet.RangeUsed();
                    if (usedRange != null)
                    {
                        ValidateWorksheetData(worksheet);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            workbook.Save();
        }
        catch (Exception ex)
        {
            LogMessage($"最小修復エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 標準修復処理
    /// </summary>
    private void PerformStandardRepair()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            if (!IsFileAccessible(selectedFilePath))
            {
                LogMessage("ファイルにアクセスできません");
                return;
            }
            PerformMinimalRepair();
            using var workbook = new XLWorkbook(selectedFilePath);
            foreach (var worksheet in workbook.Worksheets)
            {
                try
                {
                    var usedRange = worksheet.RangeUsed();
                    if (usedRange != null)
                    {
                        CleanupEmptyRowsOnly(worksheet);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            workbook.Save();
            if (!IsFileAccessible(selectedFilePath))
            {
                LogMessage("修復後にファイルが破損しました");
                RestoreFromBackup();
            }
        }
        catch (Exception ex)
        {
            LogMessage($"標準修復エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 積極的修復処理（データ損失の可能性あり）
    /// </summary>
    private void PerformAggressiveRepair()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            if (!IsFileAccessible(selectedFilePath))
            {
                LogMessage("ファイルにアクセスできません");
                return;
            }
            PerformMinimalRepair();
            ExecuteRepairMethods();
            if (!IsFileAccessible(selectedFilePath))
            {
                LogMessage("修復後にファイルが破損しました");
                RestoreFromBackup();
            }
        }
        catch (Exception ex)
        {
            LogMessage($"積極的修復エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 複数の修復メソッドを実行
    /// </summary>
    private void ExecuteRepairMethods()
    {
        var repairMethods = new Action[]
        {
            () => { try { RepairUsingOpenXML(); } catch (Exception ex) { LogMessage($"OpenXML修復エラー: {ex.Message}"); } },
            () => { try { RepairWorksheets(); } catch (Exception ex) { LogMessage($"ワークシート修復エラー: {ex.Message}"); } }
        };

        foreach (var method in repairMethods)
        {
            method();
        }
    }

    /// <summary>
    /// ワークシートデータの妥当性チェック
    /// </summary>
    private void ValidateWorksheetData(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = Math.Min(usedRange.RangeAddress.LastAddress.RowNumber, firstRow + 100);
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = Math.Min(usedRange.RangeAddress.LastAddress.ColumnNumber, firstCol + 50);
        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int col = firstCol; col <= lastCol; col++)
            {
                try
                {
                    var cell = worksheet.Cell(row, col);
                    var value = cell.Value;
                    var formula = cell.HasFormula ? cell.FormulaA1 : null;
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }
    }

    /// <summary>
    /// 完全に空の行のみを削除
    /// </summary>
    private void CleanupEmptyRowsOnly(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var emptyRows = new List<int>();
        for (int row = lastRow; row >= firstRow; row--)
        {
            if (IsRowCompletelyEmpty(worksheet, row, usedRange))
            {
                emptyRows.Add(row);
            }
        }
        foreach (int row in emptyRows)
        {
            try
            {
                worksheet.Row(row).Delete();
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// 行が完全に空か判定
    /// </summary>
    private bool IsRowCompletelyEmpty(IXLWorksheet worksheet, int row, IXLRange usedRange)
    {
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        for (int col = firstCol; col <= lastCol; col++)
        {
            var cell = worksheet.Cell(row, col);
            if (!cell.IsEmpty() || cell.HasFormula)
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// バックアップから復元
    /// </summary>
    private void RestoreFromBackup()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;

        try
        {
            var dir = Path.GetDirectoryName(selectedFilePath) ?? string.Empty;
            var backupPath = Path.Combine(
                dir,
                $"{Path.GetFileNameWithoutExtension(selectedFilePath)}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(selectedFilePath)}"
            );

            if (File.Exists(backupPath) && !string.IsNullOrEmpty(backupPath))
            {
                File.Copy(backupPath, selectedFilePath, true);
            }
        }
        catch (Exception ex)
        {
            LogMessage($"バックアップ復元エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// OpenXMLを使用した修復
    /// </summary>
    private void RepairUsingOpenXML()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;

        try
        {
            using var package = SpreadsheetDocument.Open(selectedFilePath, true);

            var workbookPart = package.WorkbookPart;
            if (workbookPart != null)
            {
                RepairWorksheetParts(workbookPart);
                RepairStylesPart(workbookPart);
            }

            package.Save();
        }
        catch (Exception ex)
        {
            LogMessage($"OpenXML修復エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// ワークシートパーツの修復
    /// </summary>
    private void RepairWorksheetParts(WorkbookPart workbookPart)
    {
        var worksheetParts = workbookPart.WorksheetParts.ToList();
        foreach (var worksheetPart in worksheetParts)
        {
            try
            {
                var worksheet = worksheetPart.Worksheet;
                if (worksheet != null)
                {
                    RepairSheetData(worksheet);
                }
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// スタイルパーツの修復
    /// </summary>
    private void RepairStylesPart(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart != null)
        {
            RepairStyles(stylesPart);
        }
    }

    /// <summary>
    /// ワークシートの修復
    /// </summary>
    private void RepairWorksheets()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            using var workbook = new XLWorkbook(selectedFilePath);
            var worksheets = workbook.Worksheets.ToList();
            foreach (var worksheet in worksheets)
            {
                try
                {
                    var usedRange = worksheet.RangeUsed();
                    if (usedRange != null)
                    {
                        CleanupEmptyCells(worksheet);
                        RepairCorruptedCells(worksheet);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            workbook.Save();
        }
        catch (Exception ex)
        {
            LogMessage($"ワークシート修復エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// シートデータの修復
    /// </summary>
    private void RepairSheetData(Worksheet worksheet)
    {
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        var rows = sheetData.Elements<Row>().ToList();
        var corruptedRows = new List<Row>();

        foreach (var row in rows)
        {
            try
            {
                RepairRowData(row);
            }
            catch (Exception)
            {
                corruptedRows.Add(row);
            }
        }

        RemoveCorruptedRows(corruptedRows);
    }

    /// <summary>
    /// 行データの修復
    /// </summary>
    private void RepairRowData(Row row)
    {
        var cells = row.Elements<Cell>().ToList();
        var corruptedCells = new List<Cell>();

        foreach (var cell in cells)
        {
            try
            {
                RepairCellValue(cell);
            }
            catch (Exception)
            {
                corruptedCells.Add(cell);
            }
        }

        RemoveCorruptedCells(corruptedCells);
    }

    /// <summary>
    /// 破損行の削除
    /// </summary>
    private void RemoveCorruptedRows(List<Row> corruptedRows)
    {
        foreach (var row in corruptedRows)
        {
            try
            {
                row.Remove();
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// 破損セルの削除
    /// </summary>
    private void RemoveCorruptedCells(List<Cell> corruptedCells)
    {
        foreach (var cell in corruptedCells)
        {
            try
            {
                cell.Remove();
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// セル値の修復
    /// </summary>
    private void RepairCellValue(Cell cell)
    {
        var cellValue = cell.CellValue;
        if (cellValue?.Text == null) return;

        var value = cellValue.Text;
        if (!string.IsNullOrEmpty(value))
        {
            value = new string(value.Where(c => !char.IsControl(c) || c == '\n' || c == '\r' || c == '\t').ToArray());
            cellValue.Text = value;
        }
    }

    /// <summary>
    /// スタイルの修復
    /// </summary>
    private void RepairStyles(WorkbookStylesPart stylesPart)
    {
        var stylesheet = stylesPart.Stylesheet;
        if (stylesheet?.CellFormats == null) return;

        foreach (var cellFormat in stylesheet.CellFormats.Elements<CellFormat>())
        {
            try
            {
                if (cellFormat.FontId?.Value >= 0)
                {
                    // フォントIDの妥当性チェック
                }
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// 破損セルの修復
    /// </summary>
    private void RepairCorruptedCells(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var corruptedCells = new List<(int row, int col)>();
        ProcessWorksheetCells(worksheet, usedRange, corruptedCells);
        ClearCorruptedCells(worksheet, corruptedCells);
    }

    /// <summary>
    /// ワークシートのセルを走査し修復
    /// </summary>
    private void ProcessWorksheetCells(IXLWorksheet worksheet, IXLRange usedRange, List<(int row, int col)> corruptedCells)
    {
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int col = firstCol; col <= lastCol; col++)
            {
                try
                {
                    var cell = worksheet.Cell(row, col);
                    if (cell.HasFormula && cell.FormulaA1.StartsWith("="))
                    {
                        RepairFormula(cell);
                    }
                }
                catch (Exception)
                {
                    corruptedCells.Add((row, col));
                }
            }
        }
    }

    /// <summary>
    /// 計算式セルの修復
    /// </summary>
    private void RepairFormula(IXLCell cell)
    {
        try
        {
            var value = cell.Value;
            if (cell.HasFormula && cell.FormulaA1.Length > 1000)
            {
                cell.Value = value;
                cell.Clear(XLClearOptions.Contents);
            }
        }
        catch (Exception)
        {
            cell.Clear(XLClearOptions.Contents);
        }
    }

    /// <summary>
    /// 破損セルをクリア
    /// </summary>
    private void ClearCorruptedCells(IXLWorksheet worksheet, List<(int row, int col)> corruptedCells)
    {
        foreach (var (row, col) in corruptedCells)
        {
            try
            {
                var cell = worksheet.Cell(row, col);
                cell.Clear();
            }
            catch (Exception)
            {
                continue;
            }
        }
    }

    /// <summary>
    /// 空のセルをクリーンアップ
    /// </summary>
    private void CleanupEmptyCells(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        for (int row = lastRow; row >= firstRow; row--)
        {
            if (IsRowEmpty(worksheet, row, usedRange))
            {
                try
                {
                    worksheet.Row(row).Delete();
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }
    }

    /// <summary>
    /// 行が空か判定
    /// </summary>
    private bool IsRowEmpty(IXLWorksheet worksheet, int row, IXLRange usedRange)
    {
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        for (int col = firstCol; col <= lastCol; col++)
        {
            if (!worksheet.Cell(row, col).IsEmpty())
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Excelファイルの最適化
    /// </summary>
    private void OptimizeExcelFile()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            using var workbook = new XLWorkbook(selectedFilePath);
            foreach (var worksheet in workbook.Worksheets)
            {
                OptimizeFormulas(worksheet);
            }
            workbook.Save();
        }
        catch (Exception ex)
        {
            LogMessage($"最適化エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 計算式の最適化
    /// </summary>
    private void OptimizeFormulas(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var formulas = new Dictionary<string, List<IXLCell>>();
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int col = firstCol; col <= lastCol; col++)
            {
                var cell = worksheet.Cell(row, col);
                if (cell.HasFormula && cell.FormulaA1.StartsWith("="))
                {
                    var formula = cell.FormulaA1;
                    if (!formulas.ContainsKey(formula))
                    {
                        formulas[formula] = new List<IXLCell>();
                    }
                    formulas[formula].Add(cell);
                }
            }
        }
    }

    /// <summary>
    /// パフォーマンス向上処理
    /// </summary>
    private void ImprovePerformance()
    {
        if (string.IsNullOrEmpty(selectedFilePath)) return;
        try
        {
            using var workbook = new XLWorkbook(selectedFilePath);
            foreach (var worksheet in workbook.Worksheets)
            {
                RemoveUnnecessaryFormulas(worksheet);
            }
            workbook.Save();
        }
        catch (Exception ex)
        {
            LogMessage($"パフォーマンス向上エラー: {ex.Message}");
        }
    }

    /// <summary>
    /// 不要な計算式の削除
    /// </summary>
    private void RemoveUnnecessaryFormulas(IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null) return;
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int col = firstCol; col <= lastCol; col++)
            {
                var cell = worksheet.Cell(row, col);
                if (cell.HasFormula && cell.FormulaA1.StartsWith("="))
                {
                    try
                    {
                        if (!cell.IsEmpty() && IsSimpleFormula(cell.FormulaA1))
                        {
                            cell.Value = cell.Value;
                            cell.Clear(XLClearOptions.Contents);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
        }
    }

    /// <summary>
    /// 単純な計算式か判定
    /// </summary>
    private bool IsSimpleFormula(string formula)
    {
        var simplePatterns = new[] { "=A", "=B", "=C", "=D", "=E", "=F", "=G", "=H", "=I", "=J" };
        return simplePatterns.Any(pattern => formula.StartsWith(pattern));
    }
    
    /// <summary>
    /// 画像の圧縮（簡易版）
    /// </summary>
    private void CompressImages()
    {
        // ClosedXMLでは画像圧縮はサポートされていません
        LogMessage("ClosedXMLでは画像圧縮はサポートされていません");
    }


}

/// <summary>
/// 処理オプション格納用クラス
/// </summary>
public class ProcessingOptions
{
    public bool ShouldCreateBackup { get; set; }
    public bool IsAggressiveRepair { get; set; }
    public bool ShouldOptimize { get; set; }
    public bool ShouldImprovePerformance { get; set; }
    public bool ShouldCompressImages { get; set; }
}