using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using EXLSXS;

namespace ExcelAddinTestHost
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Excel を起動中...");
                var app = new Application();
                app.Visible = true;

                var addIn = new ThisAddIn();
                Console.WriteLine("アドインを初期化中...");
                addIn.AutoOpen();

                Console.WriteLine("新しいワークブックを作成中...");
                var workbook = app.Workbooks.Add();
                var worksheet = (Worksheet)workbook.Sheets[1];

                Console.WriteLine("テストデータを入力中...");
                worksheet.Range["A1"].Value = "Excel DNA Test";
                worksheet.Range["A2"].Value = "Sheet 1";
                worksheet.Range["A3"].Value = "Sheet 2";
                worksheet.Range["A4"].Value = "Sheet 3";

                Console.WriteLine("\nアドインが正常に読み込まれました！");
                Console.WriteLine("Ctrl+C を押すまでアドインが実行を続けます...");
                Console.ReadKey();

                Console.WriteLine("\nクリーンアップ中...");
                addIn.AutoClose();
                workbook.Close(false);
                app.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"エラー: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
