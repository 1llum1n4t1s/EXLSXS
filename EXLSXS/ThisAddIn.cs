using System;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace EXLSXS
{
	public sealed partial class ThisAddIn
	{
		internal string AssemblyTitle { get; private set; }

		internal string AssemblyProduct { get; private set; }

		internal string PublishVersion { get; private set; }

		internal string AssemblyCopyright { get; private set; }

		private void SetupAddInInfoProperty()
		{
			Assembly executingAssembly = Assembly.GetExecutingAssembly();
			object[] customAttributes = executingAssembly.GetCustomAttributes(false);
			this.AssemblyTitle = customAttributes.OfType<AssemblyTitleAttribute>().FirstOrDefault()?.Title ?? string.Empty;
			this.AssemblyProduct = customAttributes.OfType<AssemblyProductAttribute>().FirstOrDefault()?.Product ?? string.Empty;
			// Velopack + vstolocal 配布では ClickOnce の ApplicationDeployment.IsNetworkDeployed が常に
			// false になり版数が取れないため、アセンブリ版（Directory.Build.props の <Version> 由来）を表示する。
			Version assemblyVersion = executingAssembly.GetName().Version;
			this.PublishVersion = (assemblyVersion != null)
				? $"{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}"
				: "Unknown";
			this.AssemblyCopyright = customAttributes.OfType<AssemblyCopyrightAttribute>().FirstOrDefault()?.Copyright ?? string.Empty;
		}

		private MyRibbon AddInRibbon
		{
			get
			{
				return Globals.Ribbons.GetRibbon<MyRibbon>();
			}
		}

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			// EmbedInteropTypes=True のため、文字列ベースの ComAwareEventInfo は
			// イベントメタデータが埋め込まれず NullReferenceException になる。
			// コンパイラが COM イベントバインドを生成する通常の += で購読する。
			this.Application.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(this.Application_WorkbookActivate);
			this.Application.WorkbookDeactivate += new AppEvents_WorkbookDeactivateEventHandler(this.Application_WorkbookDeactivate);
			try
			{
				this.SetupAddInInfoProperty();
				MyRibbon ribbon = Globals.Ribbons.GetRibbon<MyRibbon>();
				ribbon.RefreshStatus(this.Application.Workbooks.Count > 0);
			}
			catch (Exception ex)
			{
				UIHelper.ShowErrorDialog(ex);
			}
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
		}

		private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
		{
			this.AddInRibbon.RefreshStatus(true);
		}

		private void Application_WorkbookDeactivate(Microsoft.Office.Interop.Excel.Workbook Wb)
		{
			bool flag = this.Application.Workbooks.Count == 1;
			this.AddInRibbon.RefreshStatus(!flag);
		}

		internal void DoFinish()
		{
			Microsoft.Office.Interop.Excel.Application application = this.Application;
			Microsoft.Office.Interop.Excel.Workbook workbook = application.ActiveWorkbook;
			if (workbook == null)
			{
				return;
			}

			MyRibbon addInRibbon = this.AddInRibbon;
			XlWindowView adjustView = addInRibbon.AdjustView;
			int adjustZoom = addInRibbon.AdjustZoom;
			bool adjustFont = addInRibbon.AdjustFont;
			string adjustFontName = adjustFont ? addInRibbon.AdjustFontName : string.Empty;
			bool restoreScreenUpdating = false;
			bool previousScreenUpdating = true;

			try
			{
				previousScreenUpdating = application.ScreenUpdating;
				application.ScreenUpdating = false;
				restoreScreenUpdating = true;
			}
			catch
			{
			}

			try
			{
				Microsoft.Office.Interop.Excel.Sheets worksheets = workbook.Worksheets;
				int count = worksheets.Count;
				for (int i = count; i > 0; i--)
				{
					try
					{
						object obj = worksheets[i];
						if (obj is Microsoft.Office.Interop.Excel._Worksheet worksheet)
						{
							if (adjustFont && !IsWorksheetProtected(worksheet))
							{
								worksheet.Cells.Font.Name = adjustFontName;
							}

							// 非表示シートはアクティブ化できないため、表示倍率・選択位置の変更はスキップする。
							if (!IsWorksheetVisible(worksheet))
							{
								continue;
							}

							worksheet.Activate();
							Microsoft.Office.Interop.Excel.Window activeWindow = application.ActiveWindow;
							if (activeWindow != null)
							{
								activeWindow.View = adjustView;
								activeWindow.Zoom = adjustZoom;
							}

							worksheet.get_Range("A1", Missing.Value).Select();

							// 各表示シートのスクロール位置を左上へ戻す。A1 を選択しても
							// ウィンドウ枠を固定/分割しているシートはスクロール対象ペインが
							// 先頭に戻らないため、シートごとに（＝アクティブな今のうちに）明示リセットする。
							// ここでの設定は ScreenUpdating=false でも各シートの表示状態として保持され、
							// 後でそのシートに切り替えた際に左上で表示される。最後に見るシートは
							// finally で画面更新を戻した後に再リセットして確実に反映させる。
							ScrollWindowToTopLeft(activeWindow);
						}
					}
					catch
					{
					}
				}
			}
			finally
			{
				// 仕上げ完了後は一番左 (先頭) の表示シートをアクティブにする。
				ActivateLeftmostVisibleWorksheet(workbook);

				if (restoreScreenUpdating)
				{
					try
					{
						application.ScreenUpdating = previousScreenUpdating;
					}
					catch
					{
					}
				}

				// 画面更新を戻した後、最後に見ている（先頭）シートのスクロールを左上へ確実に戻す。
				// ScrollRow/ScrollColumn は ScreenUpdating=false 中は表示へ即時反映されないことがあるため、
				// このタイミングでもう一度リセットして、ユーザーが見るシートを左上表示で確定させる。
				ScrollWindowToTopLeft(application.ActiveWindow);
			}
		}

		// ブックの一番左 (先頭) の表示シートをアクティブにする。
		// 非表示シートはアクティブ化できないため読み飛ばし、最初に見つかった表示シートをアクティブにする。
		private static void ActivateLeftmostVisibleWorksheet(Microsoft.Office.Interop.Excel.Workbook workbook)
		{
			try
			{
				Microsoft.Office.Interop.Excel.Sheets worksheets = workbook.Worksheets;
				int count = worksheets.Count;
				for (int i = 1; i <= count; i++)
				{
					if (worksheets[i] is Microsoft.Office.Interop.Excel._Worksheet worksheet && IsWorksheetVisible(worksheet))
					{
						worksheet.Activate();
						return;
					}
				}
			}
			catch
			{
			}
		}

		// ウィンドウのスクロール位置を左上へ戻す。
		// Window.ScrollColumn / ScrollRow は「固定枠を除いたスクロール領域」の先頭列 / 行を指す
		// (公式仕様: panes are frozen → ScrollRow excludes the frozen areas)。= 1 を指定すると
		// スクロール対象ペインが最初の非固定セルまで戻る (Excel が有効な最小値へ丸める)。
		// 行・列を固定したウィンドウでも、固定行の列スクロール / 固定列の行スクロールは
		// 対応ペイン間で同期されるため、Window への 1 回の指定で全ペインが左上で揃う。
		// ScrollColumn と ScrollRow は片方が失敗してももう片方を進めたいので個別に try/catch する。
		private static void ScrollWindowToTopLeft(Microsoft.Office.Interop.Excel.Window window)
		{
			if (window == null)
			{
				return;
			}

			try
			{
				window.ScrollColumn = 1;
			}
			catch
			{
			}

			try
			{
				window.ScrollRow = 1;
			}
			catch
			{
			}
		}

		private static bool IsWorksheetVisible(Microsoft.Office.Interop.Excel._Worksheet worksheet)
		{
			try
			{
				return worksheet.Visible == XlSheetVisibility.xlSheetVisible;
			}
			catch
			{
				return true;
			}
		}

		private static bool IsWorksheetProtected(Microsoft.Office.Interop.Excel._Worksheet worksheet)
		{
			try
			{
				return worksheet.ProtectContents;
			}
			catch
			{
				return false;
			}
		}

		private void InternalStartup()
		{
			base.Startup += this.ThisAddIn_Startup;
			base.Shutdown += this.ThisAddIn_Shutdown;
		}
	}
}
