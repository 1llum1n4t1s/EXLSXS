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

				// 最後に先頭シートのスクロールを左上へ戻す。
				// ウィンドウ枠を固定していると A1 を選択してもスクロール対象ペイン（固定枠の外側）は
				// 先頭に戻らないため、ここで明示的に ScrollRow/ScrollColumn を先頭にする。
				// ScrollRow/ScrollColumn は ScreenUpdating=false 中は確実に反映されないので、
				// 画面更新を戻した後のこのタイミングで実施する。設定できない表示状態（特殊なビュー等）は
				// catch で諦める。
				try
				{
					Microsoft.Office.Interop.Excel.Window finalWindow = application.ActiveWindow;
					if (finalWindow != null)
					{
						finalWindow.ScrollRow = 1;
						finalWindow.ScrollColumn = 1;
					}
				}
				catch
				{
				}
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
