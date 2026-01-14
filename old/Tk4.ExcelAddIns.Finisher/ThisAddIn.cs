using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Deployment.Application;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace Tk4.ExcelAddIns.Finisher
{
	// Token: 0x02000005 RID: 5
	[StartupObject(0)]
	[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
	public sealed class ThisAddIn : AddInBase
	{
		// Token: 0x17000006 RID: 6
		// (get) Token: 0x0600001A RID: 26 RVA: 0x00003175 File Offset: 0x00001375
		// (set) Token: 0x0600001B RID: 27 RVA: 0x0000317D File Offset: 0x0000137D
		internal string AssemblyTitle { get; private set; }

		// Token: 0x17000007 RID: 7
		// (get) Token: 0x0600001C RID: 28 RVA: 0x00003186 File Offset: 0x00001386
		// (set) Token: 0x0600001D RID: 29 RVA: 0x0000318E File Offset: 0x0000138E
		internal string AssemblyProduct { get; private set; }

		// Token: 0x17000008 RID: 8
		// (get) Token: 0x0600001E RID: 30 RVA: 0x00003197 File Offset: 0x00001397
		// (set) Token: 0x0600001F RID: 31 RVA: 0x0000319F File Offset: 0x0000139F
		internal string PublishVersion { get; private set; }

		// Token: 0x17000009 RID: 9
		// (get) Token: 0x06000020 RID: 32 RVA: 0x000031A8 File Offset: 0x000013A8
		// (set) Token: 0x06000021 RID: 33 RVA: 0x000031B0 File Offset: 0x000013B0
		internal string AssemblyCopyright { get; private set; }

		// Token: 0x06000022 RID: 34 RVA: 0x000031BC File Offset: 0x000013BC
		private void SetupAddInInfoProperty()
		{
			Assembly executingAssembly = Assembly.GetExecutingAssembly();
			object[] customAttributes = executingAssembly.GetCustomAttributes(false);
			AssemblyTitleAttribute[] array = customAttributes.OfType<AssemblyTitleAttribute>().ToArray<AssemblyTitleAttribute>();
			this.AssemblyTitle = ((array.Length > 0) ? array[0].Title : string.Empty);
			AssemblyProductAttribute[] array2 = customAttributes.OfType<AssemblyProductAttribute>().ToArray<AssemblyProductAttribute>();
			this.AssemblyProduct = ((array2.Length > 0) ? array2[0].Product : string.Empty);
			if (ApplicationDeployment.IsNetworkDeployed)
			{
				Version currentVersion = ApplicationDeployment.CurrentDeployment.CurrentVersion;
				this.PublishVersion = string.Format("{0}.{1}.{2}.{3}", new object[]
				{
					currentVersion.Major,
					currentVersion.Minor,
					currentVersion.Build,
					currentVersion.Revision
				});
			}
			else
			{
				this.PublishVersion = "Unknown(Not published)";
			}
			AssemblyCopyrightAttribute[] array3 = customAttributes.OfType<AssemblyCopyrightAttribute>().ToArray<AssemblyCopyrightAttribute>();
			this.AssemblyCopyright = ((array3.Length > 0) ? array3[0].Copyright : string.Empty);
		}

		// Token: 0x1700000A RID: 10
		// (get) Token: 0x06000023 RID: 35 RVA: 0x000032C9 File Offset: 0x000014C9
		private MyRibbon AddInRibbon
		{
			get
			{
				return Globals.Ribbons.GetRibbon<MyRibbon>();
			}
		}

		// Token: 0x06000024 RID: 36 RVA: 0x000032D8 File Offset: 0x000014D8
		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Excel.AppEvents_Event), "WorkbookActivate").AddEventHandler(this.Application, new AppEvents_WorkbookActivateEventHandler(this, (UIntPtr)ldftn(Application_WorkbookActivate)));
			new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Excel.AppEvents_Event), "WorkbookDeactivate").AddEventHandler(this.Application, new AppEvents_WorkbookDeactivateEventHandler(this, (UIntPtr)ldftn(Application_WorkbookDeactivate)));
			try
			{
				this.SetupAddInInfoProperty();
				MyRibbon ribbon = Globals.Ribbons.GetRibbon<MyRibbon>();
				ribbon.RefreshStatus(this.Application.Workbooks.Count > 0);
			}
			catch (Exception e2)
			{
				UIHelper.ShowErrorDialog(e2);
			}
		}

		// Token: 0x06000025 RID: 37 RVA: 0x00003380 File Offset: 0x00001580
		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
		}

		// Token: 0x06000026 RID: 38 RVA: 0x00003382 File Offset: 0x00001582
		private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
		{
			this.AddInRibbon.RefreshStatus(true);
		}

		// Token: 0x06000027 RID: 39 RVA: 0x00003390 File Offset: 0x00001590
		private void Application_WorkbookDeactivate(Microsoft.Office.Interop.Excel.Workbook Wb)
		{
			bool flag = this.Application.Workbooks.Count == 1;
			this.AddInRibbon.RefreshStatus(!flag);
		}

		// Token: 0x06000028 RID: 40 RVA: 0x000033C0 File Offset: 0x000015C0
		internal void DoFinish()
		{
			int count = this.Application.ActiveWorkbook.Worksheets.Count;
			for (int i = count; i > 0; i--)
			{
				try
				{
					object obj = this.Application.ActiveWorkbook.Worksheets[i];
					if (obj is Microsoft.Office.Interop.Excel._Worksheet)
					{
						Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel._Worksheet)obj;
						MyRibbon addInRibbon = this.AddInRibbon;
						worksheet.Activate();
						if (this.Application.ActiveWindow != null)
						{
							this.Application.ActiveWindow.View = addInRibbon.AdjustView;
							this.Application.ActiveWindow.Zoom = addInRibbon.AdjustZoom;
						}
						if (addInRibbon.AdjustFont)
						{
							worksheet.Cells.Font.Name = addInRibbon.AdjustFontName;
						}
						worksheet.get_Range("A1", Missing.Value).Select();
					}
				}
				catch
				{
				}
			}
		}

		// Token: 0x06000029 RID: 41 RVA: 0x000034BC File Offset: 0x000016BC
		private void InternalStartup()
		{
			base.Startup += this.ThisAddIn_Startup;
			base.Shutdown += this.ThisAddIn_Shutdown;
		}

		// Token: 0x0600002A RID: 42 RVA: 0x000034E2 File Offset: 0x000016E2
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider) : base(factory, serviceProvider, "AddIn", "ThisAddIn")
		{
			Globals.Factory = factory;
		}

		// Token: 0x0600002B RID: 43 RVA: 0x00003508 File Offset: 0x00001708
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void Initialize()
		{
			base.Initialize();
			this.Application = base.GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), "Application");
			Globals.ThisAddIn = this;
			System.Windows.Forms.Application.EnableVisualStyles();
			this.InitializeCachedData();
			this.InitializeControls();
			this.InitializeComponents();
			this.InitializeData();
		}

		// Token: 0x0600002C RID: 44 RVA: 0x00003559 File Offset: 0x00001759
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		protected override void FinishInitialization()
		{
			this.InternalStartup();
			this.OnStartup();
		}

		// Token: 0x0600002D RID: 45 RVA: 0x00003567 File Offset: 0x00001767
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void InitializeDataBindings()
		{
			this.BeginInitialization();
			this.BindToData();
			this.EndInitialization();
		}

		// Token: 0x0600002E RID: 46 RVA: 0x0000357B File Offset: 0x0000177B
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		private void InitializeCachedData()
		{
			if (base.DataHost == null)
			{
				return;
			}
			if (base.DataHost.IsCacheInitialized)
			{
				base.DataHost.FillCachedData(this);
			}
		}

		// Token: 0x0600002F RID: 47 RVA: 0x0000359F File Offset: 0x0000179F
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void InitializeData()
		{
		}

		// Token: 0x06000030 RID: 48 RVA: 0x000035A1 File Offset: 0x000017A1
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		private void BindToData()
		{
		}

		// Token: 0x06000031 RID: 49 RVA: 0x000035A3 File Offset: 0x000017A3
		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private void StartCaching(string MemberName)
		{
			base.DataHost.StartCaching(this, MemberName);
		}

		// Token: 0x06000032 RID: 50 RVA: 0x000035B2 File Offset: 0x000017B2
		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private void StopCaching(string MemberName)
		{
			base.DataHost.StopCaching(this, MemberName);
		}

		// Token: 0x06000033 RID: 51 RVA: 0x000035C1 File Offset: 0x000017C1
		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private bool IsCached(string MemberName)
		{
			return base.DataHost.IsCached(this, MemberName);
		}

		// Token: 0x06000034 RID: 52 RVA: 0x000035D0 File Offset: 0x000017D0
		[EditorBrowsable(EditorBrowsableState.Never)]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		private void BeginInitialization()
		{
			this.BeginInit();
			this.CustomTaskPanes.BeginInit();
			this.VstoSmartTags.BeginInit();
		}

		// Token: 0x06000035 RID: 53 RVA: 0x000035EE File Offset: 0x000017EE
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void EndInitialization()
		{
			this.VstoSmartTags.EndInit();
			this.CustomTaskPanes.EndInit();
			this.EndInit();
		}

		// Token: 0x06000036 RID: 54 RVA: 0x0000360C File Offset: 0x0000180C
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void InitializeControls()
		{
			this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
			this.VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
		}

		// Token: 0x06000037 RID: 55 RVA: 0x00003648 File Offset: 0x00001848
		[EditorBrowsable(EditorBrowsableState.Never)]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		private void InitializeComponents()
		{
		}

		// Token: 0x06000038 RID: 56 RVA: 0x0000364A File Offset: 0x0000184A
		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private bool NeedsFill(string MemberName)
		{
			return base.DataHost.NeedsFill(this, MemberName);
		}

		// Token: 0x06000039 RID: 57 RVA: 0x00003659 File Offset: 0x00001859
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void OnShutdown()
		{
			this.VstoSmartTags.Dispose();
			this.CustomTaskPanes.Dispose();
			base.OnShutdown();
		}

		// Token: 0x04000016 RID: 22
		internal CustomTaskPaneCollection CustomTaskPanes;

		// Token: 0x04000017 RID: 23
		internal SmartTagCollection VstoSmartTags;

		// Token: 0x04000018 RID: 24
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		private object missing = Type.Missing;

		// Token: 0x04000019 RID: 25
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		internal Microsoft.Office.Interop.Excel.Application Application;
	}
}
