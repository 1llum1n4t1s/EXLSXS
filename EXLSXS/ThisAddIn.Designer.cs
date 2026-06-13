using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Security.Permissions;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace EXLSXS
{
	[StartupObject(0)]
	[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
	public sealed partial class ThisAddIn : AddInBase
	{
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
			: base(factory, serviceProvider, "AddIn", "ThisAddIn")
		{
			Globals.Factory = factory;
		}

		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void Initialize()
		{
			base.Initialize();
			this.Application = base.GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), "Application");
			Globals.ThisAddIn = this;
			global::System.Windows.Forms.Application.EnableVisualStyles();
			this.InitializeCachedData();
			this.InitializeControls();
			this.InitializeComponents();
			this.InitializeData();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		protected override void FinishInitialization()
		{
			this.InternalStartup();
			this.OnStartup();
		}

		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void InitializeDataBindings()
		{
			this.BeginInitialization();
			this.BindToData();
			this.EndInitialization();
		}

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

		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void InitializeData()
		{
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		private void BindToData()
		{
		}

		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private void StartCaching(string MemberName)
		{
			base.DataHost.StartCaching(this, MemberName);
		}

		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private void StopCaching(string MemberName)
		{
			base.DataHost.StopCaching(this, MemberName);
		}

		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private bool IsCached(string MemberName)
		{
			return base.DataHost.IsCached(this, MemberName);
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		private void BeginInitialization()
		{
			this.BeginInit();
			this.CustomTaskPanes.BeginInit();
			this.VstoSmartTags.BeginInit();
		}

		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void EndInitialization()
		{
			this.VstoSmartTags.EndInit();
			this.CustomTaskPanes.EndInit();
			this.EndInit();
		}

		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		[EditorBrowsable(EditorBrowsableState.Never)]
		private void InitializeControls()
		{
			this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
			this.VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[DebuggerNonUserCode]
		private void InitializeComponents()
		{
		}

		[EditorBrowsable(EditorBrowsableState.Advanced)]
		[DebuggerNonUserCode]
		private bool NeedsFill(string MemberName)
		{
			return base.DataHost.NeedsFill(this, MemberName);
		}

		[DebuggerNonUserCode]
		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		[EditorBrowsable(EditorBrowsableState.Never)]
		protected override void OnShutdown()
		{
			this.VstoSmartTags.Dispose();
			this.CustomTaskPanes.Dispose();
			base.OnShutdown();
		}

		internal CustomTaskPaneCollection CustomTaskPanes;

		internal SmartTagCollection VstoSmartTags;

		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		private object missing = Type.Missing;

		[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
		internal Microsoft.Office.Interop.Excel.Application Application;
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
	internal sealed class Globals
	{
		private Globals()
		{
		}

		internal static ThisAddIn ThisAddIn
		{
			get
			{
				return Globals._ThisAddIn;
			}
			set
			{
				if (Globals._ThisAddIn == null)
				{
					Globals._ThisAddIn = value;
					return;
				}
				throw new NotSupportedException();
			}
		}

		internal static ApplicationFactory Factory
		{
			get
			{
				return Globals._factory;
			}
			set
			{
				if (Globals._factory == null)
				{
					Globals._factory = value;
					return;
				}
				throw new NotSupportedException();
			}
		}

		internal static ThisRibbonCollection Ribbons
		{
			get
			{
				if (Globals._ThisRibbonCollection == null)
				{
					Globals._ThisRibbonCollection = new ThisRibbonCollection(Globals._factory.GetRibbonFactory());
				}
				return Globals._ThisRibbonCollection;
			}
		}

		private static ThisAddIn _ThisAddIn;

		private static ApplicationFactory _factory;

		private static ThisRibbonCollection _ThisRibbonCollection;
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
	[DebuggerNonUserCode]
	internal sealed partial class ThisRibbonCollection : RibbonCollectionBase
	{
		internal ThisRibbonCollection(RibbonFactory factory)
			: base(factory)
		{
		}
	}
}
