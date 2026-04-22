using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools.Excel;

namespace EXLSXS
{
	// Token: 0x02000008 RID: 8
	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
	internal sealed class Globals
	{
		// Token: 0x06000042 RID: 66 RVA: 0x0000371D File Offset: 0x0000191D
		private Globals()
		{
		}

		// Token: 0x1700000F RID: 15
		// (get) Token: 0x06000043 RID: 67 RVA: 0x00003725 File Offset: 0x00001925
		// (set) Token: 0x06000044 RID: 68 RVA: 0x0000372C File Offset: 0x0000192C
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

		// Token: 0x17000010 RID: 16
		// (get) Token: 0x06000045 RID: 69 RVA: 0x00003741 File Offset: 0x00001941
		// (set) Token: 0x06000046 RID: 70 RVA: 0x00003748 File Offset: 0x00001948
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

		// Token: 0x17000011 RID: 17
		// (get) Token: 0x06000047 RID: 71 RVA: 0x0000375D File Offset: 0x0000195D
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

		// Token: 0x04000021 RID: 33
		private static ThisAddIn _ThisAddIn;

		// Token: 0x04000022 RID: 34
		private static ApplicationFactory _factory;

		// Token: 0x04000023 RID: 35
		private static ThisRibbonCollection _ThisRibbonCollection;
	}
}
