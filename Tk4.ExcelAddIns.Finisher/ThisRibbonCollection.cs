using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools.Ribbon;

namespace Tk4.ExcelAddIns.Finisher
{
	// Token: 0x02000004 RID: 4
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
	[DebuggerNonUserCode]
	internal sealed class ThisRibbonCollection : RibbonCollectionBase
	{
		// Token: 0x17000005 RID: 5
		// (get) Token: 0x06000018 RID: 24 RVA: 0x00003164 File Offset: 0x00001364
		internal MyRibbon MyRibbon
		{
			get
			{
				return base.GetRibbon<MyRibbon>();
			}
		}

		// Token: 0x06000019 RID: 25 RVA: 0x0000316C File Offset: 0x0000136C
		internal ThisRibbonCollection(RibbonFactory factory) : base(factory)
		{
		}
	}
}
