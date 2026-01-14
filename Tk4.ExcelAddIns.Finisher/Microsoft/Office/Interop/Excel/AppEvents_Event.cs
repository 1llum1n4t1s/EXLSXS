using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x0200000E RID: 14
	[TypeIdentifier("00020813-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Excel.AppEvents_Event")]
	[ComEventInterface(typeof(AppEvents), typeof(AppEvents))]
	[CompilerGenerated]
	[ComImport]
	public interface AppEvents_Event
	{
		// Token: 0x06000054 RID: 84
		void _VtblGap1_18();

		// Token: 0x14000001 RID: 1
		// (add) Token: 0x06000055 RID: 85
		// (remove) Token: 0x06000056 RID: 86
		event AppEvents_WorkbookActivateEventHandler WorkbookActivate;

		// Token: 0x14000002 RID: 2
		// (add) Token: 0x06000057 RID: 87
		// (remove) Token: 0x06000058 RID: 88
		event AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivate;
	}
}
