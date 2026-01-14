using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000019 RID: 25
	[CompilerGenerated]
	[TypeIdentifier]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	[Guid("00024413-0000-0000-C000-000000000046")]
	[ComImport]
	public interface AppEvents
	{
		// Token: 0x06000073 RID: 115
		[DispId(1568)]
		[PreserveSig]
		void WorkbookActivate([MarshalAs(UnmanagedType.Interface)] [In] Workbook Wb);

		// Token: 0x06000074 RID: 116
		[DispId(1569)]
		[PreserveSig]
		void WorkbookDeactivate([MarshalAs(UnmanagedType.Interface)] [In] Workbook Wb);
	}
}
