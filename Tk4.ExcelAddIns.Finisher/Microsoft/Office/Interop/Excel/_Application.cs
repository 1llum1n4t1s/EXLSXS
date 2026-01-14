using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x0200000D RID: 13
	[TypeIdentifier]
	[DefaultMember("_Default")]
	[Guid("000208D5-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[ComImport]
	public interface _Application
	{
		// Token: 0x0600004C RID: 76
		void _VtblGap1_10();

		// Token: 0x17000012 RID: 18
		// (get) Token: 0x0600004D RID: 77
		Window ActiveWindow { [DispId(759)] [return: MarshalAs(UnmanagedType.Interface)] get; }

		// Token: 0x17000013 RID: 19
		// (get) Token: 0x0600004E RID: 78
		Workbook ActiveWorkbook { [DispId(308)] [return: MarshalAs(UnmanagedType.Interface)] get; }

		// Token: 0x0600004F RID: 79
		void _VtblGap2_33();

		// Token: 0x17000014 RID: 20
		// (get) Token: 0x06000050 RID: 80
		Workbooks Workbooks { [DispId(572)] [return: MarshalAs(UnmanagedType.Interface)] get; }

		// Token: 0x06000051 RID: 81
		void _VtblGap3_199();

		// Token: 0x17000015 RID: 21
		// (get) Token: 0x06000052 RID: 82
		// (set) Token: 0x06000053 RID: 83
		string StandardFont { [DispId(924)] [LCIDConversion(0)] [return: MarshalAs(UnmanagedType.BStr)] get; [LCIDConversion(0)] [DispId(924)] [param: MarshalAs(UnmanagedType.BStr)] [param: In] set; }
	}
}
