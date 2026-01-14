using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000014 RID: 20
	[CompilerGenerated]
	[Guid("000208D8-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[ComImport]
	public interface _Worksheet
	{
		// Token: 0x06000065 RID: 101
		void _VtblGap1_3();

		// Token: 0x06000066 RID: 102
		[LCIDConversion(0)]
		[DispId(304)]
		void Activate();

		// Token: 0x06000067 RID: 103
		void _VtblGap2_41();

		// Token: 0x1700001A RID: 26
		// (get) Token: 0x06000068 RID: 104
		Range Cells { [DispId(238)] [return: MarshalAs(UnmanagedType.Interface)] get; }

		// Token: 0x06000069 RID: 105
		void _VtblGap3_47();

		// Token: 0x1700001B RID: 27
		// (get) Token: 0x0600006A RID: 106
		Range Range { [DispId(197)] [return: MarshalAs(UnmanagedType.Interface)] get; }
	}
}
