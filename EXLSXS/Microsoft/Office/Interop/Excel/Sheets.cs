using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000013 RID: 19
	[CompilerGenerated]
	[TypeIdentifier]
	[Guid("000208D7-0000-0000-C000-000000000046")]
	[ComImport]
	public interface Sheets : IEnumerable
	{
		// Token: 0x06000061 RID: 97
		void _VtblGap1_5();

		// Token: 0x17000018 RID: 24
		// (get) Token: 0x06000062 RID: 98
		int Count
		{
			[DispId(118)]
			get;
		}

		// Token: 0x06000063 RID: 99
		void _VtblGap2_12();

		// Token: 0x17000019 RID: 25
		[IndexerName("_Default")]
		object this[[MarshalAs(UnmanagedType.Struct)] [In] object Index]
		{
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.IDispatch)]
			get;
		}
	}
}
