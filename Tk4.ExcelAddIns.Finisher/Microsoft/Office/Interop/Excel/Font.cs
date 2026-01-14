using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000017 RID: 23
	[CompilerGenerated]
	[Guid("0002084D-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	[ComImport]
	public interface Font
	{
		// Token: 0x1700001F RID: 31
		// (get) Token: 0x06000071 RID: 113
		// (set) Token: 0x06000072 RID: 114
		object Name { [DispId(110)] [PreserveSig] [return: MarshalAs(UnmanagedType.Struct)] get; [DispId(110)] [PreserveSig] [param: MarshalAs(UnmanagedType.Struct)] [param: In] set; }
	}
}
