using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000015 RID: 21
	[CompilerGenerated]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	[Guid("00020893-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[ComImport]
	public interface Window
	{
		// Token: 0x1700001C RID: 28
		// (get) Token: 0x0600006B RID: 107
		// (set) Token: 0x0600006C RID: 108
		object Zoom { [DispId(663)] [PreserveSig] [return: MarshalAs(UnmanagedType.Struct)] get; [DispId(663)] [PreserveSig] [param: MarshalAs(UnmanagedType.Struct)] [param: In] set; }

		// Token: 0x1700001D RID: 29
		// (get) Token: 0x0600006D RID: 109
		// (set) Token: 0x0600006E RID: 110
		XlWindowView View { [DispId(1194)] [PreserveSig] get; [DispId(1194)] [PreserveSig] [param: In] set; }
	}
}
