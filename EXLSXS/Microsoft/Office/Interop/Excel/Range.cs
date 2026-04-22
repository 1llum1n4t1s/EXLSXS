using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000016 RID: 22
	[Guid("00020846-0000-0000-C000-000000000046")]
	[DefaultMember("_Default")]
	[TypeIdentifier]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	[CompilerGenerated]
	[ComImport]
	public interface Range : IEnumerable
	{
		// Token: 0x1700001E RID: 30
		// (get) Token: 0x0600006F RID: 111
		Font Font
		{
			[DispId(146)]
			[PreserveSig]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000070 RID: 112
		[DispId(235)]
		[PreserveSig]
		[return: MarshalAs(UnmanagedType.Struct)]
		object Select();
	}
}
