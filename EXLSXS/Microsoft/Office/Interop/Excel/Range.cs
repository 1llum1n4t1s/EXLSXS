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

		// RowHeight / ColumnWidth / NumberFormat: DispId は実 Excel PIA
		// (Microsoft.Office.Interop.Excel, GUID 00020846-...) の Range メンバと一致する値。
		object RowHeight
		{
			[DispId(272)]
			[PreserveSig]
			[return: MarshalAs(UnmanagedType.Struct)]
			get;
			[DispId(272)]
			[PreserveSig]
			[param: MarshalAs(UnmanagedType.Struct)]
			[param: In]
			set;
		}

		object ColumnWidth
		{
			[DispId(242)]
			[PreserveSig]
			[return: MarshalAs(UnmanagedType.Struct)]
			get;
			[DispId(242)]
			[PreserveSig]
			[param: MarshalAs(UnmanagedType.Struct)]
			[param: In]
			set;
		}

		object NumberFormat
		{
			[DispId(193)]
			[PreserveSig]
			[return: MarshalAs(UnmanagedType.Struct)]
			get;
			[DispId(193)]
			[PreserveSig]
			[param: MarshalAs(UnmanagedType.Struct)]
			[param: In]
			set;
		}

		// Token: 0x06000070 RID: 112
		[DispId(235)]
		[PreserveSig]
		[return: MarshalAs(UnmanagedType.Struct)]
		object Select();
	}
}
