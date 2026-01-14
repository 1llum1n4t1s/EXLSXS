using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel
{
	// Token: 0x02000012 RID: 18
	[Guid("000208DA-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[TypeIdentifier]
	[ComImport]
	public interface _Workbook
	{
		// Token: 0x0600005F RID: 95
		void _VtblGap1_124();

		// Token: 0x17000017 RID: 23
		// (get) Token: 0x06000060 RID: 96
		Sheets Worksheets { [DispId(494)] [return: MarshalAs(UnmanagedType.Interface)] get; }
	}
}
