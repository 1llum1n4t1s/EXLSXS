using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

// AssemblyVersion / AssemblyFileVersion は Directory.Build.props の <Version> から
// EXLSXS.csproj の GenerateVersionAttributes ターゲットで生成する (ここには固定値を持たない)。
[assembly: AssemblyDescription("表示倍率やセルの選択位置を整える Excel アドインです。")]
[assembly: AssemblyTitle("EXLSXS")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("EXLSXS")]
[assembly: AssemblyProduct("EXLSXS")]
[assembly: AssemblyCopyright("Copyright © EXLSXS")]
[assembly: AssemblyTrademark("")]
[assembly: ComVisible(false)]
[assembly: Guid("1e5d8dc8-72be-49d7-b2ac-8ce0eb134425")]
