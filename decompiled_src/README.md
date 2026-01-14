# Office アドインのソース化（復元）について

このリポジトリは ClickOnce 形式の VSTO アドイン配布物のみが含まれており、
通常の C# ソースが含まれていません。そこで、配布物に含まれるマニフェストや
DLL 内のシンボル名から、**最小限のソース構造**を復元しました。

## ここに含めたもの

- `Tk4.ExcelAddIns.Finisher` の**名前空間/型/主要メソッド名**の復元
- Excel VSTO アドインで想定される `ThisAddIn` / `MyRibbon` の雛形
- 復元根拠となるマニフェスト/文字列抽出情報の一覧

## 注意点

- DLL からの完全なデコンパイルは、.NET デコンパイラが必要です。
- 現在の環境ではデコンパイラを導入できないため、**メソッド本体は未復元**です。
- そのため、本成果物は**ソース構造の復元**に主眼を置いたものです。

必要に応じて、ローカル環境で ILSpy 等を用いて DLL を開き、
`ThisAddIn` や `MyRibbon` のメソッド本体を補完してください。

（復元対象）
- `Application Files/Tk4.ExcelAddIns.Finisher_1_2_0_0/Tk4.ExcelAddIns.Finisher.dll.deploy`
- `Tk4.ExcelAddIns.Finisher.vsto`
- `Application Files/Tk4.ExcelAddIns.Finisher_1_2_0_0/Tk4.ExcelAddIns.Finisher.dll.manifest`
