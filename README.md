# EXLSXS

Excel 用 VSTO add-in を Velopack で配布する構成です。

## 構成

- `EXLSXS/`
  - .NET Framework 4.8.1 の VSTO add-in 本体
- `EXLSXS.Host/`
  - Velopack で配布されるホスト
  - インストール/更新時の VSTO 登録
  - Windows 起動時のサイレント更新確認
- `build/pack-velopack.ps1`
  - VSTO publish
  - host publish
  - Velopack package 生成

## ローカルビルド

```powershell
MSBuild EXLSXS.sln /t:Restore,Rebuild /p:Configuration=Release /p:Platform="Any CPU"
powershell -NoProfile -ExecutionPolicy Bypass -File .\build\pack-velopack.ps1
```

出力先:

- `C:\Users\szk\Work\EXLSXS\artifacts\velopack-releases\EXLSXS-win-Setup.exe`

ローカルでは Velopack 署名パラメータを渡さない限り未署名パッケージになります。サイレント更新の自動適用は、署名済み package と期待 publisher 設定が揃っている場合だけ通ります。

## 配布時の前提条件

クライアントには以下が必要です。

- .NET Framework 4.8.1
- Visual Studio Tools for Office Runtime

Velopack package には `vsto\setup.exe` を同梱しています。host 側でも前提条件チェックを行い、不足時は登録を失敗として扱います。

## GitHub Actions 用 secret

`.github/workflows/velopack-release.yml` は以下の secret を前提にしています。

- `EXLSXS_VSTO_SIGNING_PFX_BASE64`
  - VSTO manifest signing 用 PFX の base64
- `EXLSXS_VSTO_SIGNING_PFX_PASSWORD`
  - 上記 PFX のパスワード
- `EXLSXS_VSTO_SIGNING_THUMBPRINT`
  - CurrentUser\\My に import された証明書の thumbprint
- `EXLSXS_VELOPACK_SIGN_PARAMS`
  - `signtool sign ...` 相当の Velopack 用署名パラメータ
- `EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE`
  - Azure Trusted Signing を使う場合はこちら
- `EXLSXS_EXPECTED_PUBLISHER_THUMBPRINT`
  - クライアントが更新適用前に照合する publisher thumbprint
- `EXLSXS_EXPECTED_PUBLISHER_SUBJECT`
  - 追加照合したい subject 文字列。任意

`release/**` ブランチの workflow では build job / publish job を分離し、どちらも `release` environment 配下で動かす前提です。

## 更新フロー

1. Velopack install / update
2. host が前提条件を確認
3. host が VSTO manifest を登録
4. Windows 起動時に `EXLSXS.Host.exe --update-check`
5. 更新 package を download
6. package 内の `EXLSXS.Host.exe` / `EXLSXS.Host.dll` / `EXLSXS.dll(.deploy)` の署名者 thumbprint を検証
7. 一致した場合だけサイレント更新を適用

## Excel 側の挙動

`DoFinish()` は以下を行います。

- リボンで選んだ view / zoom を各シートへ適用
- 必要ならフォントを統一
- `A1` を選択
- 失敗したシートは最後にまとめて警告表示

非表示シートや保護シートは、できなかった理由をユーザーに返します。
