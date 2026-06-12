# EXLSXS

Excel 用 VSTO add-in を Velopack で配布する構成です。自動更新の配信元は Cloudflare R2 (`https://exlsxs.nephilim.jp`) です。

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
MSBuild EXLSXS.slnx /t:Restore,Rebuild /p:Configuration=Release /p:Platform="Any CPU"
powershell -NoProfile -ExecutionPolicy Bypass -File .\build\pack-velopack.ps1
```

出力先: `artifacts\velopack-releases\EXLSXS-win-Setup.exe`

`pack-velopack.ps1` に署名パラメータを渡さない場合は未署名パッケージになります。サイレント更新の自動適用は、署名済み package と期待 publisher 設定が揃っている場合だけ通ります。

## リリース (ローカル署名)

配信元の Cloudflare R2 へは、コード署名付きで `scripts/release-local.ps1` から公開します (SimplySign のトークンログインが必要なため CI では署名できません)。

```powershell
pwsh -NoProfile -File scripts\release-local.ps1 -SkipUpload  # ビルド + 署名のみ確認
pwsh -NoProfile -File scripts\release-local.ps1              # 署名 + R2 アップロード + 配信確認 + 旧版掃除
```

`/vava` でバージョンを上げると、precheck (署名証明書確認) → このスクリプトの自動実行までを行います。

## 配布時の前提条件

クライアントには以下が必要です。

- .NET Framework 4.8.1
- Visual Studio Tools for Office Runtime

Velopack package には `vsto\setup.exe` を同梱しています。host 側でも前提条件チェックを行い、不足時は登録を失敗として扱います。

## リリースの前提

リリースは `scripts/release-local.ps1` でローカル実行します (GitHub Actions では署名できないため CI リリースは持ちません)。

- **SimplySign Desktop** がトークンログイン済みで、コード署名証明書 (`CN=Open Source Developer Yuichiro Shinozaki`) が `Cert:\CurrentUser\My` に見えていること
- `C:\Users\IMT\dev\Secret\secrets.json` の `cloudflare.api_token` に R2 アップロード用トークンがあること
- ClickOnce manifest 署名 (VSTO) と Velopack package の Authenticode 署名は同一証明書で行い、クライアントはその publisher thumbprint を更新適用前に照合します

## 更新フロー

1. Velopack install / update
2. host が前提条件を確認
3. host が VSTO manifest を登録
4. Windows 起動時に `EXLSXS.Host.exe --update-check`
5. Cloudflare R2 (`exlsxs.nephilim.jp`) から更新 package を download
6. package 内の `EXLSXS.Host.exe` / `EXLSXS.Host.dll` / `EXLSXS.dll(.deploy)` の署名者 thumbprint を検証
7. 一致した場合だけサイレント更新を適用

## Excel 側の挙動

`DoFinish()` は以下を行います。

- リボンで選んだ view / zoom を各シートへ適用
- 必要ならフォントを統一
- `A1` を選択
- 失敗したシートは最後にまとめて警告表示

非表示シートや保護シートは、できなかった理由をユーザーに返します。
