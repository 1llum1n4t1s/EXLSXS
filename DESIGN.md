# EXLSXS システム設計

この文書は、現在のコードと配布設定から確認できるEXLSXSの構造、責務、境界、データフロー、設計判断を記録する。開発・検証コマンドと運用上の制約は[AGENTS.md](AGENTS.md)を正本とする。

## 目的

EXLSXSは、Excelブック内の全ワークシートへ表示モード、表示倍率、フォント、セル寸法、表示形式、選択位置を一括適用するWindows用VSTOアドインである。アドイン本体とインストール・登録・自動更新を担うホストを分離し、Velopackパッケージとして配布する。

## 主要コンポーネント

| コンポーネント | 実行環境 | 責務と境界 |
| --- | --- | --- |
| `EXLSXS/` | .NET Framework 4.8.1 / Excel VSTO | リボンUI、設定値の保持、Excel COMを介した全シート整形を担う。配布・更新処理は持たない |
| `EXLSXS.Host/` | .NET 10 Windows | Velopackのinstall/update/uninstall callback、前提条件確認、VSTO登録、Windows起動時登録、自動更新、更新パッケージの署名検証を担う。Excelのシート操作は行わない |
| `build/pack-velopack.ps1` | PowerShell | VSTO publish、ホストpublish、フラットなstaging構築、署名、`vpk pack`を一つのパッケージング処理へまとめる |
| `scripts/release-local.ps1` | PowerShell | リリース前検査、署名付きpack、署名検証、R2 upload、Cloudflare cache purge、公開manifest確認、旧世代ファイルの整理を直列実行する |
| `web/` | Cloudflare Worker + R2 custom domain | `/`と`/index.html`ではランディングページを返し、それ以外の更新ファイル要求は加工せずR2へ委譲する |
| `EXLSXS.Host.Tests/` | .NET 10 / xUnit 4 / Microsoft Testing Platform | 更新元判定、更新設定、署名検証用の正規化規則など、ホストの外部境界に近い規則を検証する |

## データフロー

### Excelブックの一括整形

1. `MyRibbon`が表示モードと倍率をユーザー単位のレジストリから復元し、リボンの選択肢を構築する。
2. 実行ボタンが`ThisAddIn.DoFinish`を呼び、現在のリボン設定を処理用の値へ確定する。
3. アクティブブックのワークシートを走査し、保護されていないシートへフォント、行高・列幅、表示形式を適用する。
4. 表示中のシートだけをアクティブ化し、表示モード、倍率、A1選択、スクロール位置を更新する。非表示シートはアクティブ化しない。
5. 最後に一番左の表示シートをアクティブ化し、Excelの`ScreenUpdating`を処理前の状態へ戻す。

### インストールと登録

1. Velopackのinstall callbackは、.NET Framework 4.8.1とVSTO Runtimeを確認し、不足時だけ同梱bootstrapperを信頼検証して実行する。
2. `VstoRegistration`が32-bit/64-bit Excel用のHKCU add-in keyへ`Manifest`、`FriendlyName`、`Description`、`LoadBehavior`を書き込む。manifestはローカル配置の`.vsto|vstolocal`を参照する。
3. installまたは明示的な`--register`では`LoadBehavior=3`を強制する。通常起動やupdate callbackでは既存の無効化状態を上書きしない。
4. `StartupRegistration`がHKCUのWindows Run keyへ`EXLSXS.Host.exe --update-check`を登録する。uninstall callbackは起動時登録とVSTO登録を解除する。

### サイレント自動更新

1. Windows起動時にホストが`--update-check`で実行され、VSTO登録を保守してから更新確認を始める。この経路では前提条件installerを起動せず、予期しないUAC表示を避ける。
2. `UpdateSettings`が`appsettings.json`を読み、その後`EXLSXS_UPDATE_*`環境変数で上書きする。更新元が未設定、またはVelopack install外なら確認を終了する。
3. `UpdateChecker`が設定に応じて通常URLまたはGitHub sourceを選び、timeout付きで更新確認とdownloadを行う。
4. download後、`UpdatePackageTrustVerifier`が`EXLSXS.Host.exe`、`EXLSXS.Host.dll`、`EXLSXS.dll`の署名、証明書chain、期待thumbprintとsubjectを検証する。
5. 信頼検証に成功したパッケージだけを`ApplyUpdatesAndExit`へ渡す。未信頼または検証設定なしのパッケージは適用しない。

### パッケージ作成と公開

1. `Directory.Build.props`の`Version`からVSTO、ホスト、Velopack成果物のversionを導出する。
2. VSTO publish成果物を`.vsto`、deployment manifest、アセンブリ、依存DLLが同じ階層にある形へ整え、ホスト成果物とともにVelopack stagingへ配置する。
3. `vpk`がinstaller、portable package、release manifestを生成する。VelopackライブラリとCLIは同じ検証済み安定版を使う。
4. release scriptが署名を再検証してR2へuploadし、URLが固定された成果物のCloudflare cacheをpurgeする。公開manifestがローカル成果物と一致してから旧世代を整理する。
5. Cloudflare Workerはランディングページだけを処理し、manifest、package、installerなどの配信特性はR2へ委ねる。

## 状態と外部境界

| 状態・外部系 | 正本または境界 |
| --- | --- |
| 製品version | `Directory.Build.props`の`Version` |
| リボンの表示モード・倍率 | `HKCU\Software\EXLSXS` |
| VSTO登録 | HKCUのExcel add-in key（通常viewと`WOW6432Node`） |
| 起動時更新確認 | HKCUのWindows Run key |
| 更新設定 | 配布物の`appsettings.json`を基底に、`EXLSXS_UPDATE_*`環境変数を優先 |
| 更新成果物 | R2 bucket `exlsxs-updates`とVelopack release manifest |
| 公開入口 | `exlsxs.kagayoi.com`。移行期限までは`exlsxs.nephilim.jp`も同じWorker routeで維持 |

## 重要な不変条件

- VSTO stagingはフラット構成とし、publish時は`MapFileExtensions=false`を明示する。`.vsto`とロード対象DLLの物理位置を分離しない。
- VSTOのExcel host宣言とdesigner生成物、および`.slnx`のType属性なし構成を維持する。
- COM eventは埋め込みinterop metadataと互換な`+=`で購読する。
- install時だけ前提条件の自動導入と`LoadBehavior`強制を許可し、update・通常起動ではユーザーまたはOfficeが無効化した状態を保持する。
- 更新パッケージはpublisher trust設定と必須3ファイルの署名検証を通過するまで適用しない。
- リリースの固定名ファイルを更新した後はCloudflare cacheをpurgeし、公開manifestとの一致を確認する。
- 旧`nephilim.jp`の更新配信経路は2027-05-31まで維持し、配信ファイルをroot redirectへ巻き込まない。

## 採用済みの設計判断

- **VSTO本体とVelopackホストの分離**: Excel統合は.NET Framework/VSTOへ残し、installer・更新・署名検証は.NET 10ホストへ集約する。実行環境は二重になるが、VSTO互換性と現行の配布機能を両立できる。
- **ユーザー単位の登録**: 設定、VSTO登録、起動時更新をHKCUへ置く。端末全体への管理者権限を不要にする代わりに、Windows userごとに登録状態を持つ。
- **シート単位のbest-effort処理**: COM例外や保護・非表示状態でブック全体を中断せず、適用可能なシートを処理する。部分適用になり得るため、保護シートでは書式変更を明示的に避ける。
- **Worker routeとR2 custom domainの重ね合わせ**: 同じhostで案内ページと更新配信を提供し、非root pathはR2へ透過委譲する。構成は簡潔になる一方、Workerは更新ファイルのRange、cache、Content-Typeを変更しないことが前提となる。
- **versionと配布toolの固定**: 製品versionを一箇所へ集約し、WranglerとVelopack CLIは検証済みversionへ固定する。自動追随より再現可能な署名・配布を優先する。
