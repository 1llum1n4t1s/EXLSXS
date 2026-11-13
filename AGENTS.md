# AGENTS.md — EXLSXS

This file provides guidance to Codex and other coding agents working in this repository.

Excel 用 VSTO アドイン（リボンから全シートの表示倍率・表示モード・フォント・選択位置を一括整形）を Velopack で配布するプロジェクト。

## 構成

- `EXLSXS/` — VSTO アドイン本体 (.NET Framework 4.8.1・legacy csproj)
- `EXLSXS.Host/` — Velopack ホスト (.NET 11)。インストール/更新時の VSTO 登録・起動時サイレント更新
- `build/pack-velopack.ps1` — VSTO publish → host publish → Velopack pack（staging はフラット構成必須、下記参照）
- `scripts/release-local.ps1` — 署名付きローカルリリース（ビルド → 署名 → 検証 → R2 アップロード）
- `../vps-web/lp/exlsxs/` — VPS配信のランディングページ（`exlsxs.kagayoi.com`）
- `Directory.Build.props` — **バージョンの唯一の定義場所**（`<Version>`、他は全部ここから導出）
- `DESIGN.md` — 現行システムの構造、責務、データフロー、設計判断の正本

## 主要コマンド

```powershell
# ビルド
MSBuild EXLSXS.slnx /t:Restore,Rebuild /p:Configuration=Release /p:Platform="Any CPU"
dotnet build EXLSXS.Host/EXLSXS.Host.csproj -c Release

# ホストのテスト
dotnet test EXLSXS.Host.Tests/EXLSXS.Host.Tests.csproj -c Release

# 署名付きパック（R2 アップロード無し）
pwsh -NoProfile -File scripts/release-local.ps1 -SkipUpload

# リリース（署名 + R2 アップロード）。通常は GitHub Actions の release.yml を手動起動する
gh workflow run release.yml --ref main
# ローカルで同じ手順を実行する場合（Sigilune 1.1.0 以降と `sigilune setup` 済みが前提）
pwsh -NoProfile -File scripts/release-local.ps1

# 実機検証済み成果物の配信（変数には実機検証時に記録した full.nupkg の SHA-256 を設定）
pwsh -NoProfile -File scripts/release-local.ps1 -SkipBuild -ExpectedPackageSha256 $verifiedPackageSha256

# ランディングページのデプロイ（リポジトリルートから EXLSXS だけを指定）
pwsh -NoProfile -File ../vps-web/deploy/deploy-lp.ps1 -Product exlsxs
```

## 守ること（実機で踏んだ罠）

- **vsto staging はフラット構成を維持する**: `.vsto` / `.dll.manifest` / `EXLSXS.dll` / 依存 DLL を同一フォルダに並べる。vstolocal 登録は .vsto と同じフォルダを AppBase にするため、ClickOnce のネスト構成 (`Application Files\<name>_<ver>\`) のまま登録すると `Assembly.Load("EXLSXS")` が FileNotFoundException でアドインが読み込めない
- **publish は `MapFileExtensions=false` をコマンドライン `/p:` で渡す**: csproj の PropertyGroup 値は VSTO publish ターゲットに上書きされる。`.deploy` 拡張子が付くと vstolocal がファイルを見つけられない
- **署名後に VSTO マニフェストを再生成する**: `build/pack-velopack.ps1` の「Rebuild → DLL の Authenticode 署名 → `/t:ResolveReferences,VisualStudioForApplicationsBuild,PublishOnly`」の順序を維持する。`PublishOnly` 単独へ省略せず、フラットな `.vsto` は同じ `$vstoBuildDir` から取得する。梱包前の版数・identity・ハッシュ・DLL サイズ検証を維持し、配布処理を変更したら署名付き pack（`-SkipUpload`）がこの検証まで通ることを確認する。理由と検証範囲は [DESIGN.md](DESIGN.md#パッケージ作成と公開) を参照。
- **実機検証済み成果物を配信するときは再ビルドしない**: `-SkipBuild -ExpectedPackageSha256` を使い、`artifacts/velopack-releases/EXLSXS-<Version>-full.nupkg` と実機検証時に記録した SHA-256 を照合する。配信直前に計算した値で検証記録を代用しない。この経路でも署名証明書のプリフライトと成果物の署名検証は必要。
- **VSTO ランタイム検出は `v4` と `v4R` の両キーを見る**: Office/VS 同梱導入は `v4`、再頒布パッケージは `v4R` に登録される
- **EmbedInteropTypes=True のため COM イベントは `+=` で購読する**: 文字列ベースの `ComAwareEventInfo` はイベントメタデータが埋め込まれず NullReferenceException になる
- **バージョンは `Directory.Build.props` の `<Version>` だけを更新する**: csproj の `ApplicationVersion` / `AssemblyVersion` / host の版数はすべて導出
- **Velopack ライブラリと vpk CLI は同じ安定版に揃える**: ライブラリ版は `EXLSXS.Host/EXLSXS.Host.csproj`、CLI 版は `build/pack-velopack.ps1` の `$vpkVersion` が正本。片方だけ更新せず、同じ batch で更新して pack 動作を確認する。pack 冒頭と CI の版一致検査が、片方だけの更新を止める。更新時は、Update.exe の展開処理（`lib-rust/src/bundle.rs` と zip クレートの、エントリ名や EOCD の扱い）が `EXLSXS.Host/ZipEntryNameConsistency.cs` と `HasAmbiguousWindowsEntryName` の前提から変わっていないかを確認する。あわせて、新しい pack 出力がこの検査で拒否されないことも確かめる（コメント付き EOCD・Zip64・拡張フィールドが出るようになると、全クライアントで更新が止まる）
- **VS2026 で開くには (1) csproj に標準 VSTO デザイナー構成を持たせ、(2) `.slnx` の VSTO プロジェクト行に `Type` 属性を付けない**: VS2026 でも VSTO は正式サポート（公式テンプレートが同梱され、新規テンプレートは正常にロードする）。EXLSXS が読み込めなかった原因は 2 つあり両方を満たす必要がある。
  - **(1) csproj 側**: `<ProjectExtensions>` の `<FlavorProperties GUID="{BAA0C2D2-18E2-41B9-852F-F413020CAA33}">` に `<ProjectProperties>` と `<Host Name="Excel" GeneratedCodeNamespace="EXLSXS"><HostItem ... Blueprint="ThisAddIn.Designer.xml" GeneratedCode="ThisAddIn.Designer.cs" /></Host>` を持たせ、`ThisAddIn.Designer.xml`（Blueprint）と `ThisAddIn.Designer.cs`（生成コード相当: `partial class ThisAddIn` の生成メンバ + `Globals` + `ThisRibbonCollection`）を実体として置く。`ThisAddIn.cs` はユーザーコードのみにする。この `<Host>`/`<HostItem>` 宣言が無いとフレーバー初期化がアサート (`GUID が空です。 パラメーター名:serviceGuid`) で失敗する（MSBuild ビルドは Host 宣言が無くても通るため気付きにくい）。
  - **(2) `.slnx` 側**: VSTO プロジェクト行は `<Project Path="EXLSXS/EXLSXS.csproj" />` と **Type 属性なし** で書く。Type 属性なしだと VS は拡張子で C# 基底型を判定し csproj の `<ProjectTypeGuids>`（`{BAA0C2D2};{FAE04EC0}` の 2 段チェーン）を読んでフレーバーをアグリゲートする。`Type="{BAA0C2D2-...}"` を付けると**外側フレーバー GUID だけ**が指定され基底型が欠けてアグリゲーションが壊れ「読み込みに失敗しました」になる（過去にこの Type 属性を回避策として入れていたが逆効果だった）。
  - 検証: `EXLSXS.csproj` を直接開く / `.slnx`（Type 無し）を開く のどちらでも `ソリューション 'EXLSXS' (3/3 のプロジェクト)` で EXLSXS が Excel ホストノード付きでロードされること。SDK のホスト/テストは Type 属性不要。
- **リボン/コードを変えたら `%LOCALAPPDATA%\assembly\dl3` を消してから Excel を起動して動作確認する**: EXLSXS は strong-name 付きで、同じ `<Version>` から生成した `$(Version).0` のビルドは同じアセンブリ identity を持つ。VSTO/Fusion は**同一バージョンの新ビルドを「同じ物」とみなして `dl3` の旧コピーを読み続ける**ため、bin\Debug/Release を再ビルドしても Excel に反映されない。dev 反復での確認手順は「Excel 終了 → `%LOCALAPPDATA%\assembly\dl3` のキャッシュ削除 → Excel 起動」。Debug と Release は identity が同一なので Fusion がどちらのコピーを使うか不定 → 確実を期すなら両構成を再ビルドしてからキャッシュを消す。リリース時は `/vava` でバージョンが上がり identity が変わるため、この罠はエンドユーザーには出ない（dev 専用）
- 製品ページの配信は `vps-web/deploy/deploy-lp.ps1` を使う。公開ホスト・更新ファイルの既存経路を維持する。
- **Velopack 配信の固定名ファイルはアップロード後に Cloudflare キャッシュをパージする**: `EXLSXS-win-Setup.exe` / `releases.win.json` / `RELEASES` / `EXLSXS-win-Portable.zip` / `assets.win.json` は URL 不変で毎リリース中身が変わる。`scripts/release-local.ps1` は新旧両ホストの固定名 URL に対する `purge_cache` API の受理を確認する。バージョン付き nupkg は URL が一意なのでパージ不要。手動アップロード時も固定名 URL をパージする。

## ドメイン移行（2026-07 開始・期限 2027/05/31）

配信ドメインを `nephilim.jp` から `kagayoi.com` へ移行中。公開入口の構造は [DESIGN.md](DESIGN.md#状態と外部境界) を参照する。

- **旧ドメイン `nephilim.jp` はレジストラで廃止申請済みで 2027/05/31 に失効する**（延長しない）。それまでに出荷済みバイナリを新ドメインへ移行しきる。
- 旧ホストの Worker route / custom domain は**期限まで消さない**。消すと出荷済みアプリの自動更新が止まる。
- `nephilim.jp` の Redirect Rules は `/` だけを 301 する。`releases.*.json` / `*.nupkg` / `*-Setup.exe` は転送せず R2 が配信を続ける。
- 配信は `exlsxs.kagayoi.com`（R2 `exlsxs-updates`）。旧 `exlsxs.nephilim.jp` は route に併記して残してある。

## 製品ページの配信先

製品ページのHTML・gateway・service bindingの構造は [DESIGN.md](DESIGN.md#製品ページの配信先) を参照する。配信には「主要コマンド」の `deploy-lp.ps1 -Product exlsxs` を使い、公開URLと既存のR2・ライセンス通信を維持する。

### ローカルリリースの実行

署名は Sigilune（Certum のクラウド鍵）で行い、SimplySign Desktop は使わない。Authenticode 署名に加え、VSTO の `EXLSXS.dll.manifest` / `EXLSXS.vsto` も `sigilune sign-manifest` で署名する（MSBuild の publish はこの実行だけの使い捨て証明書で行い、ステージング後に署名し直す）。公開依頼時は GitHub Actions の `release.yml` を手動起動すると、承認待ちなしで実行される（Environment `release` は承認者なし・ブランチ制限だけ）。ローカルで実行するときは、先に `sigilune status` で登録を確認してから実行する。

```powershell
sigilune status
if ($LASTEXITCODE -eq 0) { pwsh -NoProfile -File scripts/release-local.ps1 } else { throw 'Sigilune の登録が見つからないよ。sigilune setup で Certum アカウントを登録してから /vava を再実行してね' }
```
