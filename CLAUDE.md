# EXLSXS

Excel 用 VSTO アドイン（リボンから全シートの表示倍率・表示モード・フォント・選択位置を一括整形）を Velopack で配布するプロジェクト。

## 構成

- `EXLSXS/` — VSTO アドイン本体 (.NET Framework 4.8.1・legacy csproj)
- `EXLSXS.Host/` — Velopack ホスト (.NET 10)。インストール/更新時の VSTO 登録・起動時サイレント更新
- `build/pack-velopack.ps1` — VSTO publish → host publish → Velopack pack（staging はフラット構成必須、下記参照）
- `scripts/release-local.ps1` — 署名付きローカルリリース（ビルド → 署名 → 検証 → R2 アップロード）
- `web/` — ランディングページ + Cloudflare Worker（`exlsxs.nephilim.jp`）
- `Directory.Build.props` — **バージョンの唯一の定義場所**（`<Version>`、他は全部ここから導出）

## 主要コマンド

```powershell
# ビルド
MSBuild EXLSXS.slnx /t:Restore,Rebuild /p:Configuration=Release /p:Platform="Any CPU"
dotnet build EXLSXS.Host/EXLSXS.Host.csproj -c Release

# 署名付きパック（R2 アップロード無し）
pwsh -NoProfile -File scripts/release-local.ps1 -SkipUpload

# リリース（署名 + R2 アップロード。SimplySign ログイン必須）
pwsh -NoProfile -File scripts/release-local.ps1

# ランディングページのデプロイ（リポジトリルートから実行・--config でファイル明示が必須）
pnpm dlx wrangler@4 deploy --config web/wrangler.toml
```

## 守ること（実機で踏んだ罠）

- **vsto staging はフラット構成を維持する**: `.vsto` / `.dll.manifest` / `EXLSXS.dll` / 依存 DLL を同一フォルダに並べる。vstolocal 登録は .vsto と同じフォルダを AppBase にするため、ClickOnce のネスト構成 (`Application Files\<name>_<ver>\`) のまま登録すると `Assembly.Load("EXLSXS")` が FileNotFoundException でアドインが読み込めない
- **publish は `MapFileExtensions=false` をコマンドライン `/p:` で渡す**: csproj の PropertyGroup 値は VSTO publish ターゲットに上書きされる。`.deploy` 拡張子が付くと vstolocal がファイルを見つけられない
- **VSTO ランタイム検出は `v4` と `v4R` の両キーを見る**: Office/VS 同梱導入は `v4`、再頒布パッケージは `v4R` に登録される
- **EmbedInteropTypes=True のため COM イベントは `+=` で購読する**: 文字列ベースの `ComAwareEventInfo` はイベントメタデータが埋め込まれず NullReferenceException になる
- **バージョンは `Directory.Build.props` の `<Version>` だけを更新する**: csproj の `ApplicationVersion` / `AssemblyVersion` / host の版数はすべて導出
- **Velopack パッケージは prerelease 固定** (`0.0.1369-g1d5c984`): 安定版への更新は要手動判断（vpk CLI とのバージョン整合を確認してから）
- **VS2026 で開くには (1) csproj に標準 VSTO デザイナー構成を持たせ、(2) `.slnx` の VSTO プロジェクト行に `Type` 属性を付けない**: VS2026 でも VSTO は正式サポート（公式テンプレートが同梱され、新規テンプレートは正常にロードする）。EXLSXS が読み込めなかった原因は 2 つあり両方を満たす必要がある。
  - **(1) csproj 側**: `<ProjectExtensions>` の `<FlavorProperties GUID="{BAA0C2D2-18E2-41B9-852F-F413020CAA33}">` に `ProjectCreationSetting="1"` 付き `<ProjectProperties>` と `<Host Name="Excel" GeneratedCodeNamespace="EXLSXS"><HostItem ... Blueprint="ThisAddIn.Designer.xml" GeneratedCode="ThisAddIn.Designer.cs" /></Host>` を持たせ、`ThisAddIn.Designer.xml`（Blueprint）と `ThisAddIn.Designer.cs`（生成コード相当: `partial class ThisAddIn` の生成メンバ + `Globals` + `ThisRibbonCollection`）を実体として置く。`ThisAddIn.cs` はユーザーコードのみにする。この `<Host>`/`<HostItem>` 宣言が無いとフレーバー初期化がアサート (`GUID が空です。 パラメーター名:serviceGuid`) で失敗する（MSBuild ビルドは Host 宣言が無くても通るため気付きにくい）。
  - **(2) `.slnx` 側**: VSTO プロジェクト行は `<Project Path="EXLSXS/EXLSXS.csproj" />` と **Type 属性なし** で書く。Type 属性なしだと VS は拡張子で C# 基底型を判定し csproj の `<ProjectTypeGuids>`（`{BAA0C2D2};{FAE04EC0}` の 2 段チェーン）を読んでフレーバーをアグリゲートする。`Type="{BAA0C2D2-...}"` を付けると**外側フレーバー GUID だけ**が指定され基底型が欠けてアグリゲーションが壊れ「読み込みに失敗しました」になる（過去にこの Type 属性を回避策として入れていたが逆効果だった）。
  - 検証: `EXLSXS.csproj` を直接開く / `.slnx`（Type 無し）を開く のどちらでも `ソリューション 'EXLSXS' (3/3 のプロジェクト)` で EXLSXS が Excel ホストノード付きでロードされること。SDK のホスト/テストは Type 属性不要。
- **リボン/コードを変えたら `%LOCALAPPDATA%\assembly\dl3` を消してから Excel を起動して動作確認する**: EXLSXS は strong-name 付き + バージョン固定 (`1.0.2.0`) なので、VSTO/Fusion はアセンブリ identity でシャドウコピーをキャッシュし、**同一バージョンの新ビルドを「同じ物」とみなして `dl3` の旧コピーを読み続ける**（bin\Debug/Release を再ビルドしても Excel に反映されない）。dev 反復での確認手順は「Excel 終了 → `rm -rf %LOCALAPPDATA%/assembly/dl3` → Excel 起動」。Debug と Release は identity が同一なので Fusion がどちらのコピーを使うか不定 → 確実を期すなら両構成を再ビルドしてからキャッシュを消す。リリース時は `/vava` でバージョンが上がり identity が変わるため、この罠はエンドユーザーには出ない（dev 専用）
- **ランディングページのデプロイは `pnpm dlx wrangler@4 deploy --config web/wrangler.toml` をリポジトリルートから実行する**: `pnpm -C web dlx wrangler@4 deploy` は `-C` が効かず CWD がリポジトリルートのまま wrangler が起動し、`web/wrangler.toml` を読まずに「assets 配信 Worker」を `my-worker` 名で `*.workers.dev` に誤デプロイする（同時にルートへ `wrangler.jsonc` / `.wrangler/` を生成し `.gitignore` を書き換える）。`--config web/wrangler.toml` でファイルを明示すると設定ディレクトリ基準で動き、`exlsxs-landing` を `exlsxs.nephilim.jp/*` route に正しく載せられる。本番前に `--dry-run` で `Total Upload` が ~64 KiB（`index.html` が Text モジュールでバンドルされた証拠）になることを確認する（0.4 KiB なら誤った assets モードに落ちている）。`git push` は GitHub を更新するだけで本番サイトは変わらないため、`web/index.html` を変えたらこのコマンドで明示デプロイする
