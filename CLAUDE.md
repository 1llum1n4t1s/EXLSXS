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

# ランディングページのデプロイ
pnpm -C web dlx wrangler@4 deploy
```

## 守ること（実機で踏んだ罠）

- **vsto staging はフラット構成を維持する**: `.vsto` / `.dll.manifest` / `EXLSXS.dll` / 依存 DLL を同一フォルダに並べる。vstolocal 登録は .vsto と同じフォルダを AppBase にするため、ClickOnce のネスト構成 (`Application Files\<name>_<ver>\`) のまま登録すると `Assembly.Load("EXLSXS")` が FileNotFoundException でアドインが読み込めない
- **publish は `MapFileExtensions=false` をコマンドライン `/p:` で渡す**: csproj の PropertyGroup 値は VSTO publish ターゲットに上書きされる。`.deploy` 拡張子が付くと vstolocal がファイルを見つけられない
- **VSTO ランタイム検出は `v4` と `v4R` の両キーを見る**: Office/VS 同梱導入は `v4`、再頒布パッケージは `v4R` に登録される
- **EmbedInteropTypes=True のため COM イベントは `+=` で購読する**: 文字列ベースの `ComAwareEventInfo` はイベントメタデータが埋め込まれず NullReferenceException になる
- **バージョンは `Directory.Build.props` の `<Version>` だけを更新する**: csproj の `ApplicationVersion` / `AssemblyVersion` / host の版数はすべて導出
- **Velopack パッケージは prerelease 固定** (`0.0.1369-g1d5c984`): 安定版への更新は要手動判断（vpk CLI とのバージョン整合を確認してから）
