# 変更履歴

Git のバージョン記録・コミット差分と既存の変更履歴をもとに、確認できた版ごとの変更点をまとめています。「Git 記録日」は公開日ではありません。番号の欠番だけから未確認のリリースは補っていません。

## 未リリース

## [1.2.0] — Git 記録日: 2026-04-23

- 配布ホスト EXLSXS.Host の初期版。VSTO 登録・前提条件確認・Velopack による更新基盤を追加。

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/96013f2cb40788ff27c3a7467315e8f883471ff9) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/b5a13c76d5966696b95d2a2ddd364b4bba3728bf...96013f2cb40788ff27c3a7467315e8f883471ff9)。

## [1.0.10] — Git 記録日: 2026-09-01

- アプリアイコンの透過境界を調整。

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/12d2844878133523654f94898346693478d302e1) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/ca856de4957b10d7907ba4b70837e014d3219ccb...12d2844878133523654f94898346693478d302e1)。

## [1.0.9] — Git 記録日: 2026-08-31

- アプリアイコンをクリスタル調に刷新する
- 新旧配信ホストの反映確認を強化する
- Excel整形中のイベントと再計算を抑制する
- インストール後のVSTO登録を再試行する
- Velopack依存の更新契約を修正する
- xUnit 4とMTPへ移行する

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/ca856de4957b10d7907ba4b70837e014d3219ccb) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/d0a0d16c41948025d1f1392499e091c320ceb488...ca856de4957b10d7907ba4b70837e014d3219ccb)。

## [1.0.8] — Git 記録日: 2026-08-30

- 自動更新とVSTO処理を堅牢化する
- 配信ドメインを nephilim.jp から kagayoi.com へ移行

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/d0a0d16c41948025d1f1392499e091c320ceb488) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/ab155a1f2539f6f335633deda3246cf27c8e4629...d0a0d16c41948025d1f1392499e091c320ceb488)。

## [1.0.7] — Git 記録日: 2026-07-02

- Excel のウィンドウ表示・ズーム・フォント取得を保護し、一部の COM 操作が失敗しても整形全体が中断しないよう改善。

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/ab155a1f2539f6f335633deda3246cf27c8e4629) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/120118cdc7badd97f07e0e6e6f2203d1744e30b8...ab155a1f2539f6f335633deda3246cf27c8e4629)。

## [1.0.6] — Git 記録日: 2026-07-01

- セルの方眼紙化・表示形式の一括設定機能を追加
- R2 アップロード後に Cloudflare キャッシュをパージして旧バージョン配信を防ぐ

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/120118cdc7badd97f07e0e6e6f2203d1744e30b8) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/e8828203d6f418d1075ffc9cfcfbb8cb406f7113...120118cdc7badd97f07e0e6e6f2203d1744e30b8)。

## [1.0.5] — Git 記録日: 2026-06-18

- ウィンドウ枠固定シートのスクロールが左上に戻らない不具合を修正
- ウィンドウ枠固定シートのスクロールを各シート左上へ戻す

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/e8828203d6f418d1075ffc9cfcfbb8cb406f7113) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/0362045b040f7689104e1c0d54d1e3b24b5f5be2...e8828203d6f418d1075ffc9cfcfbb8cb406f7113)。

## [1.0.4] — Git 記録日: 2026-06-15

- 仕上げ後に先頭シートをアクティブ化・設定永続化・LP モーション刷新

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/0362045b040f7689104e1c0d54d1e3b24b5f5be2) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/c206a5dc0c7e1df976c38292fe5854069644683f...0362045b040f7689104e1c0d54d1e3b24b5f5be2)。

## [1.0.3] — Git 記録日: 2026-06-14

- VS2026 で EXLSXS が読み込めるよう標準デザイナー構成を復元 + UI/コード整理
- vpk CLI を動作確認済みの 1.2.0 に固定（ライブラリ版数依存の誤りを修正）

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/c206a5dc0c7e1df976c38292fe5854069644683f) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/63079e4f5f3256e0a07fe70ab33fad09e88e2ef4...c206a5dc0c7e1df976c38292fe5854069644683f)。

## [1.0.2] — Git 記録日: 2026-06-13

- 証明書の年次更新で自動更新が止まる問題を修正 (thumbprint + subject + 証明書チェーン検証)
- 前提ブートストラッパー setup.exe の署名検証を追加 (差し替えた exe の昇格起動を防止)
- vpk CLI をライブラリと同一バージョンに固定 / PFX パスワードを SecureString 化
- DoFinish 後に元のアクティブシートへ復帰 / フォントプレビューを軽量化して起動を高速化
- リリースの旧 nupkg を 2 世代保持してロールバック可能に
- 登録例外が Velopack インストール/更新を巻き込まないよう隔離

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/63079e4f5f3256e0a07fe70ab33fad09e88e2ef4) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/616874b93b86962be92b2ef030678c8d78e84795...63079e4f5f3256e0a07fe70ab33fad09e88e2ef4)。

## [1.0.1] — Git 記録日: 2026-06-13

- VSTO アドインが読み込めない不具合を修正
- 配信元を Cloudflare R2 に移行しローカル署名リリースを構築

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/616874b93b86962be92b2ef030678c8d78e84795) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/96013f2cb40788ff27c3a7467315e8f883471ff9...616874b93b86962be92b2ef030678c8d78e84795)。

## [1.0.0] — Git 記録日: 2026-01-14

- アプリ名を EXLSXS に変更し、Excel アドインの画面と COM 連携を再構成。

出典: [版の記録](https://github.com/1llum1n4t1s/EXLSXS/commit/b5a13c76d5966696b95d2a2ddd364b4bba3728bf) / [変更差分](https://github.com/1llum1n4t1s/EXLSXS/compare/d6005086ae26a9ab4c757e40ba3cf6dfebee86f0...b5a13c76d5966696b95d2a2ddd364b4bba3728bf)。
