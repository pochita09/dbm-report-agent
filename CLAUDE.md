# CLAUDE.md

MODE: personal

---

## プロジェクト概要

DBM報告書に現場写真を自動配置するWebアプリ。  
ExcelテンプレートをアップロードしてGemini APIで写真を分類、openpyxlで配置したExcelを返す。  
社内ツール（単一ユーザー想定）。本番デプロイ先はRender。

---

## 技術スタック

| 層 | 技術 |
|----|------|
| Webフレームワーク | Flask + Flask-Limiter |
| 非同期I/O | gunicorn + gevent（本番のみ） |
| 画像AI | Google Gemini API（HTTPダイレクト呼び出し） |
| Excel操作 | openpyxl |
| デプロイ | Render（`render.yaml` で設定済み） |
| リアルタイム通知 | SSE（Server-Sent Events） |

主要ファイル:
- `app.py` — Flask サーバー、認証、SSEエンドポイント
- `classify_photos.py` — Gemini APIによる写真分類ロジック
- `place_photos.py` — Excelへの写真配置ロジック

---

## 重要な制約（Gotchas）

### workers=1 は絶対に変更しないこと

`app.py` のモジュールレベル辞書 `original_names = {}` はプロセス間で共有されない。  
`--workers 2` 以上にするとダウンロード時のファイル名がデフォルト値に化ける。  
詳細は `docs/02_decisions.md` ADR-001 を参照。

### APP_PASSWORD は必ず設定すること

未設定（空文字列）の場合、認証がスキップされてインターネットに無防備に公開される。  
Renderダッシュボードで `APP_PASSWORD` / `SECRET_KEY` / `GEMINI_API_KEY` を設定済みか確認すること。

### ダウンロードは1回限り

ダウンロード処理後にサーバー上のExcelファイルは即削除される（`app.py:267-268`）。  
UIにその旨を表示していないため、ユーザーへの周知が必要。

### ローカル実行 vs 本番の差異

| 項目 | ローカル (`python app.py`) | 本番 (Render + gunicorn) |
|------|--------------------------|--------------------------|
| サーバー | Flask開発サーバー | gunicorn + gevent |
| geventモンキーパッチ | 適用されない | 適用される |
| workers | 1（Flask単一スレッド） | 1（gunicorn固定） |

---

## 環境変数

`.env` ファイル（`.gitignore`対象）に記載。  
`.env.example` を参照してコピーすること。

本番（Render）ではダッシュボードで直接管理。`render.yaml` に `sync: false` で定義済み。

---

## コミット前チェックリスト

- [ ] `.env` / APIキー / パスワードがコードや設定ファイルに含まれていないか
- [ ] `render.yaml` の `--workers` が `1` のままか
- [ ] `classify_photos.py` の `__main__` ブロックに個人PCのパスが追加されていないか
- [ ] `requirements.txt` に新しいパッケージを追加した場合、バージョンを固定したか
- [ ] 設計判断を行った場合、`docs/02_decisions.md` に追記したか

---

## 既知の未対応事項（SEレビュー指摘）

優先度「高」の未対応:
- `render.yaml` に `SECRET_KEY` / `APP_PASSWORD` の `sync: false` エントリが未追加（H-1）
- 複数タブ同時処理による二重API呼び出し防止が未実装（H-4）

詳細は `docs/se-review-report.md` を参照。
