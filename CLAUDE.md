# DBM Report Agent

清掃作業の写真をExcelテンプレートに自動配置するWebアプリ。Gemini AIで写真を分類し、openpyxlでExcelに貼り付ける。

## アーキテクチャ

```
Browser → Flask (app.py) → classify_photos.py → Gemini API
                        → place_photos.py     → openpyxl → Excel出力
```

**SSEで進捗をリアルタイム配信**（アップロード→AI分類→Excel配置→ダウンロード）

## ファイル構成

| ファイル | 役割 |
|---|---|
| `app.py` | Flaskサーバー。認証・アップロード・SSE・ダウンロード |
| `classify_photos.py` | Gemini API呼び出し、写真↔スロットのマッピング |
| `place_photos.py` | Excelスロット検出、写真配置（EMU計算含む） |
| `analyze_template.py` | デバッグ用テンプレート解析ツール |
| `static/index.html` | メインUI（ドラッグ&ドロップ、SSE受信） |
| `static/login.html` | パスワード認証UI |

## 開発環境セットアップ

```bash
pip install -r requirements.txt
cp .env.example .env  # GEMINI_API_KEYを設定
python app.py         # http://localhost:5000
```

**必須環境変数（`.env`）:**
- `GEMINI_API_KEY` — Google Gemini APIキー（必須）
- `APP_PASSWORD` — パスワード認証（省略時は無認証で公開）
- `SECRET_KEY` — Flaskセッションキー（省略時は起動ごとに再生成→再起動でセッション無効）

## 主要な処理フロー

### 写真分類（classify_photos.py）

1. `detect_photo_slots(ws)` — 行高さ・列幅の中央値からスロットを検出
2. `parse_slot_info(slots)` — section/categoryから作業内容・状態を分離
3. `build_prompt(parsed_slots, filenames)` — Geminiへのプロンプト構築（基本層＋ガラス清掃補足層）
4. `call_gemini_api(prompt, photo_paths, api_key)` — 全写真を一括送信、JSON受信
5. `assign_photos(api_response, parsed_slots, photo_paths)` — スロット順リストに変換

### スロット検出のロジック（place_photos.py）

- **行高さ** の中央値 × 3.0 以上の行 = 写真行
- **列幅** の中央値 × 1.2 以上の列 = コンテンツ列
- 交点がスロット。上方向に最大3行ずつ category → section を探索

### 座標系

Excel画像配置はEMU（English Metric Units）で計算:
- `col_width_to_emu()` — Excel列幅文字数 → EMU
- `row_height_to_emu()` — 行高さpt → EMU
- `OneCellAnchor` + `AnchorMarker` でセル内中央配置

## 本番環境

| 項目 | 内容 |
|---|---|
| ホスティング | Render.com 無料プラン |
| 利用者 | 社内5名程度 |
| 利用頻度 | 1日1〜3件 |
| データ規模 | 1件あたり画像約20枚、Excelテンプレート1ファイル |

**無料プランの制約に注意:**
- アイドル15分でスリープ → 最初のリクエストに50秒程度かかる
- 一時ファイルはサーバー再起動（スリープ復帰含む）でリセットされる（処理中のデータは消える可能性あり）
- ディスク容量・メモリ制限があるため、セッションファイルの放置は避けること（現状クリーンアップ未実装）
- Gemini API呼び出し（画像20枚）＋Excel生成で30〜60秒かかるため、timeout=300は必要

## デプロイ（Render.com）

`render.yaml`設定済み。gevent×1ワーカー（SSE対応）。

```bash
# ワーカー1つ固定（geventでSSE対応、複数ワーカーはSECRET_KEY共有問題に注意）
gunicorn app:app --worker-class gevent --workers 1 --timeout 300
```

**Render環境変数として必ず設定:** `GEMINI_API_KEY`, `SECRET_KEY`

## コーディング規約

- 日本語コメント・ログで統一
- 関数ごとにdocstringを書く（Args/Returns形式）
- エラーはユーザー向けメッセージ（日本語）でSSE `error_event` に流す
- `secure_filename()` を必ずサーバー保存パスに使用し、元ファイル名は別途保持
- 画像はAPI送信前に1024px以内にリサイズ（`encode_photo()`）

## よくある落とし穴

- `SECRET_KEY`未設定のままデプロイ→再起動のたびにセッション無効
- gunicornワーカーを複数にすると `original_names` dict が共有されない（各ワーカーが独立）
- Excelの結合セルは `get_slot_merged_range()` で正確なサイズを取得すること
- MDW（Maximum Digit Width）はフォントごとに異なる → `detect_mdw()` で自動取得

## テスト実行

```bash
# 単体テスト（テスト用パスをclassify_photos.pyの__main__ブロックに設定して実行）
python classify_photos.py

# テンプレート解析デバッグ
python analyze_template.py
```
