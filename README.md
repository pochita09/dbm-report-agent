# dbm-report-agent

DBM（データベースマーケティング）報告書に、現場写真を自動配置するWebアプリ。  
Excelテンプレートをアップロードし、写真をAI（Gemini）で分類・配置したExcelをダウンロードできる。

---

## 動作要件

- Python 3.12
- 本番デプロイ先: [Render](https://render.com)（`render.yaml` に設定済み）
- ローカル実行はgunicornなしのFlask開発サーバーで可能（後述）

---

## 必要な環境変数

| 変数名 | 用途 | 備考 |
|--------|------|------|
| `GEMINI_API_KEY` | Google Gemini API（写真分類） | Google AI StudioまたはVertex AIで発行 |
| `APP_PASSWORD` | Webアプリへのログインパスワード | **未設定時は認証がスキップされる** — 必ず設定すること |
| `SECRET_KEY` | Flaskセッション署名用 | 未設定時は起動のたびに再生成され、全セッションが無効化される |

> **注意:** 本番（Render）では `render.yaml` の `sync: false` エントリに対し、  
> Renderダッシュボードの「Environment」から値を手動で設定すること。

---

## ローカル実行手順

```bash
# 1. リポジトリをクローン
git clone <repo-url>
cd dbm-report-agent

# 2. 仮想環境を作成・有効化
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 3. 依存パッケージをインストール
pip install -r requirements.txt

# 4. 環境変数ファイルを作成
copy .env.example .env   # Windows
# cp .env.example .env   # macOS/Linux
# .env を編集して各変数を設定する

# 5. 開発サーバーを起動
python app.py
```

ブラウザで `http://localhost:5000` にアクセスする。  
ローカル実行時はgunicorn/geventは使われない（Flask開発サーバーのみ）。

---

## Renderへのデプロイ

`render.yaml` に基づき、以下の設定が適用される:

- ランタイム: Python 3.12
- 起動コマンド: `gunicorn app:app --worker-class gevent --workers 1 --timeout 300`
- **workers は必ず 1 にすること**（`docs/02_decisions.md` ADR-001 参照）

Renderダッシュボードで以下の環境変数を設定:
- `GEMINI_API_KEY`
- `APP_PASSWORD`
- `SECRET_KEY`

---

## ファイル構成

```
dbm-report-agent/
├── app.py                  # Flaskサーバー（エンドポイント、認証、SSE）
├── classify_photos.py      # Gemini APIによる写真分類
├── place_photos.py         # Excelへの写真配置（openpyxl）
├── render.yaml             # Renderデプロイ設定
├── requirements.txt        # 依存パッケージ一覧
├── static/
│   ├── index.html          # メイン画面
│   └── login.html          # ログイン画面
└── docs/
    └── 02_decisions.md     # 設計判断ログ（ADR）
```

---

## 主な制約・注意事項

- **写真は最大30枚まで**（Gemini APIのペイロード制限による）
- **ダウンロードは1回限り**（ダウンロード後にサーバー上のファイルは即削除される）
- **同時処理は想定外**（workers=1のシングルプロセス構成、詳細は `docs/02_decisions.md`）
