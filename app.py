"""
報告書写真配置Webアプリ - Flask サーバー
SSEで進行状況をリアルタイム表示し、完了後にExcelをダウンロード
"""
 
import os
import sys
import json
import uuid
import tempfile
import traceback
import secrets
import shutil
from pathlib import Path
from functools import wraps
 
from flask import Flask, request, jsonify, send_file, Response, stream_with_context, session, redirect, url_for
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
 
load_dotenv()
 
app = Flask(__name__, static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", secrets.token_hex(32))
 
limiter = Limiter(get_remote_address, app=app, default_limits=[])
 
# 一時ファイル管理用
UPLOAD_DIR = tempfile.mkdtemp(prefix="dbm_agent_")
RESULT_DIR = tempfile.mkdtemp(prefix="dbm_result_")
 
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
 
# セッション内の元ファイル名を保持する辞書
original_names = {}
 
 
# ============================================================
# パスワード認証
# ============================================================
 
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not APP_PASSWORD:
            return f(*args, **kwargs)
        if not session.get("authenticated"):
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated
 
 
@app.route("/login", methods=["GET"])
def login_page():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    return send_file("static/login.html")
 
 
@app.route("/login", methods=["POST"])
@limiter.limit("5 per minute")
def login_submit():
    data = request.get_json() or {}
    password = data.get("password", "")
    if password == APP_PASSWORD:
        session["authenticated"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "パスワードが違います"}), 401
 
 
@app.errorhandler(429)
def ratelimit_handler(e):
    return jsonify({"ok": False, "error": "しばらく待ってから再試行してください"}), 429
 
 
@app.route("/")
@login_required
def index():
    return send_file("static/index.html")
 
 
@app.route("/upload", methods=["POST"])
@login_required
def upload():
    """写真とテンプレートをアップロードし、セッションIDを返す"""
    session_id = str(uuid.uuid4())
    session_dir = Path(UPLOAD_DIR) / session_id
    photo_dir = session_dir / "photos"
    photo_dir.mkdir(parents=True, exist_ok=True)
 
    # テンプレート保存
    template_file = request.files.get("template")
    if not template_file:
        return jsonify({"error": "テンプレートが選択されていません"}), 400
 
    # 元のファイル名を保持（日本語含む）
    original_template_name = template_file.filename or "template.xlsx"
    original_names[session_id] = original_template_name
 
    # サーバー保存用にはsecure_filenameを使用
    template_name = secure_filename(template_file.filename) or "template.xlsx"
    template_path = session_dir / template_name
    template_file.save(str(template_path))
 
    # 写真保存
    photos = request.files.getlist("photos")
    if not photos or photos[0].filename == "":
        return jsonify({"error": "写真が選択されていません"}), 400
 
    MAX_PHOTOS = 30
    if len(photos) > MAX_PHOTOS:
        return jsonify({"error": f"写真は{MAX_PHOTOS}枚までです（{len(photos)}枚選択されています）"}), 400
 
    photo_paths = []
    for photo in photos:
        fname = secure_filename(photo.filename) or f"photo_{len(photo_paths)}.jpg"
        fpath = photo_dir / fname
        photo.save(str(fpath))
        photo_paths.append(str(fpath))
 
    return jsonify({
        "session_id": session_id,
        "template": template_name,
        "photo_count": len(photo_paths),
    })
 
 
@app.route("/process/<session_id>")
@login_required
def process(session_id):
    """SSEで処理の進行状況を配信する"""
    session_dir = Path(UPLOAD_DIR) / session_id
 
    if not session_dir.exists():
        return jsonify({"error": "セッションが見つかりません"}), 404
 
    # テンプレートと写真のパスを取得
    template_path = None
    for f in session_dir.iterdir():
        if f.suffix == ".xlsx" and f.is_file():
            template_path = str(f)
            break
 
    if not template_path:
        return Response(
            sse_event("error_event", {"message": "テンプレートファイルが見つかりません"}),
            mimetype="text/event-stream",
        )
 
    photo_dir = session_dir / "photos"
    exts = {".jpg", ".jpeg", ".png"}
    photo_paths = sorted([
        str(f) for f in photo_dir.iterdir()
        if f.suffix.lower() in exts
    ])
 
    # ダウンロード用ファイル名（元のテンプレート名を使用）
    download_name = original_names.get(session_id, "output.xlsx")
 
    def generate():
        try:
            # Step 1: テンプレート解析
            yield sse_event("progress", {"step": "template", "message": "テンプレート解析中..."})
 
            from classify_photos import classify_and_assign
            from place_photos import place_photos
 
            # Step 2: AI分類
            yield sse_event("progress", {"step": "classify", "message": f"AI分類中... ({len(photo_paths)}枚の写真を分析)"})
 
            assigned, parsed_slots, slots_by_sheet = classify_and_assign(template_path, photo_paths, GEMINI_API_KEY)
 
            assigned_count = sum(1 for p in assigned if p is not None)
            yield sse_event("progress", {"step": "classified", "message": f"分類完了: {assigned_count}/{len(parsed_slots)}スロットに割り当て"})
 
            # 分類結果の詳細を送信
            details = []
            for i, (slot, photo) in enumerate(zip(parsed_slots, assigned)):
                photo_name = Path(photo).name if photo else "(空)"
                state_str = slot["state"] if slot["state"] else "なし"
                details.append(f'スロット{i}: {slot["work_type"]}[{state_str}] → {photo_name}')
            yield sse_event("details", {"assignments": details})
 
            # Step 3: 写真配置
            yield sse_event("progress", {"step": "place", "message": "Excelに写真を配置中..."})
 
            # 保存用ファイル名はsession_idベース（一意性確保）
            save_name = f"output_{session_id[:8]}.xlsx"
            output_path = str(Path(RESULT_DIR) / save_name)
            place_photos(template_path, output_path, assigned, precomputed_slots=slots_by_sheet)
 
            yield sse_event("progress", {"step": "done", "message": "完了! ダウンロードを開始します"})
            yield sse_event("complete", {
                "download_url": f"/download/{session_id}/{save_name}",
                "download_name": download_name,
            })
 
            # アップロードファイルを削除（結果ファイルはダウンロード後に削除）
            shutil.rmtree(session_dir, ignore_errors=True)
            original_names.pop(session_id, None)
        except Exception as e:
            traceback.print_exc()
            yield sse_event("error_event", {"message": str(e)})
            shutil.rmtree(session_dir, ignore_errors=True)
            original_names.pop(session_id, None)
 
    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )
 
 
@app.route("/download/<session_id>/<filename>")
@login_required
def download(session_id, filename):
    """生成されたExcelファイルをダウンロード"""
    filepath = (Path(RESULT_DIR) / filename).resolve()
    # パストラバーサル防止: RESULT_DIR配下であることを検証
    if not str(filepath).startswith(str(Path(RESULT_DIR).resolve())):
        return jsonify({"error": "不正なパスです"}), 403
    if not filepath.exists():
        return jsonify({"error": "ファイルが見つかりません"}), 404
 
    # 元のテンプレートファイル名でダウンロード（取得後にdictから削除）
    dl_name = original_names.pop(session_id, filename)

    # ファイルをメモリに読み込んでから削除（after_this_requestはFlask 3.1で廃止）
    import io
    data = filepath.read_bytes()
    filepath.unlink(missing_ok=True)

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=dl_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
 
def sse_event(event_type, data):
    """SSEイベント文字列を生成"""
    return f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"
 
 
if __name__ == "__main__":
    if not GEMINI_API_KEY:
        print("エラー: .envにGEMINI_API_KEYを設定してください")
        sys.exit(1)
    print("サーバー起動: http://localhost:5000")
    app.run(debug=True, port=5000)