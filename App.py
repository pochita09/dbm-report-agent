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
from pathlib import Path
 
from flask import Flask, request, jsonify, send_file, Response, stream_with_context
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
 
load_dotenv()
 
app = Flask(__name__, static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB
 
# 一時ファイル管理用
UPLOAD_DIR = tempfile.mkdtemp(prefix="dbm_agent_")
RESULT_DIR = tempfile.mkdtemp(prefix="dbm_result_")
 
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
 
 
@app.route("/")
def index():
    return send_file("static/index.html")
 
 
@app.route("/upload", methods=["POST"])
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
 
    template_name = secure_filename(template_file.filename) or "template.xlsx"
    template_path = session_dir / template_name
    template_file.save(str(template_path))
 
    # 写真保存
    photos = request.files.getlist("photos")
    if not photos or photos[0].filename == "":
        return jsonify({"error": "写真が選択されていません"}), 400
 
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
 
    def generate():
        try:
            # Step 1: テンプレート解析
            yield sse_event("progress", {"step": "template", "message": "テンプレート解析中..."})
 
            from classify_photos import classify_and_assign
            from place_photos import place_photos
 
            # Step 2: AI分類
            yield sse_event("progress", {"step": "classify", "message": f"AI分類中... ({len(photo_paths)}枚の写真を分析)"})
 
            assigned, parsed_slots = classify_and_assign(template_path, photo_paths, GEMINI_API_KEY)
 
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
 
            output_name = f"output_{session_id[:8]}.xlsx"
            output_path = str(Path(RESULT_DIR) / output_name)
            place_photos(template_path, output_path, assigned)
 
            yield sse_event("progress", {"step": "done", "message": "完了! ダウンロードを開始します"})
            yield sse_event("complete", {"download_url": f"/download/{session_id}/{output_name}"})
 
        except Exception as e:
            traceback.print_exc()
            yield sse_event("error_event", {"message": str(e)})
 
    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )
 
 
@app.route("/download/<session_id>/<filename>")
def download(session_id, filename):
    """生成されたExcelファイルをダウンロード"""
    filepath = Path(RESULT_DIR) / filename
    if not filepath.exists():
        return jsonify({"error": "ファイルが見つかりません"}), 404
    return send_file(
        str(filepath),
        as_attachment=True,
        download_name=filename,
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