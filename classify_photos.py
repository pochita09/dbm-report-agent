"""
Step 3: AI写真分類モジュール
Gemini APIを使って作業写真を分類し、テンプレートのスロットに割り当てる
 
処理フロー:
1. detect_photo_slots でスロット情報を取得
2. section/category から「作業内容」と「状態」を分離
3. Gemini APIに全写真を一括送信し分類
4. スロット順の画像パスリスト（割り当てなし=None）を返す
"""
 
import sys
import base64
import json
import os
import re
from pathlib import Path
 
import requests
from PIL import Image as PILImage
 
import openpyxl
from place_photos import detect_photo_slots
 
 
# ============================================================
# 定数
# ============================================================
 
# 「状態」を表すキーワード（これ以外は「作業内容」として扱う）
STATE_KEYWORDS = ["作業前", "作業中", "作業後"]
 
# Gemini APIエンドポイント
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
 
 
# ============================================================
# 補足層: ガラス清掃の分類ルール
# ============================================================
 
GLASS_CLEANING_RULES = """
## 作業内容の判定ルール（ガラス清掃）
以下の優先順位で判定すること。上位に該当する場合、下位には分類しない。
 
### 優先度1: 館銘板（館名板）
建物名称・ロゴ・住所が刻印またはプリントされたプレートや看板。
ステンレス、石材、アクリル、ガラス製など。建物入口付近に設置。
「館銘板」「館名板」は同じカテゴリとして扱うこと。
 
### 優先度2: 立入禁止区域の設置
作業現場の安全確保のために設置された規制機材が主役の写真。
カラーコーン（ロードコーン）、コーンバー、虎ロープ（黄黒ロープ）、プラスチックチェーン。
作業員ではなく、地面に並べられた機材がメインの構図。
 
### 優先度3: ロープ作業 / ゴンドラ作業
建物外壁を屋上から吊り下がって清掃する手法。
ロープ（ブランコ）: 作業員が椅子状の板に座り、ロープで降下している。
ゴンドラ: 金属製の箱状昇降機に乗り、ワイヤーで吊られている。
 
### 優先度4: 高所作業車 / リフター
動力を持つ昇降機械を使用した作業。
高所作業車: 走行車の荷台から伸縮アームが伸びているもの。
リフター: シザーリフトなど、垂直に昇降する自走式の台。
 
### 優先度5: ローリングタワー
鋼管パイプを組み上げた移動式足場。
足元にキャスター付き、人力で移動可能。吹き抜けやロビーなど屋内で多用。
 
### 優先度6: エントランス他（上記のいずれにも該当しない場合）
スクイジー清掃: 手持ちの水切り道具で至近距離のガラスを拭いている。
ポール作業: 伸縮式の長い棒の先に道具をつけて地面から清掃。
その他: 窓枠（サッシ）清掃や上記に分類できない作業。
"""
 
 
# ============================================================
# スロット情報の解析
# ============================================================
 
def _strip_numbering(text):
    """セクション名から番号部分を除去する
    例: '（１）ロープ作業' → 'ロープ作業'
         '(2) エントランス他' → 'エントランス他'
    """
    return re.sub(r'^[（\(]\s*[０-９0-9]+\s*[）\)]\s*', '', text).strip()
 
 
def parse_slot_info(slots):
    """detect_photo_slotsの結果から「作業内容」と「状態」を分離する
 
    Returns:
        list[dict]: 各スロットに以下のキーを追加した情報
            - work_type: 作業内容（ロープ作業、エントランス他 etc.）
            - state: 状態（作業前/作業中/作業後/None）
    """
    parsed = []
    for slot in slots:
        section = slot.get("section", "")
        category = slot.get("category", "")
 
        work_type = ""
        state = None
 
        # categoryを先に見る。STATE_KEYWORDSでなければwork_type、
        # STATE_KEYWORDSならstateにし、sectionからwork_typeを取る
        for text in [category, section]:
            stripped = text.strip()
            if not stripped:
                continue
            if stripped in STATE_KEYWORDS:
                state = stripped
            else:
                if not work_type:
                    work_type = _strip_numbering(stripped)
 
        parsed.append({
            "slot_index": len(parsed),
            "row": slot["row"],
            "col": slot["col"],
            "work_type": work_type,
            "state": state,
            "section": section,
            "category": category,
        })
 
    return parsed
 
 
# ============================================================
# プロンプト構築
# ============================================================
 
def build_prompt(parsed_slots, photo_filenames):
    """基本層 + 補足層のプロンプトを構築する"""
 
    # --- スロット情報テキスト ---
    slot_lines = []
    for s in parsed_slots:
        state_str = f'"{s["state"]}"' if s["state"] else "なし"
        slot_lines.append(
            f'スロット{s["slot_index"]}: 作業内容="{s["work_type"]}" 状態={state_str}'
        )
    slot_text = "\n".join(slot_lines)
 
    # --- 写真一覧テキスト ---
    photo_text = ", ".join(photo_filenames)
 
    # --- 基本層 ---
    base_prompt = f"""あなたは清掃作業の写真を分類する専門家です。
以下の写真を分析し、テンプレートの各スロットに割り当ててください。
 
## スロット情報
{slot_text}
 
## 送信された写真
{photo_text}
 
## 分類ルール（優先度順に適用）
 
### ステップ1: 完全一致を優先
各写真の「作業内容」と「状態」を判定し、両方が一致するスロットに割り当てる。
 
### ステップ2: 空きスロットを埋める
ステップ1で空きスロットが残り、かつ未使用の写真がある場合:
- 同じ作業内容で状態が異なる写真を割り当てる（例: 作業中のスロットに作業前の写真）
- 同じ作業内容の写真もない場合は、近いカテゴリの写真を割り当てる
 
### ステップ3: nullにする条件
以下の場合のみnullとする。無理に当てはめてはならない:
- 未使用の写真が1枚も残っていない
- 残っている写真がスロットの作業内容と全く関係ない（例: ロープ作業のスロットにエントランスの写真）
 
### その他のルール
- 同じ写真を複数スロットに使い回さない
- 同じ作業内容に複数枚割り当てる場合、構図がなるべく異なるものを選ぶ
 
## 状態の判定基準
- 状態="作業中" のスロット → 人（作業員）が写っている写真を割り当てる
- 状態="作業前" または "作業後" のスロット → 人が写っていない作業箇所の写真を割り当てる
- 状態=なし → 状態は考慮不要、作業内容のみで判定する
 
## 作業前・作業中・作業後の横並びルール
同じ作業内容で「作業前」「作業中」「作業後」が横並びになっている場合:
- この3スロットには同じ場所（同じ窓、同じガラス面、同じアングル）の写真をセットで割り当てること
- 「作業前」「作業後」の区別は不要。人なし写真をどちらに入れてもよい
- セットが組めない場合は、可能な範囲で近い写真を割り当てる
 
## エントランス優先配置ルール
作業内容に「エントランス」を含む場合、そのグループの最初のスロットには
建物のエントランス（自動ドア・メイン入口）の写真を優先的に配置すること。
該当写真がない場合は他の写真で埋める。
 
## 回答形式
以下のJSON形式のみで回答すること。説明文やマークダウンの装飾は一切不要。
{{
  "assignments": [
    {{"slot_index": 0, "file": "ファイル名 or null"}},
    {{"slot_index": 1, "file": "ファイル名 or null"}},
    ...
  ]
}}
"""
 
    # --- 基本層 + 補足層を結合 ---
    return base_prompt + "\n" + GLASS_CLEANING_RULES
 
 
# ============================================================
# 画像エンコード
# ============================================================
 
def encode_photo(photo_path, max_long_side=1024):
    """画像をリサイズしてbase64エンコードする
 
    Args:
        photo_path: 画像ファイルパス
        max_long_side: 長辺の最大ピクセル数（API送信量削減）
 
    Returns:
        tuple: (base64文字列, MIMEタイプ)
    """
    with PILImage.open(photo_path) as img:
        img = img.convert("RGB")
 
        # 長辺がmax_long_sideを超える場合はリサイズ
        w, h = img.size
        if max(w, h) > max_long_side:
            ratio = max_long_side / max(w, h)
            img = img.resize((int(w * ratio), int(h * ratio)), PILImage.LANCZOS)
 
        # JPEG → base64
        import io
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
 
    return b64, "image/jpeg"
 
 
# ============================================================
# Gemini API呼び出し
# ============================================================
 
def call_gemini_api(prompt, photo_paths, api_key):
    """Gemini APIに写真を一括送信して分類結果を取得する
 
    Args:
        prompt: 構築済みプロンプト
        photo_paths: 画像ファイルパスのリスト
        api_key: Gemini APIキー
 
    Returns:
        dict: パース済みのJSON（assignments含む）
    """
    # リクエストボディ構築
    parts = []
 
    # テキストプロンプト
    parts.append({"text": prompt})
 
    # 各写真をインライン画像として追加
    for photo_path in photo_paths:
        filename = Path(photo_path).name
        b64, mime = encode_photo(photo_path)
 
        # ファイル名ラベル
        parts.append({"text": f"[写真: {filename}]"})
        # 画像データ
        parts.append({
            "inline_data": {
                "mime_type": mime,
                "data": b64,
            }
        })
 
    body = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "temperature": 0.1,  # 分類タスクなので低温
            "maxOutputTokens": 8192,
        },
    }
 
    # API呼び出し（APIキーはヘッダーで送信し、ログへの漏洩を防止）
    print(f"  Gemini API呼び出し中... ({len(photo_paths)}枚送信)")
 
    response = requests.post(
        GEMINI_API_URL,
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": api_key,
        },
        json=body,
        timeout=120,
    )
    if not response.ok:
        raise RuntimeError(f"Gemini API error: {response.status_code} {response.reason}")
    result = response.json()
 
    # レスポンスからテキスト部分を抽出
    try:
        text = result["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Gemini APIレスポンスの解析に失敗: {e}\n{json.dumps(result, ensure_ascii=False, indent=2)}")
 
    print(f"  Gemini APIレスポンス受信")
 
    # JSONをパース（マークダウンコードブロックの除去）
    text = text.strip()
    text = re.sub(r'^```json\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
 
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Gemini APIレスポンスのJSON解析に失敗: {e}\nレスポンス:\n{text}")
 
    return parsed
 
 
# ============================================================
# 割り当て結果の変換
# ============================================================
 
def assign_photos(api_response, parsed_slots, photo_paths):
    """APIレスポンスをスロット順の画像パスリストに変換する
 
    Args:
        api_response: Gemini APIのパース済みレスポンス
        parsed_slots: parse_slot_infoの結果
        photo_paths: 画像ファイルパスのリスト
 
    Returns:
        list: スロット順の画像パスリスト（未割り当て=None）
    """
    # ファイル名→パスのマッピング
    name_to_path = {Path(p).name: p for p in photo_paths}
 
    assignments = api_response.get("assignments", [])
    result = [None] * len(parsed_slots)
    used_files = set()
 
    for entry in assignments:
        idx = entry.get("slot_index")
        filename = entry.get("file")
 
        if idx is None or idx >= len(parsed_slots):
            continue
        if filename is None or filename == "null":
            continue
        if filename in used_files:
            print(f"  [!] 警告: {filename} が重複割り当て（スキップ）")
            continue
        if filename not in name_to_path:
            print(f"  [!] 警告: {filename} が写真一覧に見つかりません")
            continue
 
        result[idx] = name_to_path[filename]
        used_files.add(filename)
 
    return result
 
 
# ============================================================
# メイン関数
# ============================================================
 
def classify_and_assign(template_path, photo_paths, api_key):
    """テンプレートと写真からAI分類を実行し、スロット順の画像パスリストを返す
 
    Args:
        template_path: Excelテンプレートのパス
        photo_paths: 画像ファイルパスのリスト
        api_key: Gemini APIキー
 
    Returns:
        tuple: (assigned_photos, parsed_slots, slots_by_sheet)
            assigned_photos: スロット順の画像パスリスト（未割り当て=None）
            parsed_slots: 解析済みスロット情報
            slots_by_sheet: シートごとのスロット情報 {sheet_name: [slot, ...]}
    """
    print(f"\n{'='*60}")
    print(f"AI写真分類開始")
    print(f"{'='*60}")
 
    # 1. テンプレート解析
    wb = openpyxl.load_workbook(template_path)
    all_slots = []
    slots_by_sheet = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        slots, _, _ = detect_photo_slots(ws)
        if slots:
            print(f"  シート '{sheet_name}': {len(slots)}スロット検出")
            slots_by_sheet[sheet_name] = slots
            all_slots.extend(slots)
    wb.close()
 
    if not all_slots:
        print("  [!] 写真スロットが見つかりません")
        return [], [], {}
 
    # 2. スロット情報の解析（作業内容＋状態の分離）
    parsed_slots = parse_slot_info(all_slots)
    print(f"\n  スロット情報:")
    for s in parsed_slots:
        state_str = s["state"] if s["state"] else "なし"
        print(f'    スロット{s["slot_index"]}: 作業内容="{s["work_type"]}" 状態={state_str}')
 
    # 3. プロンプト構築
    photo_filenames = [Path(p).name for p in photo_paths]
    prompt = build_prompt(parsed_slots, photo_filenames)
 
    print(f"\n  写真: {len(photo_paths)}枚")
    for name in photo_filenames:
        print(f"    {name}")
 
    # 4. Gemini API呼び出し
    api_response = call_gemini_api(prompt, photo_paths, api_key)
 
    # 5. 割り当て結果の変換
    assigned = assign_photos(api_response, parsed_slots, photo_paths)
 
    # 結果表示
    print(f"\n  割り当て結果:")
    for i, (slot, photo) in enumerate(zip(parsed_slots, assigned)):
        photo_name = Path(photo).name if photo else "(空)"
        state_str = slot["state"] if slot["state"] else "なし"
        print(f'    スロット{i}: {slot["work_type"]}[{state_str}] → {photo_name}')
 
    assigned_count = sum(1 for p in assigned if p is not None)
    print(f"\n  完了: {assigned_count}/{len(parsed_slots)}スロットに割り当て")
 
    return assigned, parsed_slots, slots_by_sheet
 
 
# ============================================================
# テスト実行
# ============================================================
 
if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()
 
    API_KEY = os.environ.get("GEMINI_API_KEY")
    if not API_KEY:
        print("エラー: .envにGEMINI_API_KEYを設定してください")
        sys.exit(1)
 
    # テスト用パス（実行環境に合わせて変更）
    TEMPLATE = r"c:\Users\nyaaa\OneDrive\デスクトップ\報告書テンプレ\20260226_○○現場_ｶﾞﾗｽ定期特別清掃作業報告書.xlsx"
    PHOTO_DIR = r"c:\Users\nyaaa\OneDrive\デスクトップ\DBM"
 
    exts = {'.jpg', '.jpeg', '.png'}
    photo_paths = sorted([
        os.path.join(PHOTO_DIR, f)
        for f in os.listdir(PHOTO_DIR)
        if os.path.splitext(f)[1].lower() in exts
    ])
 
    print(f"\n{len(photo_paths)}枚の写真を使用します")
    for p in photo_paths:
        print(f"  {Path(p).name}")
 
    assigned, parsed_slots = classify_and_assign(TEMPLATE, photo_paths, API_KEY)
 
    # 分類結果を使ってExcelに配置
    if any(p is not None for p in assigned):
        OUTPUT = r"c:\Users\nyaaa\OneDrive\デスクトップ\test_output_classified.xlsx"
        from place_photos import place_photos
        place_photos(TEMPLATE, OUTPUT, assigned)