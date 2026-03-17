"""
Step 2: 写真配置ロジック
ダミー画像をテンプレートのスロットに配置して確認する
 
旧ツールの問題を解消するためのポイント:
- 結合セルの実際のサイズを行高さ・列幅の合計から計算（セルオブジェクトに依存しない）
- OneCellAnchorで位置+画像EMUサイズを指定して配置（ストレッチ防止）
- アスペクト比を維持してセル内に収まるよう縮小、中央配置（トリミングなし）
"""
 
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from PIL import Image as PILImage
import statistics
import struct
import zlib
import os
import shutil
from pathlib import Path
 
 
# ============================================================
# EMU変換ユーティリティ
# ============================================================
 
# 標準フォント(11pt)ごとのMaximum Digit Width
_MDW_TABLE = {
    'ＭＳ Ｐゴシック': 8, 'MS PGothic': 8,
    'ＭＳ ゴシック': 7, 'MS Gothic': 7,
    '游ゴシック': 8, 'Yu Gothic': 8, 'Yu Gothic UI': 7,
    'メイリオ': 9, 'Meiryo': 9, 'Meiryo UI': 8,
    'Arial': 7, 'Calibri': 7, 'Aptos': 7,
    'Times New Roman': 7, 'Verdana': 8, 'Tahoma': 7,
}
 
 
def detect_mdw(wb):
    """ワークブックの標準スタイルフォントからMDWを取得"""
    for style in wb._named_styles:
        if style.name in ('標準', 'Normal'):
            font_name = style.font.name
            if font_name and font_name in _MDW_TABLE:
                return _MDW_TABLE[font_name]
    return 7  # 不明な場合はデフォルト
 
 
def col_width_to_emu(width_chars, mdw):
    """Excel公式の列幅→ピクセル→EMU変換"""
    if width_chars == 0:
        return 0
    px = int(((256 * width_chars + int(128 / mdw)) / 256) * mdw)
    return px * 9525
 
 
def row_height_to_emu(height_pt):
    return int(height_pt * 12700)
 
 
# ============================================================
# テンプレート解析
# ============================================================
 
def get_merged_cell_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return str(cell.value).strip()
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if top_left.value is not None:
                return str(top_left.value).strip()
    return ""
 
 
def get_row_height(ws, row):
    rd = ws.row_dimensions.get(row)
    if rd and rd.height is not None:
        return rd.height
    return 15.0
 
 
def get_col_width(ws, col):
    col_letter = openpyxl.utils.get_column_letter(col)
    cd = ws.column_dimensions.get(col_letter)
    if cd and cd.width is not None:
        return cd.width
    return 8.43
 
 
def detect_photo_slots(ws):
    max_row = ws.max_row
    max_col = ws.max_column
 
    row_heights = {r: get_row_height(ws, r) for r in range(1, max_row + 1)}
    col_widths = {c: get_col_width(ws, c) for c in range(1, max_col + 1)}
 
    heights = list(row_heights.values())
    median_h = statistics.median(heights)
    threshold_h = median_h * 3.0
    photo_rows = sorted([r for r, h in row_heights.items() if h >= threshold_h])
 
    if not photo_rows:
        return [], row_heights, col_widths
 
    widths = list(col_widths.values())
    median_w = statistics.median(widths)
    threshold_w = median_w * 1.2
    content_cols = sorted([c for c, w in col_widths.items() if w >= threshold_w])
 
    if not content_cols:
        return [], row_heights, col_widths
 
    slots = []
    for idx, r in enumerate(photo_rows):
        row_min = photo_rows[idx - 1] + 1 if idx > 0 else 1
        for c in content_cols:
            # カテゴリ探索（写真行から最大3行上まで）
            category = ""
            cat_row = r
            for look_up in range(1, 4):
                search_row = r - look_up
                if search_row < row_min:
                    break
                candidate = get_merged_cell_value(ws, search_row, c)
                if candidate:
                    category = candidate
                    cat_row = search_row
                    break
 
            # カテゴリが見つからなければスロットから除外
            if not category:
                continue
 
            # セクション探索（カテゴリ行から最大3行上まで）
            section = ""
            for look_up in range(1, 4):
                search_row = cat_row - look_up
                if search_row < row_min:
                    break
                candidate = get_merged_cell_value(ws, search_row, c)
                if not candidate:
                    candidate = get_merged_cell_value(ws, search_row, content_cols[0])
                if candidate and candidate != category:
                    section = candidate
                    break
 
            slots.append({
                "row": r,
                "col": c,
                "category": category,
                "section": section,
                "row_heights": row_heights,
                "col_widths": col_widths,
            })
 
    return slots, row_heights, col_widths
 
 
# ============================================================
# スロット結合範囲取得
# ============================================================
 
def get_slot_merged_range(ws, row, col):
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
            return (merged_range.min_row, merged_range.min_col,
                    merged_range.max_row, merged_range.max_col)
    return (row, col, row, col)
 
 
# ============================================================
# ダミーPNG生成
# ============================================================
 
def make_dummy_png(filepath, width=400, height=300, color=(200, 100, 50), label=""):
    def png_chunk(name, data):
        c = struct.pack('>I', len(data)) + name + data
        crc = zlib.crc32(name + data) & 0xffffffff
        return c + struct.pack('>I', crc)
 
    header = b'\x89PNG\r\n\x1a\n'
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)
    ihdr = png_chunk(b'IHDR', ihdr_data)
 
    raw_rows = []
    for y in range(height):
        row = b'\x00'
        row += bytes(color) * width
        raw_rows.append(row)
 
    raw_data = b''.join(raw_rows)
    compressed = zlib.compress(raw_data)
    idat = png_chunk(b'IDAT', compressed)
    iend = png_chunk(b'IEND', b'')
 
    with open(filepath, 'wb') as f:
        f.write(header + ihdr + idat + iend)
 
    print(f"  ダミー画像生成: {filepath}")
 
 
# ============================================================
# 写真配置メイン
# ============================================================
 
def place_photos(template_path, output_path, photo_paths):
    import tempfile
 
    print(f"\n{'='*60}")
    print(f"写真配置開始: {Path(template_path).name}")
    print(f"{'='*60}")
 
    wb = openpyxl.load_workbook(template_path)
    mdw = detect_mdw(wb)
    print(f"  MDW={mdw}（標準フォントから自動検出）")
    tmp_dir = tempfile.mkdtemp()
    total_placed = 0
 
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            slots, row_heights, col_widths = detect_photo_slots(ws)
 
            if not slots:
                print(f"  スキップ: '{sheet_name}'")
                continue

            # 既存画像を全て削除（重複配置を防止）
            existing_count = len(ws._images)
            if existing_count > 0:
                ws._images.clear()
                print(f"  シート '{sheet_name}': 既存画像{existing_count}枚を削除")

            print(f"\n  シート '{sheet_name}': {len(slots)}スロットに配置")
 
            photo_iter = iter(photo_paths)
 
            for i, slot in enumerate(slots):
                try:
                    img_path = next(photo_iter)
                except StopIteration:
                    print(f"  [!] 写真が足りません（スロット{i+1}以降はスキップ）")
                    break
 
                # 未割り当てスロット（None）はスキップ
                if img_path is None:
                    print(f"    スロット{i+1}: (未割り当て)")
                    continue
 
                row_heights = slot["row_heights"]
                col_widths = slot["col_widths"]
 
                min_row, min_col, max_row, max_col = get_slot_merged_range(ws, slot["row"], slot["col"])
 
                slot_w_emu = sum(col_width_to_emu(col_widths.get(c, 8.43), mdw) for c in range(min_col, max_col + 1))
                slot_h_emu = sum(row_height_to_emu(row_heights.get(r, 15.0)) for r in range(min_row, max_row + 1))
 
                # --- アスペクト比維持フィット（トリミングなし・中央配置） ---
                with PILImage.open(img_path) as img:
                    img_w, img_h = img.size
 
                img_ratio = img_w / img_h
                slot_ratio = slot_w_emu / slot_h_emu
 
                if img_ratio > slot_ratio:
                    # 画像のほうが横長 → 幅に合わせる
                    fit_w_emu = slot_w_emu
                    fit_h_emu = int(slot_w_emu / img_ratio)
                else:
                    # 画像のほうが縦長 → 高さに合わせる
                    fit_h_emu = slot_h_emu
                    fit_w_emu = int(slot_h_emu * img_ratio)
 
                # EMU→px変換は切り捨てで誤差が出るため、アスペクト比をpx段階で再計算
                fit_w_px = max(1, fit_w_emu // 9525)
                fit_h_px = max(1, round(fit_w_px * img_h / img_w) if img_ratio > slot_ratio else fit_h_emu // 9525)
                # EMUをpxから逆算して一致させる
                fit_w_emu = fit_w_px * 9525
                fit_h_emu = fit_h_px * 9525
 
                # オフセットはEMU確定後に計算（中央配置）
                x_off = (slot_w_emu - fit_w_emu) // 2
                y_off = (slot_h_emu - fit_h_emu) // 2
 
                with PILImage.open(img_path) as img:
                    resized = img.convert("RGB").resize((fit_w_px, fit_h_px), PILImage.LANCZOS)
 
                tmp_path = str(Path(tmp_dir) / f"_tmp_{Path(img_path).stem}.jpg")
                resized.save(tmp_path, "JPEG", quality=95)
 
                xl_img = XLImage(tmp_path)
 
                from_marker = AnchorMarker(
                    col=min_col - 1,
                    colOff=x_off,
                    row=min_row - 1,
                    rowOff=y_off,
                )
                ext = XDRPositiveSize2D(fit_w_emu, fit_h_emu)
                xl_img.anchor = OneCellAnchor(_from=from_marker, ext=ext)
                ws.add_image(xl_img)
 
                print(f"    スロット{i+1}: section='{slot['section']}' category='{slot['category']}'")
                print(f"      セル(row={min_row},col={min_col}) サイズ=({fit_w_px}x{fit_h_px}px) offset=({x_off},{y_off})")
                print(f"      画像: {Path(img_path).name}")
                total_placed += 1
 
        wb.save(output_path)
        print(f"\n{'='*60}")
        print(f"完了: {total_placed}枚配置 → {output_path}")
        print(f"{'='*60}")
 
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)