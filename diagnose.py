import openpyxl

wb = openpyxl.load_workbook(r"c:\Users\nyaaa\OneDrive\デスクトップ\報告書テンプレ\20260226_○○現場_ｶﾞﾗｽ定期特別清掃作業報告書.xlsx")
ws = wb["A棟"]

print("=== 結合セル一覧 ===")
for mr in ws.merged_cells.ranges:
    print(f"  {mr}")

print("\n=== 行高さ ===")
for row in range(1, 25):
    rd = ws.row_dimensions.get(row)
    h = rd.height if rd and rd.height else 15.0
    print(f"  行{row}: {h}pt")