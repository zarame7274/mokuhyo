"""売上目標トラッカーの Excel を生成する。

使い方: python tools/make_sheet.py 出力先.xlsx

このシートは Web アプリのデータ保存先も兼ねる。
apps-script/Code.gs はここで決めたセル位置を前提にしているので、
レイアウトを変えるときは両方を合わせて直すこと。
"""
import datetime
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

out = sys.argv[1]
wb = Workbook()
ws = wb.active
ws.title = "売上目標"

bold = Font(bold=True)
note = Font(color="666666", size=9)
inp = PatternFill("solid", fgColor="FFF4C2")  # 入力欄（黄色）
head = PatternFill("solid", fgColor="DDE7F5")
thin = Side(style="thin", color="BBBBBB")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center")
yen = "#,##0"
qty = '#,##0.0"本"'

ws["A1"] = "売上目標トラッカー"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "黄色のセルだけ入力してください。目標は自動で再計算されます。"
ws["A2"].font = Font(color="666666")

# ---- 設定（B3:B6）----
settings = [
    ("単価", 3000, yen, True),
    ("目標本数", 1000, '#,##0"本"', True),
    ("月間目標", "=B3*B4", yen, False),
    ("対象月（1日）", datetime.date(2026, 11, 1), "yyyy/mm", True),
]
for i, (label, value, fmt, editable) in enumerate(settings):
    r = 3 + i
    ws.cell(r, 1, label).font = bold
    c = ws.cell(r, 2, value)
    c.number_format = fmt
    c.border = box
    if editable:
        c.fill = inp
ws["B5"].font = bold

# ---- サマリー（D3:E8）----
FIRST, LAST = 12, 42
A, C, D = f"$A${FIRST}:$A${LAST}", f"$C${FIRST}:$C${LAST}", f"$D${FIRST}:$D${LAST}"
summary = [
    ("累計実績", f"=SUM({D})", yen),
    ("達成率", "=IFERROR(E3/B5,0)", "0.0%"),
    ("残り目標", "=MAX(0,B5-E3)", yen),
    ("残り営業日", f'=COUNTIFS({A},">0",{D},"",{C},">0")', '0"日"'),
    ("残り1日あたり目標", f'=IFERROR(E5/SUMIFS({C},{D},"",{A},">0"),0)', yen),
    ("残り1日あたり本数", "=IFERROR(E7/$B$3,0)", qty),
]
for i, (label, f, fmt) in enumerate(summary):
    r = 3 + i
    ws.cell(r, 4, label).font = bold
    c = ws.cell(r, 5, f)
    c.number_format = fmt
    c.border = box
for c in ("E7", "E8"):
    ws[c].font = Font(bold=True, size=13, color="C0392B")
ws["D9"] = "※重み1の日の目標"
ws["D9"].font = note

# ---- 曜日別の重み（G3:M4）----
ws["G2"] = "曜日別の重み（売れる曜日は大きく）"
ws["G2"].font = bold
for i, w in enumerate("月火水木金土日"):
    h = ws.cell(3, 7 + i, w)
    h.fill, h.alignment, h.border = head, center, box
    v = ws.cell(4, 7 + i, 1)
    v.fill, v.alignment, v.border = inp, center, box
ws["G5"] = "※特定の日を休みにするときは、表の「重み」に 0 を入力"
ws["G5"].font = note

# ---- 日別テーブル（11行目が見出し、12〜42行目）----
headers = ["日付", "曜日", "重み", "売上実績", "目標", "差額（実績−目標）", "累計実績", "目標本数", "実績本数"]
for i, h in enumerate(headers):
    c = ws.cell(FIRST - 1, 1 + i, h)
    c.font, c.fill, c.border, c.alignment = bold, head, box, center

for r in range(FIRST, LAST + 1):
    a = f"A{r}"
    ws[a] = "=B6" if r == FIRST else f'=IF(A{r-1}="","",IF(MONTH(A{r-1}+1)=MONTH($B$6),A{r-1}+1,""))'
    ws[a].number_format = "m/d"
    ws[f"B{r}"] = f'=IF({a}="","",CHOOSE(WEEKDAY({a}),"日","月","火","水","木","金","土"))'
    ws[f"C{r}"] = f'=IF({a}="","",INDEX($G$4:$M$4,WEEKDAY({a},2)))'
    # 実績入力済みの日 → その日の朝時点の目標 / 未入力の日 → 現在の残り目標を重みで配分
    ws[f"E{r}"] = (
        f'=IF({a}="","",IF(N(C{r})=0,0,IF(D{r}<>"",'
        f'MAX(0,$B$5-SUMIFS({D},{A},"<"&{a}))*C{r}/SUMIFS({C},{A},">="&{a}),'
        f'IFERROR($E$5*C{r}/SUMIFS({C},{D},"",{A},">0"),0))))'
    )
    ws[f"F{r}"] = f'=IF(OR({a}="",D{r}=""),"",D{r}-E{r})'
    ws[f"G{r}"] = f'=IF(OR({a}="",D{r}=""),"",SUMIFS({D},{A},"<="&{a}))'
    ws[f"H{r}"] = f'=IF({a}="","",IFERROR(E{r}/$B$3,0))'
    ws[f"I{r}"] = f'=IF(OR({a}="",D{r}=""),"",IFERROR(D{r}/$B$3,0))'
    ws[f"D{r}"].fill = inp
    for col in "ABCDEFGHI":
        ws[f"{col}{r}"].border = box
    for col in "DEFG":
        ws[f"{col}{r}"].number_format = yen
    for col in "HI":
        ws[f"{col}{r}"].number_format = qty
    ws[f"B{r}"].alignment = ws[f"C{r}"].alignment = center

rng = f"F{FIRST}:F{LAST}"
ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C0392B")))
ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(color="1E8449")))

for col, w in zip("ABCDEFGHIJKLM", [14, 8, 8, 14, 18, 18, 14, 10, 10, 6, 6, 6, 6]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = f"A{FIRST}"
wb.save(out)
print("saved")
