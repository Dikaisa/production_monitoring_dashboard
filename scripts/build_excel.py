"""Bangun workbook Production Monitoring dari data simulasi (semua angka rekap pakai formula)."""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
df = pd.read_csv(ROOT / "data" / "produksi_harian.csv", parse_dates=["tanggal"])
N = len(df)
LAST = N + 1  # baris terakhir data (header di baris 1)

FONT = "Arial"
NAVY = "1F3864"
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_BOLD = Font(name=FONT, size=10, bold=True)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="DDEBF7")
FILL_CALC = PatternFill("solid", fgColor="FFF2CC")
FILL_TOTAL = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

D = "'Data Produksi'"
BULAN = [(5, "Mei 2026"), (6, "Jun 2026"), (7, "Jul 2026"),
         (8, "Agu 2026"), (9, "Sep 2026"), (10, "Okt 2026")]
PARTS = ["Part A", "Part B", "Part C", "Part D", "Part E"]


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    ws.row_dimensions[row].height = 32


def body(c, fmt=None, font=F_BODY, fill=None):
    c.font, c.border = font, BORDER
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_SUB
    ws.sheet_view.showGridLines = False


wb = Workbook()

# ---------------------------------------------------------------- Tentang
ws = wb.active
ws.title = "Tentang"
title(ws, "Production Monitoring Dashboard",
      "Template monitoring produksi harian lini perakitan, dengan data SIMULASI (bukan data asli perusahaan)")
info = [
    ("Latar belakang", "Rekap produksi di lini perakitan sebelumnya dikerjakan manual di kertas/Excel per hari. "
                       "Workbook ini menyatukan input harian, rekap otomatis, dan dashboard dalam satu file."),
    ("Periode data", "1 Mei 2026 s.d. 31 Oktober 2026 (hari kerja + Sabtu masuk)"),
    ("Jumlah baris", f"{N} baris (1 baris = 1 part per 1 hari kerja)"),
    ("Anonimisasi", "Nama perusahaan, model, line, dan part sudah diganti menjadi generik (Part A-E, Line 2/3)."),
    ("Data", "Data dibuat dengan script Python (scripts/generate_data.py), bukan data produksi asli."),
    ("", ""),
    ("ISI SHEET", ""),
    ("Dashboard", "KPI utama + grafik tren pencapaian, NG rate, dan kekurangan."),
    ("Laporan Harian", "Pilih tanggal di sel C4, laporan per part terisi otomatis dari Data Produksi."),
    ("Rekap Bulanan", "Ringkasan per bulan + heatmap pencapaian per part per bulan."),
    ("Rekap Part", "Ringkasan per part + perbandingan hari kerja normal, lembur, dan Sabtu."),
    ("Data Produksi", "Tabel data mentah. Tambah baris baru di bawah untuk input hari berikutnya."),
    ("", ""),
    ("KETERANGAN WARNA", ""),
    ("Biru", "Sel input manual (diisi operator)."),
    ("Kuning", "Sel otomatis (formula), jangan diketik manual."),
    ("", ""),
    ("RUMUS UTAMA", ""),
    ("OK", "Produksi - NG"),
    ("Total Tersedia", "Stok Awal + Stok Supplier + OK"),
    ("Kekurangan", "MAX(0, Planning - Total Tersedia)"),
    ("Stok Akhir", "MAX(0, Total Tersedia - Planning), jadi stok awal hari kerja berikutnya"),
    ("Pencapaian", "OK / Planning"),
    ("NG Rate", "NG / Produksi"),
]
for i, (k, v) in enumerate(info, start=4):
    a, b = ws.cell(row=i, column=1, value=k), ws.cell(row=i, column=2, value=v)
    a.font = F_BOLD if k and not v else Font(name=FONT, size=10, bold=True, color=NAVY)
    b.font = F_BODY
    b.alignment = Alignment(wrap_text=True, vertical="top")
    a.alignment = Alignment(vertical="top")
ws["A18"].fill = FILL_INPUT
ws["A19"].fill = FILL_CALC
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 95

# ---------------------------------------------------------------- Data Produksi
wd = wb.create_sheet("Data Produksi")
cols = ["Tanggal", "Bulan", "Hari", "Tipe Hari", "Lembur", "Part", "Line Tujuan",
        "Planning", "Stok Awal", "Stok Supplier", "Produksi", "NG",
        "OK", "Total Tersedia", "Kekurangan", "Stok Akhir", "Pencapaian", "Status"]
header(wd, 1, cols)
for r, row in enumerate(df.itertuples(index=False), start=2):
    vals = [row.tanggal.to_pydatetime(), f"=MONTH(A{r})", row.hari, row.tipe_hari, row.lembur,
            row.part, row.line_tujuan, row.planning, row.stok_awal, row.stok_supplier,
            row.produksi, row.ng,
            f"=K{r}-L{r}", f"=I{r}+J{r}+M{r}", f"=MAX(0,H{r}-N{r})", f"=MAX(0,N{r}-H{r})",
            f"=IFERROR(M{r}/H{r},0)", f'=IF(O{r}>0,"KURANG","AMAN")']
    for c, v in enumerate(vals, start=1):
        cell = wd.cell(row=r, column=c, value=v)
        is_input = c in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
        body(cell, font=F_INPUT if is_input else F_BODY)
        if c == 1:
            cell.number_format = "dd/mm/yyyy"
        elif c == 17:
            cell.number_format = "0.0%"
        elif c >= 8:
            cell.number_format = "#,##0"
for col in range(1, 19):
    wd.cell(row=1, column=col).fill = FILL_HEAD
widths = [12, 7, 9, 11, 8, 9, 11, 10, 10, 10, 10, 8, 9, 11, 11, 10, 11, 10]
for i, w in enumerate(widths, start=1):
    wd.column_dimensions[get_column_letter(i)].width = w
wd.freeze_panes = "A2"
wd.auto_filter.ref = f"A1:R{LAST}"
wd.conditional_formatting.add(f"R2:R{LAST}", CellIsRule(operator="equal", formula=['"KURANG"'],
                              font=Font(name=FONT, color="C00000", bold=True),
                              fill=PatternFill("solid", fgColor="F8CBAD")))
wd.conditional_formatting.add(f"Q2:Q{LAST}", CellIsRule(operator="lessThan", formula=["0.9"],
                              font=Font(name=FONT, color="C00000")))
for col, val in (("D", '"Hari Kerja,Sabtu"'), ("E", '"Ya,Tidak"'), ("F", '"' + ",".join(PARTS) + '"')):
    dv = DataValidation(type="list", formula1=val, allow_blank=False)
    wd.add_data_validation(dv)
    dv.add(f"{col}2:{col}{LAST + 500}")


def rng_(col):
    return f"{D}!${col}$2:${col}${LAST}"


# ---------------------------------------------------------------- Laporan Harian
wl = wb.create_sheet("Laporan Harian")
title(wl, "Laporan Harian Produksi", "Ganti tanggal di sel biru (C4), semua angka di bawah otomatis")
wl["B4"], wl["C4"] = "Tanggal", df["tanggal"].iloc[-1].to_pydatetime()
wl["B4"].font = F_BOLD
body(wl["C4"], "dd/mm/yyyy", F_INPUT, FILL_INPUT)
wl["E4"] = "Hari"
wl["E4"].font = F_BOLD
wl["F4"] = f'=IFERROR(INDEX({rng_("C")},MATCH(C4,{rng_("A")},0)),"Libur / tidak ada data")'
wl["F4"].font = F_BODY
wl["H4"] = "Lembur"
wl["H4"].font = F_BOLD
wl["I4"] = f'=IFERROR(INDEX({rng_("E")},MATCH(C4,{rng_("A")},0)),"-")'
wl["I4"].font = F_BODY
lh = ["No", "Part", "Line Tujuan", "Planning", "Stok Awal", "Stok Supplier",
      "Produksi", "NG", "OK", "Kekurangan", "Pencapaian", "Status"]
header(wl, 6, lh)
line_map = {"Part A": "Line 3"}
for i, p in enumerate(PARTS):
    r = 7 + i
    crit = f'{rng_("A")},$C$4,{rng_("F")},$B{r}'
    vals = [i + 1, p, line_map.get(p, "Line 2"),
            f"=SUMIFS({rng_('H')},{crit})", f"=SUMIFS({rng_('I')},{crit})",
            f"=SUMIFS({rng_('J')},{crit})", f"=SUMIFS({rng_('K')},{crit})",
            f"=SUMIFS({rng_('L')},{crit})", f"=SUMIFS({rng_('M')},{crit})",
            f"=SUMIFS({rng_('O')},{crit})", f"=IFERROR(I{r}/D{r},0)",
            f'=IF(D{r}=0,"-",IF(J{r}>0,"KURANG","AMAN"))']
    for c, v in enumerate(vals, start=1):
        body(wl.cell(row=r, column=c, value=v), "0.0%" if c == 11 else "#,##0")
    wl.cell(row=r, column=1).alignment = CENTER
tr = 12
wl.cell(row=tr, column=1, value="TOTAL")
wl.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=3)
for c in range(1, 13):
    cell = wl.cell(row=tr, column=c)
    body(cell, "#,##0", F_BOLD, FILL_TOTAL)
for c in range(4, 11):
    L = get_column_letter(c)
    wl.cell(row=tr, column=c, value=f"=SUM({L}7:{L}11)")
wl.cell(row=tr, column=11, value=f"=IFERROR(I{tr}/D{tr},0)").number_format = "0.0%"
wl.cell(row=tr, column=12, value=f'=IF(D{tr}=0,"-",IF(J{tr}>0,"KURANG","AMAN"))')
wl.conditional_formatting.add("L7:L12", CellIsRule(operator="equal", formula=['"KURANG"'],
                              font=Font(name=FONT, color="C00000", bold=True),
                              fill=PatternFill("solid", fgColor="F8CBAD")))
wl.conditional_formatting.add("L7:L12", CellIsRule(operator="equal", formula=['"AMAN"'],
                              font=Font(name=FONT, color="006100", bold=True),
                              fill=PatternFill("solid", fgColor="C6EFCE")))
wl["A14"] = "Catatan / Kendala:"
wl["A14"].font = F_BOLD
for r in range(15, 18):
    for c in range(1, 13):
        wl.cell(row=r, column=c).fill = FILL_INPUT
for i, w in enumerate([6, 10, 12, 11, 11, 12, 11, 8, 10, 12, 12, 11], start=1):
    wl.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------- Rekap Bulanan
wr = wb.create_sheet("Rekap Bulanan")
title(wr, "Rekap Bulanan", "Semua angka dihitung otomatis dari sheet Data Produksi")
rh = ["Bulan", "No. Bulan", "Hari Kerja", "Planning", "Produksi", "NG", "OK",
      "Kekurangan", "Pencapaian", "NG Rate", "Kejadian Kurang"]
header(wr, 4, rh)
for i, (m, lab) in enumerate(BULAN):
    r = 5 + i
    vals = [lab, m,
            f'=COUNTIFS({rng_("B")},$B{r},{rng_("F")},"Part A")',
            f"=SUMIFS({rng_('H')},{rng_('B')},$B{r})", f"=SUMIFS({rng_('K')},{rng_('B')},$B{r})",
            f"=SUMIFS({rng_('L')},{rng_('B')},$B{r})", f"=SUMIFS({rng_('M')},{rng_('B')},$B{r})",
            f"=SUMIFS({rng_('O')},{rng_('B')},$B{r})", f"=IFERROR(G{r}/D{r},0)",
            f"=IFERROR(F{r}/E{r},0)", f'=COUNTIFS({rng_("B")},$B{r},{rng_("O")},">0")']
    for c, v in enumerate(vals, start=1):
        fmt = "0.0%" if c == 9 else ("0.00%" if c == 10 else "#,##0")
        body(wr.cell(row=r, column=c, value=v), fmt)
tr = 11
wr.cell(row=tr, column=1, value="TOTAL")
for c in range(1, 12):
    body(wr.cell(row=tr, column=c), "#,##0", F_BOLD, FILL_TOTAL)
for c in (3, 4, 5, 6, 7, 8, 11):
    L = get_column_letter(c)
    wr.cell(row=tr, column=c, value=f"=SUM({L}5:{L}10)")
wr.cell(row=tr, column=9, value=f"=IFERROR(G{tr}/D{tr},0)").number_format = "0.0%"
wr.cell(row=tr, column=10, value=f"=IFERROR(F{tr}/E{tr},0)").number_format = "0.00%"
wr.cell(row=tr, column=2, value="")
wr.conditional_formatting.add("I5:I10", ColorScaleRule(start_type="min", start_color="F8696B",
                              mid_type="percentile", mid_value=50, mid_color="FFEB84",
                              end_type="max", end_color="63BE7B"))

wr["A13"] = "Pencapaian per Part per Bulan (OK / Planning)"
wr["A13"].font = Font(name=FONT, size=12, bold=True, color=NAVY)
header(wr, 14, ["Bulan"] + PARTS)
for i, (m, lab) in enumerate(BULAN):
    r = 15 + i
    body(wr.cell(row=r, column=1, value=lab))
    for j, p in enumerate(PARTS):
        c = 2 + j
        f = (f'=IFERROR(SUMIFS({rng_("M")},{rng_("B")},{m},{rng_("F")},"{p}")'
             f'/SUMIFS({rng_("H")},{rng_("B")},{m},{rng_("F")},"{p}"),0)')
        body(wr.cell(row=r, column=c, value=f), "0.0%")
wr.conditional_formatting.add("B15:F20", ColorScaleRule(start_type="num", start_value=0.85, start_color="F8696B",
                              mid_type="num", mid_value=0.95, mid_color="FFEB84",
                              end_type="num", end_value=1.0, end_color="63BE7B"))
for i, w in enumerate([12, 10, 10, 11, 11, 9, 11, 12, 12, 10, 14], start=1):
    wr.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------- Rekap Part
wp = wb.create_sheet("Rekap Part")
title(wp, "Rekap per Part", "Periode Mei - Oktober 2026")
ph = ["Part", "Line Tujuan", "Planning", "Produksi", "NG", "OK", "Kekurangan",
      "Pencapaian", "NG Rate", "Kejadian Kurang", "% Hari Kurang"]
header(wp, 4, ph)
for i, p in enumerate(PARTS):
    r = 5 + i
    crit = f'{rng_("F")},$A{r}'
    vals = [p, line_map.get(p, "Line 2"),
            f"=SUMIFS({rng_('H')},{crit})", f"=SUMIFS({rng_('K')},{crit})",
            f"=SUMIFS({rng_('L')},{crit})", f"=SUMIFS({rng_('M')},{crit})",
            f"=SUMIFS({rng_('O')},{crit})", f"=IFERROR(F{r}/C{r},0)", f"=IFERROR(E{r}/D{r},0)",
            f'=COUNTIFS({crit},{rng_("O")},">0")',
            f'=IFERROR(J{r}/COUNTIFS({crit}),0)']
    for c, v in enumerate(vals, start=1):
        fmt = {8: "0.0%", 9: "0.00%", 11: "0.0%"}.get(c, "#,##0")
        body(wp.cell(row=r, column=c, value=v), fmt)
tr = 10
wp.cell(row=tr, column=1, value="TOTAL")
for c in range(1, 12):
    body(wp.cell(row=tr, column=c), "#,##0", F_BOLD, FILL_TOTAL)
for c in (3, 4, 5, 6, 7, 10):
    L = get_column_letter(c)
    wp.cell(row=tr, column=c, value=f"=SUM({L}5:{L}9)")
wp.cell(row=tr, column=8, value=f"=IFERROR(F{tr}/C{tr},0)").number_format = "0.0%"
wp.cell(row=tr, column=9, value=f"=IFERROR(E{tr}/D{tr},0)").number_format = "0.00%"
wp.cell(row=tr, column=11, value=f"=IFERROR(J{tr}/COUNTA({rng_('F')}),0)").number_format = "0.0%"
wp.conditional_formatting.add("I5:I9", FormulaRule(formula=["I5=MAX($I$5:$I$9)"],
                              font=Font(name=FONT, color="C00000", bold=True),
                              fill=PatternFill("solid", fgColor="F8CBAD")))

wp["A12"] = "Pencapaian berdasarkan Tipe Hari"
wp["A12"].font = Font(name=FONT, size=12, bold=True, color=NAVY)
header(wp, 13, ["Tipe Hari", "Planning", "OK", "Pencapaian"])
tipe = [("Hari Kerja (normal)", '"Hari Kerja"', '"Tidak"'),
        ("Hari Kerja (lembur)", '"Hari Kerja"', '"Ya"'),
        ("Sabtu", '"Sabtu"', '"Tidak"')]
for i, (lab, t, l) in enumerate(tipe):
    r = 14 + i
    crit = f'{rng_("D")},{t},{rng_("E")},{l}'
    vals = [lab, f"=SUMIFS({rng_('H')},{crit})", f"=SUMIFS({rng_('M')},{crit})", f"=IFERROR(C{r}/B{r},0)"]
    for c, v in enumerate(vals, start=1):
        body(wp.cell(row=r, column=c, value=v), "0.0%" if c == 4 else "#,##0")
for i, w in enumerate([20, 12, 11, 11, 9, 11, 12, 12, 10, 14, 13], start=1):
    wp.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------- Dashboard
wdash = wb.create_sheet("Dashboard", 1)
title(wdash, "Dashboard Produksi · Mei - Oktober 2026", "Data simulasi · semua angka terhubung ke Data Produksi")
kpis = [
    ("Total Planning", "='Rekap Bulanan'!D11", "#,##0"),
    ("Total OK", "='Rekap Bulanan'!G11", "#,##0"),
    ("Pencapaian", "='Rekap Bulanan'!I11", "0.0%"),
    ("NG Rate", "='Rekap Bulanan'!J11", "0.00%"),
    ("Total Kekurangan", "='Rekap Bulanan'!H11", "#,##0"),
    ("Kejadian Kurang", "='Rekap Bulanan'!K11", "#,##0"),
]
for i, (lab, f, fmt) in enumerate(kpis):
    col = 1 + i * 2
    wdash.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
    wdash.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
    a = wdash.cell(row=4, column=col, value=lab)
    a.font, a.fill, a.alignment = F_HEAD, FILL_HEAD, CENTER
    b = wdash.cell(row=5, column=col, value=f)
    b.font = Font(name=FONT, size=18, bold=True, color=NAVY)
    b.fill, b.alignment, b.number_format = FILL_TOTAL, CENTER, fmt
    for rr in (4, 5):
        for cc in (col, col + 1):
            wdash.cell(row=rr, column=cc).border = BORDER
wdash.row_dimensions[4].height = 22
wdash.row_dimensions[5].height = 38
for c in range(1, 13):
    wdash.column_dimensions[get_column_letter(c)].width = 11

wdash["A7"] = "Insight utama"
wdash["A7"].font = Font(name=FONT, size=12, bold=True, color=NAVY)
insights = [
    "1. Part D punya NG rate tertinggi (sekitar 3x part lain) dan paling sering kekurangan stok.",
    "2. Agustus paling rendah karena Part D anjlok ke ~82% (skenario masalah mesin); September turun merata di semua part.",
    "3. Lembur menaikkan pencapaian di atas 100%, sedangkan Sabtu masuk paling rendah (~88%).",
    "4. Part C & E tidak pernah kurang, tapi stoknya menumpuk >1.000 pcs: order ke supplier bisa dikurangi.",
]
for i, t in enumerate(insights):
    wdash.cell(row=8 + i, column=1, value=t).font = F_BODY


def style_chart(ch, t, w=12.6, h=8.5):
    ch.title = t
    ch.width, ch.height = w, h
    ch.legend = None
    return ch


def labels(series, fmt=None):
    series.dLbls = DataLabelList()
    series.dLbls.showVal = True
    for a in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(series.dLbls, a, False)
    if fmt:
        series.dLbls.numFmt = fmt


lc = LineChart()
lc.add_data(Reference(wr, min_col=9, min_row=4, max_row=10), titles_from_data=True)
lc.set_categories(Reference(wr, min_col=1, min_row=5, max_row=10))
style_chart(lc, "Pencapaian per Bulan")
lc.y_axis.number_format = "0%"
lc.y_axis.scaling.min = 0.9
lc.y_axis.scaling.max = 1.0
lc.series[0].graphicalProperties.line.solidFill = NAVY
lc.series[0].graphicalProperties.line.width = 28000
lc.series[0].marker.symbol = "circle"
lc.series[0].smooth = False
labels(lc.series[0], "0.0%")
lc.series[0].dLbls.position = "t"
wdash.add_chart(lc, "A13")

bc = BarChart()
bc.add_data(Reference(wp, min_col=9, min_row=4, max_row=9), titles_from_data=True)
bc.set_categories(Reference(wp, min_col=1, min_row=5, max_row=9))
style_chart(bc, "NG Rate per Part")
bc.y_axis.number_format = "0.0%"
bc.series[0].graphicalProperties.solidFill = "C00000"
labels(bc.series[0], "0.0%")
wdash.add_chart(bc, "G13")

bc2 = BarChart()
bc2.add_data(Reference(wp, min_col=7, min_row=4, max_row=9), titles_from_data=True)
bc2.set_categories(Reference(wp, min_col=1, min_row=5, max_row=9))
style_chart(bc2, "Total Kekurangan per Part (pcs)")
bc2.series[0].graphicalProperties.solidFill = "ED7D31"
labels(bc2.series[0])
wdash.add_chart(bc2, "A30")

bc3 = BarChart()
bc3.add_data(Reference(wp, min_col=4, min_row=13, max_row=16), titles_from_data=True)
bc3.set_categories(Reference(wp, min_col=1, min_row=14, max_row=16))
style_chart(bc3, "Pencapaian per Tipe Hari")
bc3.y_axis.number_format = "0%"
bc3.y_axis.scaling.min = 0.8
bc3.series[0].graphicalProperties.solidFill = NAVY
labels(bc3.series[0], "0.0%")
wdash.add_chart(bc3, "G30")

for w in wb.worksheets:
    w.sheet_properties.tabColor = NAVY if w.title == "Dashboard" else None

out = ROOT / "Production_Monitoring_Dashboard.xlsx"
wb.save(out)
print("saved", out)
