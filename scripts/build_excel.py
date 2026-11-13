#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第6弾「不動産・空き家ダッシュボード」— Excel完成品の生成

data/processed/ の tidy CSV から output/realestate_vacancy_dashboard.xlsx を作る。
構成 (12シート): 表紙・使い方 / D1_KPI概要 / D2_都道府県別_空き家 /
D3_地価・着工 / D4_推移 / data_空き家_都道府県 / data_推移_全国 /
data_地価 / data_着工 / data_世帯 / 検証ログ / 出典・注記
グラフは openpyxl ネイティブグラフ。出典は各図表の下に注記する。
"""
import csv
import pathlib

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"
OUTF = ROOT / "output" / "realestate_vacancy_dashboard.xlsx"

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUBTITLE_FONT = Font(bold=True, size=12, color=NAVY)
KPI_LABEL_FONT = Font(bold=True, size=11, color="404040")
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
NOTE_FONT = Font(size=9, color="595959")
WARN_FONT = Font(size=10, bold=True, color="C00000")

SRC_HLS = "総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報）"
SRC_KOUJI = "国土交通省 地価公示（令和7年）"
SRC_CHAKKO = "国土交通省 建築着工統計調査（住宅着工統計）"


def rd(name):
    with open(PROC / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def iv(v):
    return int(float(v)) if v not in (None, "") else None


def fv(v):
    return float(v) if v not in (None, "") else None


def title(ws, cell, text):
    ws[cell] = text
    ws[cell].font = TITLE_FONT


def subtitle(ws, cell, text):
    ws[cell] = text
    ws[cell].font = SUBTITLE_FONT


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = NOTE_FONT


def hdr_row(ws, row, col0, labels):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=lab)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center")


def put_table(ws, row0, col0, headers, rows, fmts=None, widths=None):
    hdr_row(ws, row0, col0, headers)
    for j, r in enumerate(rows):
        for i, v in enumerate(r):
            c = ws.cell(row=row0 + 1 + j, column=col0 + i, value=v)
            if fmts and fmts[i]:
                c.number_format = fmts[i]
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(col0 + i)].width = w
    return row0 + 1 + len(rows)


def kpi(ws, row, col, label, value, fmt="#,##0", sub=None):
    ws.cell(row=row, column=col, value=label).font = KPI_LABEL_FONT
    c = ws.cell(row=row + 1, column=col, value=value)
    c.font = KPI_VALUE_FONT
    c.number_format = fmt
    if sub:
        ws.cell(row=row + 2, column=col, value=sub).font = NOTE_FONT


# ------------------------------------------------------------------ データ準備
snap_rows = rd("vacancy_snapshot_pref_2023.csv")
snap = {r["pref_code"]: r for r in snap_rows}
NAT = snap["00"]
trend = {int(r["year"]): r for r in rd("vacancy_trend_national.csv")}
comp = rd("vacancy_composition_national.csv")
struct = rd("vacancy_by_structure_pref_2023.csv")
dwell_pref = rd("dwellings_households_pref.csv")
hh_snap = {r["pref_code"]: r for r in rd("households_snapshot_pref_2023.csv")}
land_rate = rd("land_price_rate_pref_2025.csv")
land_snap = rd("land_price_snapshot_2025.csv")
land_trend = rd("land_price_trend_national.csv")
starts_pref = rd("starts_pref_annual.csv")
starts_nat = rd("starts_national_annual.csv")

VAC = iv(NAT["vacant_total"])
VAC_RATE = fv(NAT["vacancy_rate_pct_calc"])
VAC_OTHER = iv(NAT["vacant_other"])
TOT_DWELL = iv(NAT["total_dwellings"])
HH2023 = iv(trend[2023]["households"])
lr_nat = {(r["pref_code"], r["land_use"]): r for r in land_rate}
ls_nat = {(r["pref_code"], r["land_use"]): r for r in land_snap}
RESI_RATE25 = fv(lr_nat[("00", "住宅地")]["rate_2025_pct"])
COMM_RATE25 = fv(lr_nat[("00", "商業地")]["rate_2025_pct"])
ALL_RATE25 = fv(ls_nat[("00", "全用途")]["rate_2025_pct"])
RESI_PRICE25 = iv(ls_nat[("00", "住宅地")]["avg_price_2025_yen_m2"])
sn = {(int(r["year"]), r["year_type"]): r for r in starts_nat}
STARTS_CY25 = iv(sn[(2025, "暦年")]["total_starts"])
STARTS_CY24 = iv(sn[(2024, "暦年")]["total_starts"])
STARTS_FY24 = iv(sn[(2024, "年度")]["total_starts"])
CY_YOY = (STARTS_CY25 / STARTS_CY24 - 1) * 100
rates_sorted = sorted(
((fv(r["vacancy_rate_pct_calc"]), r["pref_jp"]) for c, r in snap.items() if c!= "00"),
reverse=True)
TOP_RATE_PREF, TOP_RATE = rates_sorted[0][1], rates_sorted[0][0]
LOW_RATE_PREF, LOW_RATE = rates_sorted[-1][1], rates_sorted[-1][0]

wb = Workbook()

# ================================================================== 表紙・使い方
ws = wb.active
ws.title = "表紙・使い方"
ws.column_dimensions["B"].width = 110
title(ws, "B2", "不動産・空き家ダッシュボード（都道府県レベル）")
ws["B4"] = "都道府県レベルで「空き家ストック × 地価 × 住宅着工フロー」を見る市場ダッシュボードのサンプルです。"
ws["B5"] = "単一エリアの出店診断（サンプル第1弾）とは異なり、都道府県単位の広域市場概観に特化しています。"
ws["B7"] = "対象読者： 不動産事業者、自治体の空き家対策担当、住宅関連事業者"
ws["B7"].font = KPI_LABEL_FONT
ws["B9"] = "【鮮度の限界について（重要）】"
ws["B9"].font = WARN_FONT
ws["B10"] = "空き家の主データである住宅・土地統計調査は5年周期の調査で、最新は2023年（令和5年）10月1日現在です。"
ws["B11"] = "2024年以降の空き家動向はこの統計では追えません（次回調査は2028年予定）。直近の市場動向は、"
ws["B12"] = "毎年更新される地価公示（2025年）と建築着工統計（2025年）で補完する設計にしています。"
for r_, t in [
(14, "シート構成"),
(15, "D1_KPI概要 ― 空き家・地価・着工の主要指標（最新値）"),
(16, "D2_都道府県別_空き家 ― 47都道府県の空き家数・空き家率・建て方別構成"),
(17, "D3_地価・着工 ― 都道府県別の地価変動率と住宅着工戸数"),
(18, "D4_推移 ― 住宅ストックと世帯の長期推移、空き家の種類別推移、地価・着工の推移"),
(19, "data_* ― 整形済みデータ（出典列つき）／検証ログ ― 公表値との突合結果／出典・注記 ― 定義・採用理由・不採用理由"),
(21, "使い方： D1→D2→D3→D4 の順に読むと「現状（ストック）→ 地域差 → 直近の動き（価格・フロー）→ 長期トレンド」が追えます。"),
(22, "数値の出典は各シート下部と「出典・注記」に明記しています。率は原則として本ダッシュボードの計算値です。"),
]:
    ws.cell(row=r_, column=2, value=t)
    if r_ == 14:
        ws.cell(row=r_, column=2).font = SUBTITLE_FONT

# ================================================================== D1 KPI概要
ws = wb.create_sheet("D1_KPI概要")
title(ws, "B2", "D1 KPI概要 ― 空き家ストック、地価、着工フロー（最新値）")
kpi(ws, 4, 2, "空き家数（2023年）", VAC, "#,##0", "過去最多（2018年 8,488,600戸から増加）")
kpi(ws, 4, 4, "空き家率（2023年）", VAC_RATE, "0.0", "％（公表値13.8％と一致・計算値）")
kpi(ws, 4, 6, "「その他」空き家（2023年）", VAC_OTHER, "#,##0", "賃貸・売却用・二次的住宅を除く空き家")
kpi(ws, 8, 2, "総住宅数（2023年）", TOT_DWELL, "#,##0", f"総世帯数 {HH2023:,}世帯")
kpi(ws, 8, 4, "地価公示 住宅地 変動率（2025年）", RESI_RATE25, "0.0", "％（全国平均・対前年）")
kpi(ws, 8, 6, "新設住宅着工戸数（2025年・暦年）", STARTS_CY25, "#,##0", f"前年比 {CY_YOY:+.1f}％")
ws["B12"] = (f"空き家率が最も高いのは{TOP_RATE_PREF}（{TOP_RATE:.1f}％）、最も低いのは{LOW_RATE_PREF}（{LOW_RATE:.1f}％）。"
             f"地価公示2025は全用途 {ALL_RATE25:+.1f}％、商業地 {COMM_RATE25:+.1f}％と上昇が続く。"
             f"住宅地の平均価格は {RESI_PRICE25:,}円/m²。")
note(ws, "B13", "出典： " + SRC_HLS + "／" + SRC_KOUJI + "／" + SRC_CHAKKO)

subtitle(ws, "B15", "空き家数と空き家率の推移（全国、1978〜2023年）")
trows = []
for y in range(1978, 2024, 5):
    r = trend[y]
    trows.append([y, iv(r["vacant_total"]), fv(r["vacancy_rate_pct_calc"])])
end = put_table(ws, 16, 2, ["年", "空き家数（戸）", "空き家率（％・計算値）"], trows,
                fmts=[None, "#,##0", "0.00"], widths=[10, 18, 16])
ch = BarChart()
ch.title = "空き家数（棒）と空き家率（線）の推移"
data = Reference(ws, min_col=3, min_row=16, max_row=end - 1)
cats = Reference(ws, min_col=2, min_row=17, max_row=end - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.y_axis.title = "空き家数（戸）"
ch.height, ch.width = 8, 16
ln = LineChart()
data2 = Reference(ws, min_col=4, min_row=16, max_row=end - 1)
ln.add_data(data2, titles_from_data=True)
ln.y_axis.axId = 200
ln.y_axis.title = "空き家率（％）"
ch += ln
ws.add_chart(ch, f"B{end + 1}")
note(ws, f"B{end + 17}", "出典： " + SRC_HLS + " 第3表・第4表。空き家率は空き家数÷総住宅数の計算値。")

subtitle(ws, f"B{end + 19}", "空き家の種類別内訳（全国、2023年）")
crows = [[r["vacancy_type"], iv(r["total"]),
          round(iv(r["total"]) / VAC * 100, 1)] for r in comp]
end2 = put_table(ws, end + 20, 2, ["種類", "戸数", "構成比（％）"], crows,
                 fmts=[None, "#,##0", "0.0"], widths=[40, 16, 12])
note(ws, f"B{end2 + 1}", "出典： " + SRC_HLS + " 第5表。「その他」は表記上「賃貸・売却用及び二次的住宅を除く空き家」。内訳合計と総数は原表の丸めで最大100戸差が生じる。")

# ================================================================== D2 都道府県別_空き家
ws = wb.create_sheet("D2_都道府県別_空き家")
title(ws, "B2", "D2 都道府県別 空き家 ― 空き家数・空き家率・建て方別（2023年）")
prows = []
for r in snap_rows:
    c = r["pref_code"]
    if c == "00":
        continue
    hh = hh_snap[c]
    prows.append([r["pref_jp"], iv(r["total_dwellings"]), iv(hh["households_total"]),
                  iv(r["vacant_total"]), fv(r["vacancy_rate_pct_calc"]),
                  iv(r["vacant_other"]), fv(r["vacant_other_rate_pct_calc"]),
                  iv(r["vacant_for_rent"])])
prows.sort(key=lambda x: -x[4])
end = put_table(ws, 4, 2,
                ["都道府県", "総住宅数", "世帯数", "空き家数", "空き家率（％）",
                 "その他空き家", "その他率（％）", "賃貸用空き家"],
                prows, fmts=[None, "#,##0", "#,##0", "#,##0", "0.00", "#,##0", "0.00", "#,##0"],
                widths=[12, 14, 14, 14, 12, 14, 12, 14])
note(ws, f"B{end + 1}", "出典： " + SRC_HLS + " 第1-2表・第2-2表。空き家率・その他率は計算値。表は空き家率の降順。")

subtitle(ws, f"B{end + 3}", "空き家率 上位10県・下位10県（％）")
rank_rows = [[p, v] for v, p in rates_sorted[:10]] + [[p, v] for v, p in rates_sorted[-10:]]
endr = put_table(ws, end + 4, 2, ["都道府県", "空き家率（％）"], rank_rows,
                 fmts=[None, "0.00"], widths=[12, 12])
ch = BarChart()
ch.type = "bar"
ch.title = "空き家率 上位10県・下位10県（2023年）"
data = Reference(ws, min_col=3, min_row=end + 4, max_row=endr - 1)
cats = Reference(ws, min_col=2, min_row=end + 5, max_row=endr - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 10, 16
ws.add_chart(ch, f"B{endr + 1}")

subtitle(ws, f"B{endr + 22}", "建て方別 空き家数（全国、2023年）")
nat_struct = {r["structure"]: iv(r["vacant_dwellings"]) for r in struct
              if r["pref_code"] == "00" and r["structure"] != "総数"}
srows = [[k, v, round(v / VAC * 100, 1)] for k, v in nat_struct.items()]
ends = put_table(ws, endr + 23, 2, ["建て方", "空き家数", "構成比（％）"], srows,
                 fmts=[None, "#,##0", "0.0"], widths=[12, 14, 12])
note(ws, f"B{ends + 1}", "出典： " + SRC_HLS + " 第4表。空き家の過半は共同住宅だが、「その他」空き家は一戸建に偏る（D1の種類別内訳・data参照）。")

# ================================================================== D3 地価・着工
ws = wb.create_sheet("D3_地価・着工")
title(ws, "B2", "D3 地価・着工 ― 都道府県別の地価変動率（2025年）と住宅着工戸数")
lrows = []
for c in [r["pref_code"] for r in land_rate if r["land_use"] == "住宅地"]:
    if c == "00":
        continue
    lrows.append([lr_nat[(c, "住宅地")]["pref_jp"],
                  fv(lr_nat[(c, "住宅地")]["rate_2025_pct"]),
                  fv(lr_nat[(c, "商業地")]["rate_2025_pct"]) if (c, "商業地") in lr_nat else None,
                  iv(ls_nat[(c, "住宅地")]["avg_price_2025_yen_m2"]) if (c, "住宅地") in ls_nat else None])
lrows.sort(key=lambda x: -(x[1] or -99))
end = put_table(ws, 4, 2,
                ["都道府県", "住宅地 変動率2025（％）", "商業地 変動率2025（％）", "住宅地 平均価格2025（円/m²）"],
                lrows, fmts=[None, "0.0", "0.0", "#,##0"], widths=[12, 16, 16, 18])
note(ws, f"B{end + 1}", "出典： " + SRC_KOUJI + " 第5表・価格/変動率推移表。平均価格は標準地の平均で、取引価格ではない。表は住宅地変動率の降順。")

subtitle(ws, f"B{end + 3}", "住宅地 変動率 上位15県（2025年、％）")
top15 = lrows[:15]
endt = put_table(ws, end + 4, 2, ["都道府県", "住宅地変動率（％）"], [[r[0], r[1]] for r in top15],
                 fmts=[None, "0.0"], widths=[12, 14])
ch = BarChart()
ch.type = "bar"
ch.title = "地価公示 住宅地変動率 上位15県（2025年）"
data = Reference(ws, min_col=3, min_row=end + 4, max_row=endt - 1)
cats = Reference(ws, min_col=2, min_row=end + 5, max_row=endt - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 9, 16
ws.add_chart(ch, f"B{endt + 1}")

srow0 = endt + 20
subtitle(ws, f"B{srow0}", "新設住宅着工戸数（全国、暦年 2014〜2025年）")
cy = {}
for r in starts_pref:
    if r["pref_code"] == "00":
        cy[int(r["year"])] = iv(r["new_housing_starts"])
crows = [[y, cy[y], round((cy[y] / cy[y - 1] - 1) * 100, 1) if y - 1 in cy else None]
         for y in sorted(cy)]
ends2 = put_table(ws, srow0 + 1, 2, ["年", "着工戸数", "前年比（％）"], crows,
                  fmts=[None, "#,##0", "0.0"], widths=[8, 14, 12])
ch = LineChart()
ch.title = "新設住宅着工戸数の推移（全国・暦年）"
data = Reference(ws, min_col=3, min_row=srow0 + 1, max_row=ends2 - 1)
cats = Reference(ws, min_col=2, min_row=srow0 + 2, max_row=ends2 - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 8, 16
ws.add_chart(ch, f"B{ends2 + 1}")
note(ws, f"B{ends2 + 17}", "出典： " + SRC_CHAKKO + " 都道府県別戸数 時系列（年計）。2025年は建築基準法改正（4月）を控えた駆け込みと反動を含む。")

# ================================================================== D4 推移
ws = wb.create_sheet("D4_推移")
title(ws, "B2", "D4 推移 ― ストック・価格・フローの長期推移")
subtitle(ws, "B4", "住宅数と世帯数（全国、1958〜2023年）")
trows = [[y, iv(trend[y]["total_dwellings"]), iv(trend[y]["households"])]
         for y in sorted(trend)]
end = put_table(ws, 5, 2, ["年", "総住宅数", "総世帯数"], trows,
                fmts=[None, "#,##0", "#,##0"], widths=[8, 14, 14])
ch = LineChart()
ch.title = "総住宅数と総世帯数の推移"
data = Reference(ws, min_col=3, min_row=5, max_row=end - 1)
data2 = Reference(ws, min_col=4, min_row=5, max_row=end - 1)
cats = Reference(ws, min_col=2, min_row=6, max_row=end - 1)
ch.add_data(data, titles_from_data=True)
ch.add_data(data2, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 8, 16
ws.add_chart(ch, f"B{end + 1}")
note(ws, f"B{end + 17}", "出典： " + SRC_HLS + " 第3表。住宅数が世帯数を上回り続けており、その差が空き家の受け皿になっている。")

r0 = end + 19
subtitle(ws, f"B{r0}", "空き家の種類別推移（全国、1983〜2023年）")
krows = []
for y in [1983, 1988, 1993, 1998, 2003, 2008, 2013, 2018, 2023]:
    r = trend[y]
    krows.append([y, iv(r["vacant_other"]), iv(r["vacant_for_rent"]),
                  iv(r["vacant_for_sale"]), iv(r["vacant_second"])])
endk = put_table(ws, r0 + 1, 2,
                 ["年", "その他（賃貸・売却・二次を除く）", "賃貸用", "売却用", "二次的住宅"],
                 krows, fmts=[None, "#,##0", "#,##0", "#,##0", "#,##0"],
                 widths=[8, 18, 14, 12, 12])
ch = BarChart()
ch.type = "col"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "空き家の種類別推移（積み上げ）"
data = Reference(ws, min_col=3, min_row=r0 + 1, max_row=endk - 1)
cats = Reference(ws, min_col=2, min_row=r0 + 2, max_row=endk - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 8, 16
ws.add_chart(ch, f"B{endk + 1}")
note(ws, f"B{endk + 17}", "出典： " + SRC_HLS + " 第4表。売却用は2003年から分離集計（それ以前は賃貸用に含まれる）。「その他」の増加が空き家問題の中核。")

r0 = endk + 19
subtitle(ws, f"B{r0}", "地価公示 平均価格の推移（全国、1975〜2025年）")
lp = {}
for r in land_trend:
    lp.setdefault(r["land_use"], {})[int(r["year"])] = iv(r["avg_price_yen_m2"])
prows = [[y, lp["住宅地"].get(y), lp["商業地"].get(y)]
         for y in [1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025]]
endp = put_table(ws, r0 + 1, 2, ["年", "住宅地 平均価格（円/m²）", "商業地 平均価格（円/m²）"],
                 prows, fmts=[None, "#,##0", "#,##0"], widths=[8, 18, 18])
note(ws, f"B{endp + 1}", "出典： " + SRC_KOUJI + " 価格推移表（5年刻み＋2025年）。完全な年次系列は data_地価 を参照。")

r0 = endp + 3
subtitle(ws, f"B{r0}", "新設住宅着工戸数（全国、年度 1994〜2024年度）")
frows = []
for y in range(1994, 2025):
    r = sn[(y, "年度")]
    frows.append([y, iv(r["total_starts"]), iv(r["owner_occupied"]),
                  iv(r["for_rent"]), iv(r["for_sale"])])
endf = put_table(ws, r0 + 1, 2, ["年度", "総計", "持家", "貸家", "分譲住宅"], frows,
                 fmts=[None, "#,##0", "#,##0", "#,##0", "#,##0"],
                 widths=[8, 14, 12, 12, 12])
ch = BarChart()
ch.title = "新設住宅着工戸数（年度・利用関係別は表参照）"
data = Reference(ws, min_col=3, min_row=r0 + 1, max_row=endf - 1)
cats = Reference(ws, min_col=2, min_row=r0 + 2, max_row=endf - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 8, 16
ws.add_chart(ch, f"B{endf + 1}")
note(ws, f"B{endf + 17}", "出典： " + SRC_CHAKKO + " 時系列表（全国）。2025年度は年度途中のため未掲載（値は data_着工 に未確定として収録）。")

# ================================================================== data シート群
def dump(sheet_name, csv_name):
    rows = rd(csv_name)
    wsx = wb.create_sheet(sheet_name)
    if not rows:
        return
    headers = list(rows[0].keys())
    hdr_row(wsx, 1, 1, headers)
    for j, r in enumerate(rows):
        for i, h in enumerate(headers):
            v = r[h]
            if v in (None, ""):
                continue
            try:
                f_ = float(v)
                v = int(f_) if f_ == int(f_) and abs(f_) >= 1 else f_
            except (TypeError, ValueError):
                pass
            wsx.cell(row=2 + j, column=1 + i, value=v)
    wsx.freeze_panes = "A2"


dump("data_空き家_都道府県", "vacancy_snapshot_pref_2023.csv")
dump("data_推移_全国", "vacancy_trend_national.csv")

ws = wb.create_sheet("data_地価")
hdr_row(ws, 1, 1, ["データ", "pref_code", "都道府県", "用途", "値1", "値2", "出典"])
row = 2
for r in land_rate:
    ws.cell(row=row, column=1, value="変動率（第5表）")
    ws.cell(row=row, column=2, value=r["pref_code"])
    ws.cell(row=row, column=3, value=r["pref_jp"])
    ws.cell(row=row, column=4, value=r["land_use"])
    ws.cell(row=row, column=5, value=fv(r["rate_2024_pct"]))
    ws.cell(row=row, column=6, value=fv(r["rate_2025_pct"]))
    ws.cell(row=row, column=7, value=r["source"])
    row += 1
for r in land_snap:
    ws.cell(row=row, column=1, value="平均価格・変動率2025（推移表）")
    ws.cell(row=row, column=2, value=r["pref_code"])
    ws.cell(row=row, column=3, value=r["pref_jp"])
    ws.cell(row=row, column=4, value=r["land_use"])
    ws.cell(row=row, column=5, value=iv(r["avg_price_2025_yen_m2"]))
    ws.cell(row=row, column=6, value=fv(r["rate_2025_pct"]))
    ws.cell(row=row, column=7, value=r["source"])
    row += 1
for r in land_trend:
    ws.cell(row=row, column=1, value="全国推移")
    ws.cell(row=row, column=2, value=r["year"])
    ws.cell(row=row, column=4, value=r["land_use"])
    ws.cell(row=row, column=5, value=iv(r["avg_price_yen_m2"]))
    ws.cell(row=row, column=6, value=fv(r["rate_pct"]))
    ws.cell(row=row, column=7, value=r["source"])
    row += 1
ws.freeze_panes = "A2"

ws = wb.create_sheet("data_着工")
hdr_row(ws, 1, 1, ["データ", "pref_code/年", "都道府県/年種別", "年", "戸数/総計", "持家", "貸家", "分譲", "出典"])
row = 2
for r in starts_pref:
    ws.cell(row=row, column=1, value="都道府県別・暦年")
    ws.cell(row=row, column=2, value=r["pref_code"])
    ws.cell(row=row, column=3, value=r["pref_jp"])
    ws.cell(row=row, column=4, value=iv(r["year"]))
    ws.cell(row=row, column=5, value=iv(r["new_housing_starts"]))
    ws.cell(row=row, column=9, value=r["source"])
    row += 1
for r in starts_nat:
    ws.cell(row=row, column=1, value="全国")
    ws.cell(row=row, column=2, value=iv(r["year"]))
    ws.cell(row=row, column=3, value=r["year_type"])
    ws.cell(row=row, column=4, value=iv(r["year"]))
    ws.cell(row=row, column=5, value=iv(r["total_starts"]))
    ws.cell(row=row, column=6, value=iv(r["owner_occupied"]))
    ws.cell(row=row, column=7, value=iv(r["for_rent"]))
    ws.cell(row=row, column=8, value=iv(r["for_sale"]))
    ws.cell(row=row, column=9, value=r["source"] + "（" + r["status"] + "）")
    row += 1
ws.freeze_panes = "A2"

ws = wb.create_sheet("data_世帯")
hdr_row(ws, 1, 1, ["pref_code", "都道府県", "year", "総住宅数", "総世帯数",
                   "世帯総数（第2-2表・2023）", "1世帯当たり人員（2023）", "出典"])
by_pref_year = {(r["pref_code"], int(r["year"])): r for r in dwell_pref}
row = 2
for c in ["00"] + [f"{i:02d}" for i in range(1, 48)]:
    for y in [2003, 2008, 2013, 2018, 2023]:
        r = by_pref_year.get((c, y))
        if not r:
            continue
        ws.cell(row=row, column=1, value=c)
        ws.cell(row=row, column=2, value=r["pref_jp"])
        ws.cell(row=row, column=3, value=y)
        ws.cell(row=row, column=4, value=iv(r["total_dwellings"]))
        ws.cell(row=row, column=5, value=iv(r["households"]))
        if y == 2023:
            ws.cell(row=row, column=6, value=iv(hh_snap[c]["households_total"]))
            ws.cell(row=row, column=7, value=fv(hh_snap[c]["persons_per_household_calc"]))
        ws.cell(row=row, column=8, value=r["source"])
        row += 1
ws.freeze_panes = "A2"

# ================================================================== 検証ログ
ws = wb.create_sheet("検証ログ")
title(ws, "B2", "検証ログ ― 公表値との突合（2026-11-13 実施）")
vrows = [
["空き家総数（2023・全国）", "9,001,600戸", "900万2千戸（基本集計 確報・結果概要）", "一致", "公表は千戸単位に丸め"],
["空き家率（2023・全国）", "13.84％（計算値）", "13.8％（同上）", "一致", "小数第1位まで一致"],
["賃貸・売却用及び二次的住宅を除く空き家（2023）", "3,856,000戸", "385万6千戸（同上）", "一致", ""],
["総住宅数（2023・全国）", "65,046,700戸", "6,502万戸（同上）", "一致", ""],
["空き家総数（2018・全国）", "8,488,600戸", "848万9千戸（同上）", "一致", ""],
["総世帯数（2023・全国）", "56,215,300世帯", "第3表と第2-2表の二源が完全一致", "一致", "定義： 住宅に居住する世帯の総数"],
["空き家 種類別内訳の合計（2023）", "9,001,500戸", "空き家総数 9,001,600戸", "差100戸", "原表の丸めによる差（各値は推計・丸め済み）"],
["建て方別合計 vs 空き家総数（都道府県別）", "差は最大100戸", "第4表と第1-2表の照合", "整合", "原表の丸めの範囲内"],
["地価公示2025 全国変動率", "住宅地2.1／商業地3.9／工業地4.8／全用途2.7％", "第5表と価格・変動率推移表の二源", "一致", "全用途2.7％は公表値とも一致"],
["地価公示2025 都道府県別変動率", "141件を二源照合", "第5表 vs 推移表", "全件一致", ""],
["着工 暦年合計（2023）", "819,623戸", "都道府県別年計表 vs 全国月次の自前集計", "完全一致", ""],
["着工 暦年合計（2024）", "792,195戸", "同上", "完全一致", ""],
["着工 暦年合計（2025）", "740,667戸", "同上", "完全一致", ""],
["着工 年度計（2024年度）", "816,388戸", "時系列表（全国）", "採用", "2025年度分 711,171戸は年度途中の未確定値として区別"],
]
end = put_table(ws, 4, 2, ["項目", "本ダッシュボードの値", "突合相手", "結果", "備考"],
vrows, widths=[36, 34, 38, 10, 44])
note(ws, f"B{end + 1}", "突合相手の公表値は、総務省「令和5年住宅・土地統計調査 住宅及び世帯に関する基本集計（確報集計）結果」（2024-09-25公表）等の一次資料による。")

# ================================================================== 出典・注記
ws = wb.create_sheet("出典・注記")
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 95
title(ws, "B2", "出典・注記 ― 定義、採用理由・不採用理由、限界")
row = 4
subtitle(ws, f"B{row}", "採用したデータと採用理由")
row += 1
adopt = [
["住宅・土地統計調査（総務省統計局、令和5年・基本集計 確報）",
"空き家数・空き家率の公式定義を持つ唯一の専門調査。都道府県別・空き家の種類別・建て方別まで一貫した定義で比較できるため、本ダッシュボードの主データとして採用。速報ではなく確報を用いた。"],
["地価公示（国土交通省、令和7年）",
"公的な地価指標で毎年公表される。都道府県別・用途別の平均価格と対前年変動率が揃い、5年周期の住調を年次で補完できるため採用。第5表と推移表の二源で相互検証できることも採用理由。"],
["建築着工統計調査／住宅着工統計（国土交通省）",
"住宅フローの公式統計。都道府県別・暦年の着工戸数（e-Stat時系列表）が長期で揃い、全国の年度・月次系列と相互検証できるため採用。"],
["世帯数（住宅・土地統計調査 第2-2表・第3表）",
"空き家（ストック）の需要側の分母として、住宅総数との対比に必要。住調の中で定義が揃うため、外部の人口統計ではなく同一調査から取得した。"],
]
for a, b in adopt:
    ws.cell(row=row, column=2, value=a).font = KPI_LABEL_FONT
    c = ws.cell(row=row, column=3, value=b)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 42
    row += 1
row += 1
subtitle(ws, f"B{row}", "採用しなかったデータと不採用理由")
row += 1
reject = [
["住宅・土地統計調査 速報集計（住宅数概数集計）",
"早期公表用の概数で、確報で値が更新される。確報（基本集計）が公表済みのため不採用。"],
["住宅・土地統計調査 市区町村別結果",
"本サンプルは都道府県粒度に限定したため不採用。市区町村粒度は自治体・事業者の買い手価値が高いため、商品化時は市区町村別レポートとして有料オプションで提供する想定。"],
["建築着工統計 構造別・建て方別・利用関係別の年別詳細ファイル",
"粒度（年×都道府県×構造×建て方×利用関係）が本ダッシュボードの粒度を大きく上回り、主要指標は時系列表で代替できるため本編では不採用。商品化時は構造別・利用関係別の深掘り分析として有料オプション提供の想定。"],
["地価公示 市区町村別の平均価格・変動率表",
"都道府県粒度に限定したため不採用。商品化時は市区町村別レポートとして有料オプション提供の想定。"],
["地価公示 地点別（標準地）生データ",
"都道府県集計は公表表（第5表・推移表）で代替できるため不採用。地点レベルの立地分析は商品化時の有料オプション想定。"],
["国勢調査の住宅・世帯項目",
"空き家の種類別内訳は国勢調査では取れず、空き家の定義は住調が専門であるため不採用。"],
["人口動態（住民基本台帳・人口推計）",
"空き家率の公式分母は住宅総数であり、需要側は世帯数で代表できるため本編では不採用。商品化時は人口動態とのクロス分析を有料オプション想定。"],
["建築着工統計 都道府県別 月次",
"年次に統一した方が住調（5年）・地価（年次）と粒度が揃うため本編では不採用。全国の月次は暦年集計の検証にのみ使用した。"],
]
for a, b in reject:
    ws.cell(row=row, column=2, value=a).font = KPI_LABEL_FONT
    c = ws.cell(row=row, column=3, value=b)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 42
    row += 1
row += 1
subtitle(ws, f"B{row}", "定義・利用上の注意")
row += 1
defs = [
"「空き家」（住調の定義）： 居住世帯のいない住宅のうち、一時現在者のみの住宅と建築中の住宅を除いたもの。種類は 賃貸用／売却用／二次的住宅／その他（賃貸・売却用及び二次的住宅を除く空き家＝長期不在・取壊し予定など）。",
"空き家率＝空き家数÷総住宅数（本ダッシュボードの計算値）。公表値13.8％と小数第1位で一致する。",
"【鮮度限界】住調は5年周期で最新は2023年10月1日現在。2024年以降の空き家動向は追えない（次回2028年調査予定）。地価・着工は年次で補完しているが、空き家そのものの代替指標ではない。",
"地価公示の平均価格は標準地の平均価格であり、実際の取引価格とは異なる。変動率は継続地点の対前年変動率の平均。",
"着工統計の2025年度（年度計）は年度途中の未確定値。暦年2025年には建築基準法改正（2025年4月）を控えた駆け込み（3月の急増）とその反動が含まれる。",
"住調の数値は標本調査に基づく推計値で、原表の丸めにより内訳の合計が総数と一致しない場合がある（最大100戸程度）。",
]
for d in defs:
    c = ws.cell(row=row, column=2, value="・" + d)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    ws.row_dimensions[row].height = 30
    row += 1

OUTF.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUTF)
print("saved", OUTF)
