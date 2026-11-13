#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第6弾「不動産・空き家ダッシュボード」 tidy CSV 生成スクリプト

入力: data/raw/ の公開ダウンロードファイル群（e-Stat / 国土交通省 / 土地情報センター）
出力: data/processed/ の tidy CSV 群

方針:
- 数値は原表の値をそのまま使い、推測で埋めない。欠測は空欄。
- 空き家率などの率は「計算値」であることを列名と注記で明示する。
- 各 CSV に出典列 (source) を付ける。
"""
import csv
import os
import re
import warnings

import openpyxl
import xlrd

warnings.filterwarnings("ignore")

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "data", "raw")
OUT = os.path.join(BASE, "data", "processed")
os.makedirs(OUT, exist_ok=True)

PREFS = [
    ("01", "北海道", "Hokkaido"), ("02", "青森県", "Aomori"), ("03", "岩手県", "Iwate"),
    ("04", "宮城県", "Miyagi"), ("05", "秋田県", "Akita"), ("06", "山形県", "Yamagata"),
    ("07", "福島県", "Fukushima"), ("08", "茨城県", "Ibaraki"), ("09", "栃木県", "Tochigi"),
    ("10", "群馬県", "Gunma"), ("11", "埼玉県", "Saitama"), ("12", "千葉県", "Chiba"),
    ("13", "東京都", "Tokyo"), ("14", "神奈川県", "Kanagawa"), ("15", "新潟県", "Niigata"),
    ("16", "富山県", "Toyama"), ("17", "石川県", "Ishikawa"), ("18", "福井県", "Fukui"),
    ("19", "山梨県", "Yamanashi"), ("20", "長野県", "Nagano"), ("21", "岐阜県", "Gifu"),
    ("22", "静岡県", "Shizuoka"), ("23", "愛知県", "Aichi"), ("24", "三重県", "Mie"),
    ("25", "滋賀県", "Shiga"), ("26", "京都府", "Kyoto"), ("27", "大阪府", "Osaka"),
    ("28", "兵庫県", "Hyogo"), ("29", "奈良県", "Nara"), ("30", "和歌山県", "Wakayama"),
    ("31", "鳥取県", "Tottori"), ("32", "島根県", "Shimane"), ("33", "岡山県", "Okayama"),
    ("34", "広島県", "Hiroshima"), ("35", "山口県", "Yamaguchi"), ("36", "徳島県", "Tokushima"),
    ("37", "香川県", "Kagawa"), ("38", "愛媛県", "Ehime"), ("39", "高知県", "Kochi"),
    ("40", "福岡県", "Fukuoka"), ("41", "佐賀県", "Saga"), ("42", "長崎県", "Nagasaki"),
    ("43", "熊本県", "Kumamoto"), ("44", "大分県", "Oita"), ("45", "宮崎県", "Miyazaki"),
    ("46", "鹿児島県", "Kagoshima"), ("47", "沖縄県", "Okinawa"),
]
PREF_BY_SHORT = {jp.replace("県", "").replace("府", "").replace("都", ""): (c, jp, en)
                 for c, jp, en in PREFS}
PREF_BY_SHORT["東京"] = ("13", "東京都", "Tokyo")
PREF_BY_SHORT["大阪"] = ("27", "大阪府", "Osaka")
PREF_BY_SHORT["京都"] = ("26", "京都府", "Kyoto")
PREF_BY_SHORT["北海道"] = ("01", "北海道", "Hokkaido")


def num(v):
    """e-Stat表のセル値を int/float に。'-' や空欄は None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "－", "…", "x"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def write_csv(name, header, rows):
    path = os.path.join(OUT, name)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"wrote {name}: {len(rows)} rows")
    return rows


SRC_HLS = "総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計"
SRC_KOUJI = "国土交通省 地価公示 令和7年(2025)"
SRC_CHAKKO = "国土交通省 建築着工統計調査(住宅着工統計)"

# ----------------------------------------------------------------------
# 1. 空き家スナップショット 2023 都道府県別 (第1-2表 令和5年シート)
#    列: 5=住宅総数 6=居住世帯あり 7=居住世帯なし総数 8=一時現在者のみ
#        9=空き家総数 10=賃貸・売却用及び二次的住宅を除く 11=賃貸用 12=売却用
#        13=二次的住宅 14=別荘 15=建築中 16=住宅以外で人が居住する建物総数
# ----------------------------------------------------------------------
def build_snapshot():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t1-2_dwellings_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb["令和5年"]
    rows = []
    for row in ws.iter_rows(min_row=21, values_only=True):
        code = str(row[1]).strip() if row[1] is not None else ""
        name_jp = str(row[2]).strip() if row[2] is not None else ""
        if code == "00":
            code, name_jp, name_en = "00", "全国", "Japan (total)"
        elif code in {c for c, _, _ in PREFS}:
            d = {c: (jp, en) for c, jp, en in PREFS}
            name_jp, name_en = d[code]
        else:
            continue  # 大都市圏などの参考地域は本編では扱わない
        total = num(row[5]); vacant = num(row[9])
        other = num(row[10])
        rate = round(vacant / total * 100, 2) if total and vacant is not None else None
        orate = round(other / total * 100, 2) if total and other is not None else None
        rows.append([
            code, name_jp, name_en, 2023,
            int(total) if total is not None else "",
            int(num(row[6])) if num(row[6]) is not None else "",
            int(vacant) if vacant is not None else "",
            int(other) if other is not None else "",
            int(num(row[11])) if num(row[11]) is not None else "",
            int(num(row[12])) if num(row[12]) is not None else "",
            int(num(row[13])) if num(row[13]) is not None else "",
            int(num(row[14])) if num(row[14]) is not None else "",
            int(num(row[15])) if num(row[15]) is not None else "",
            rate if rate is not None else "",
            orate if orate is not None else "",
            SRC_HLS + " 第1-2表",
        ])
    wb.close()
    return write_csv("vacancy_snapshot_pref_2023.csv",
                     ["pref_code", "pref_jp", "pref_en", "year", "total_dwellings",
                      "occupied_dwellings", "vacant_total", "vacant_other",
                      "vacant_for_rent", "vacant_for_sale", "vacant_second",
                      "vacant_villa", "under_construction",
                      "vacancy_rate_pct_calc", "vacant_other_rate_pct_calc", "source"],
                     rows)


# ----------------------------------------------------------------------
# 2. 全国推移: 住宅数・世帯数 (第3表) + 空き家数・種類別 (第4表 総数)
#    第3表: 地域ごとに2行 (住宅数行, 世帯数行)。列5-18=住宅数 1958..2023, 列19-32=世帯数
#    第4表総数: 地域の先頭行のみ採用。列5-18=空き家総数 1958..2023,
#      列19-27=除く空き家 1983..2023, 列28-36=賃貸用 1983..2023,
#      列37-41=売却用 2003..2023, 列42-51=二次的住宅 1978..2023
# ----------------------------------------------------------------------
YEARS_T3 = [1958, 1963, 1968, 1973, 1978, 1983, 1988, 1993, 1998, 2003, 2008, 2013, 2018, 2023]


def build_trend_national():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t3_dwellings_households_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    dwell = {}
    house = {}
    nat_rows = []
    started = False
    for row in ws.iter_rows(min_row=13, values_only=True):
        name = str(row[1]).strip() if row[1] is not None else ""
        if name == "全国":
            nat_rows.append(row)
            started = True
        elif started:
            break
    wb.close()
    # 全国は2行構成: 1行目が現行系列 (1973-2023), 2行目 (注記1つき) に 1958-1973 の旧値。
    # 重なる1973年は現行系列 (1行目) を正とし、1958-1968 のみ2行目から補う。
    for y, v in zip(YEARS_T3, nat_rows[0][5:19]):
        if num(v) is not None:
            dwell[y] = int(num(v))
    for y, v in zip(YEARS_T3, nat_rows[0][19:33]):
        if num(v) is not None:
            house[y] = int(num(v))
    if len(nat_rows) > 1:
        for y, v in zip(YEARS_T3, nat_rows[1][5:19]):
            if y not in dwell and num(v) is not None:
                dwell[y] = int(num(v))
        for y, v in zip(YEARS_T3, nat_rows[1][19:33]):
            if y not in house and num(v) is not None:
                house[y] = int(num(v))

    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t4_vacancy_by_type_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb["総数"]
    nat = None
    for row in ws.iter_rows(min_row=16, values_only=True):
        if str(row[1]).strip() == "00":
            nat = row
            break
    wb.close()
    vac = {}
    series = [
        ("total", list(range(5, 19)), YEARS_T3),
        ("other", list(range(19, 28)), [1983, 1988, 1993, 1998, 2003, 2008, 2013, 2018, 2023]),
        ("rent", list(range(28, 37)), [1983, 1988, 1993, 1998, 2003, 2008, 2013, 2018, 2023]),
        ("sale", list(range(37, 42)), [2003, 2008, 2013, 2018, 2023]),
        ("second", list(range(42, 52)), [1978, 1983, 1988, 1993, 1998, 2003, 2008, 2013, 2018, 2023]),
    ]
    for key, cols, years in series:
        vac[key] = {}
        for c, y in zip(cols, years):
            v = num(nat[c])
            if v is not None:
                vac[key][y] = int(v)
    rows = []
    for y in YEARS_T3:
        vt = vac["total"].get(y)
        td = dwell.get(y)
        rows.append([
            y,
            td if td is not None else "",
            house.get(y, ""),
            vt if vt is not None else "",
            vac["other"].get(y, ""), vac["rent"].get(y, ""),
            vac["sale"].get(y, ""), vac["second"].get(y, ""),
            round(vt / td * 100, 2) if vt and td else "",
            SRC_HLS + " 第3表・第4表",
        ])
    return write_csv("vacancy_trend_national.csv",
                     ["year", "total_dwellings", "households", "vacant_total",
                      "vacant_other", "vacant_for_rent", "vacant_for_sale",
                      "vacant_second", "vacancy_rate_pct_calc", "source"], rows)


# ----------------------------------------------------------------------
# 3. 都道府県別 住宅数・世帯数 (第3表, 2003-2023)
# ----------------------------------------------------------------------
def build_dwellings_pref():
    # 第3表は 1地域1行 (同じ行に住宅数 cols5-18 と世帯数 cols19-32)。
    # 全国のみ2行目に旧系列があるため、全国は現行行を採用し不足年を補完する。
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t3_dwellings_households_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    years_sel = [2003, 2008, 2013, 2018, 2023]
    idx = {y: YEARS_T3.index(y) for y in years_sel}
    data = {}
    nat_extra = None
    for row in ws.iter_rows(min_row=13, values_only=True):
        code = str(row[1]).strip() if row[1] is not None else ""
        name = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        if code == "全国":  # 全国行はコード列ではなく地域名が列1に入る
            code = "00"
        if code == "00":
            if "00" not in data:
                key = ("00", "全国", "Japan (total)")
                data["00"] = (key, row)
            else:
                nat_extra = row
            continue
        if code in {c for c, _, _ in PREFS}:
            d = {c: (jp, en) for c, jp, en in PREFS}
            data[code] = ((code, d[code][0], d[code][1]), row)
    rows = []
    for code in ["00"] + [c for c, _, _ in PREFS]:
        if code not in data:
            continue
        key, row = data[code]
        for y in years_sel:
            dv = num(row[5 + idx[y]])
            hv = num(row[19 + idx[y]])
            if code == "00" and nat_extra is not None:
                if dv is None:
                    dv = num(nat_extra[5 + idx[y]])
                if hv is None:
                    hv = num(nat_extra[19 + idx[y]])
            rows.append([key[0], key[1], key[2], y,
                         int(dv) if dv is not None else "",
                         int(hv) if hv is not None else "",
                         SRC_HLS + " 第3表"])
    wb.close()
    return write_csv("dwellings_households_pref.csv",
                     ["pref_code", "pref_jp", "pref_en", "year",
                      "total_dwellings", "households", "source"], rows)


# ----------------------------------------------------------------------
# 4. 建て方別 空き家数 2023 都道府県別 (第4表 建て方別シート)
#    各シート: 列5-10=総数 1998..2023, 列11-16=除く, 列17-22=賃貸用,
#              列23-27=売却用(2003..), 列28-33=二次的
#    ここでは 2023年の「総数」を採用 (列10)
# ----------------------------------------------------------------------
def build_structure_pref():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t4_vacancy_by_type_pref.xlsx"),
        read_only=True, data_only=True)
    rows = []
    for sheet, label in [("総数", "総数"), ("一戸建", "一戸建"), ("長屋建", "長屋建"),
                         ("共同住宅", "共同住宅"), ("その他", "その他")]:
        ws = wb[sheet]
        col = 18 if sheet == "総数" else 10  # 総数シートは年次列が 5-18 (2023=18), 建て方別は 5-10 (2023=10)
        for row in ws.iter_rows(min_row=16, values_only=True):
            code = str(row[1]).strip() if row[1] is not None else ""
            if code == "00":
                key = ("00", "全国", "Japan (total)")
            elif code in {c for c, _, _ in PREFS}:
                d = {c: (jp, en) for c, jp, en in PREFS}
                key = (code, d[code][0], d[code][1])
            else:
                continue
            v = num(row[col])
            if v is None:
                continue
            rows.append([key[0], key[1], key[2], 2023, label, int(v),
                         SRC_HLS + " 第4表"])
    wb.close()
    return write_csv("vacancy_by_structure_pref_2023.csv",
                     ["pref_code", "pref_jp", "pref_en", "year",
                      "structure", "vacant_dwellings", "source"], rows)


# ----------------------------------------------------------------------
# 5. 全国 空き家構成 2023 (第5表): 種類 x 建て方 x 構造 (実数)
#    行は種類ごとにブロック。行ラベルは列1付近。実数は列4-14 (総数/一戸建(総,木,非)/長屋/共同/その他)
# ----------------------------------------------------------------------
def build_composition():
    # 第5表は「年ラベル行 + 種類行 (総数/除く/賃貸用/売却用/二次的)」のブロック形式。
    # 列: 4=総数 5=一戸建(総数) 8=長屋建(総数) 11=共同住宅(総数) 14=その他(総数)
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t5_vacancy_composition_national.xlsx"),
        read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    in_2023 = False
    label_map = [
        ("空き家総数", "総数"),
        ("賃貸・売却用及び二次的住宅を除く空き家", "賃貸・売却用及び二次的住宅を除く空き家"),
        ("賃貸用の空き家", "賃貸用の空き家"),
        ("売却用の空き家", "売却用の空き家"),
        ("二次的住宅", "二次的住宅"),
    ]
    for row in ws.iter_rows(min_row=10, values_only=True):
        label = str(row[1]).replace(" ", "").replace("　", "").strip() if len(row) > 1 and row[1] is not None else ""
        if label.startswith("令和５年") or label == "2023":
            in_2023 = True
            continue
        if not in_2023:
            continue
        matched = None
        for pat, canonical in label_map:
            if label.startswith(pat.replace(" ", "")):
                matched = canonical
                break
        if matched:
            total = num(row[4])
            if total is not None:
                rows.append([2023, matched, int(total),
                             int(num(row[5])) if num(row[5]) is not None else "",
                             int(num(row[8])) if num(row[8]) is not None else "",
                             int(num(row[11])) if num(row[11]) is not None else "",
                             int(num(row[14])) if num(row[14]) is not None else "",
                             SRC_HLS + " 第5表"])
        if len(rows) >= 5:
            break
    wb.close()
    return write_csv("vacancy_composition_national.csv",
                     ["year", "vacancy_type", "total", "detached",
                      "tenement", "apartment", "other_structure", "source"], rows)


# ----------------------------------------------------------------------
# 6. 世帯スナップショット 2023 都道府県別 (第2-2表 令和5年シート)
#    列: 5=世帯数総数(住宅に居住) 6=主世帯 7=同居世帯 9=世帯人員総数 10=主世帯人員
# ----------------------------------------------------------------------
def build_households_snapshot():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "estat_hls_t2-2_households_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb["令和5年"]
    rows = []
    for row in ws.iter_rows(min_row=18, values_only=True):
        code = str(row[1]).strip() if row[1] is not None else ""
        if code == "00":
            key = ("00", "全国", "Japan (total)")
        elif code in {c for c, _, _ in PREFS}:
            d = {c: (jp, en) for c, jp, en in PREFS}
            key = (code, d[code][0], d[code][1])
        else:
            continue
        hh = num(row[5]); members = num(row[9])
        if hh is None:
            continue
        rows.append([key[0], key[1], key[2], 2023, int(hh),
                     int(num(row[6])) if num(row[6]) is not None else "",
                     int(members) if members is not None else "",
                     round(members / hh, 2) if members and hh else "",
                     SRC_HLS + " 第2-2表"])
    wb.close()
    return write_csv("households_snapshot_pref_2023.csv",
                     ["pref_code", "pref_jp", "pref_en", "year", "households_total",
                      "principal_households", "household_members",
                      "persons_per_household_calc", "source"], rows)


# ----------------------------------------------------------------------
# 7. 地価公示 2025 都道府県別・用途別 変動率 (第5表)
#    行7=全国, 以降は都道府県 (空白行区切り)。列: 0=県名 1=住宅地2024 2=住宅地2025
#    3=地点数 4=商業地2024 5=商業地2025 6=地点数 7=工業地2024 8=工業地2025 9=地点数
# ----------------------------------------------------------------------
def build_land_rate_pref():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "mlit_kouji2025_t5_rate_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(min_row=7, values_only=True):
        name = str(row[0]).replace("　", "").replace(" ", "").strip() if row[0] else ""
        if not name:
            continue
        if name == "全国":
            key = ("00", "全国", "Japan (total)")
        elif name in PREF_BY_SHORT:
            key = PREF_BY_SHORT[name]
        else:
            continue
        for use, c24, c25, cpt in [("住宅地", 1, 2, 3), ("商業地", 4, 5, 6), ("工業地", 7, 8, 9)]:
            r25 = num(row[c25])
            if r25 is None and num(row[cpt]) is None:
                continue
            rows.append([key[0], key[1], key[2], use,
                         num(row[c24]) if num(row[c24]) is not None else "",
                         r25 if r25 is not None else "",
                         int(num(row[cpt])) if num(row[cpt]) is not None else "",
                         SRC_KOUJI + " 第5表"])
    wb.close()
    return write_csv("land_price_rate_pref_2025.csv",
                     ["pref_code", "pref_jp", "pref_en", "land_use",
                      "rate_2024_pct", "rate_2025_pct", "survey_points_2025",
                      "source"], rows)


# ----------------------------------------------------------------------
# 8. 地価公示 時系列 (全国 + 都道府県): 平均価格・変動率
#    変動率推移表/価格推移表: 列0=地域コード 列1=地域名 列2=用途略称 列3-53=1975..2025
#    用途略称: 住宅/宅見/商業/準工/工業/調区/全用途 → 住宅地として「住宅」を採用
# ----------------------------------------------------------------------
USE_MAP = {"住宅": "住宅地", "商業": "商業地", "工業": "工業地", "全用途": "全用途"}


def _load_timeseries(sheet):
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "mlit_kouji_timeseries_rate_price.xlsx"),
        read_only=True, data_only=True)
    ws = wb[sheet]
    data = {}
    current_name = None  # 地域名はブロック先頭行 (住宅) にのみ記載されるため引き継ぐ
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = str(row[0]).strip() if row[0] is not None else ""
        use = str(row[2]).strip() if row[2] is not None else ""
        if row[1] is not None and str(row[1]).strip():
            current_name = str(row[1]).strip()
        if use not in USE_MAP:
            continue
        if code.startswith("0199"):
            key = ("00", "全国")
        elif code.startswith("04") and len(code) == 9:
            name = current_name.replace("県", "").replace("府", "") if current_name else ""
            if name in PREF_BY_SHORT:
                k = PREF_BY_SHORT[name]
                key = (k[0], k[1])
            else:
                continue
        else:
            continue
        vals = {}
        for i, y in enumerate(range(1975, 2026)):
            v = num(row[3 + i])
            if v is not None:
                vals[y] = v
        data[(key[0], USE_MAP[use])] = (key[1], vals)
    wb.close()
    return data


def build_land_timeseries():
    rates = _load_timeseries("変動率推移表")
    prices = _load_timeseries("価格推移表")
    # 全国推移
    rows = []
    for (code, use), (name, vals) in sorted(prices.items()):
        if code != "00":
            continue
        rvals = rates.get((code, use), (None, {}))[1]
        for y in sorted(set(vals) | set(rvals)):
            rows.append([y, use,
                         int(vals[y]) if y in vals else "",
                         rvals.get(y, ""),
                         SRC_KOUJI + " 価格・変動率推移表"])
    write_csv("land_price_trend_national.csv",
              ["year", "land_use", "avg_price_yen_m2", "rate_pct", "source"], rows)
    # 都道府県 2025 スナップショット
    rows = []
    en_map = {c: en for c, _, en in PREFS}
    for (code, use), (name, vals) in sorted(prices.items()):
        if code == "00":
            jp, en = "全国", "Japan (total)"
        else:
            jp = name
            en = en_map.get(code, "")
        rvals = rates.get((code, use), (None, {}))[1]
        if 2025 in vals or 2025 in rvals:
            rows.append([code, jp, en, use,
                         int(vals[2025]) if 2025 in vals else "",
                         rvals.get(2025, ""),
                         SRC_KOUJI + " 価格・変動率推移表"])
    return write_csv("land_price_snapshot_2025.csv",
                     ["pref_code", "pref_jp", "pref_en", "land_use",
                      "avg_price_2025_yen_m2", "rate_2025_pct", "source"], rows)


# ----------------------------------------------------------------------
# 9. 着工 都道府県別 年計 (e-Stat 時系列表 .xls, 9年ブロック x 9年)
# ----------------------------------------------------------------------
def build_starts_pref():
    wb = xlrd.open_workbook(os.path.join(RAW, "estat_chakko_starts_by_pref_annual.xls"))
    sh = wb.sheet_by_index(0)
    rows = []
    block_starts = [0, 65, 130, 195, 260, 325, 390, 455, 520]
    start_years = [1951, 1960, 1969, 1978, 1987, 1996, 2005, 2014, 2023]
    for bs, sy in zip(block_starts, start_years):
        for r in range(bs + 6, min(bs + 65, sh.nrows)):
            name = str(sh.cell_value(r, 0)).replace(" ", "").replace("　", "").strip()
            if not name:
                continue
            if name == "合計":
                key = ("00", "全国", "Japan (total)")
            elif name in PREF_BY_SHORT:
                key = PREF_BY_SHORT[name]
            else:
                continue
            for k in range(9):
                year = sy + k
                if year < 2014:
                    continue
                v = sh.cell_value(r, 1 + 2 * k)
                if isinstance(v, (int, float)) and v:
                    rows.append([key[0], key[1], key[2], year, int(v),
                                 SRC_CHAKKO + " 都道府県別戸数 時系列(年計)"])
    return write_csv("starts_pref_annual.csv",
                     ["pref_code", "pref_jp", "pref_en", "year",
                      "new_housing_starts", "source"], rows)


# ----------------------------------------------------------------------
# 10. 着工 全国 (国交省 時系列表 xlsx jyuu): 年度系列 + 利用関係別 + 暦年
#     列: 1=総計戸数 5=持家 9=貸家 11=給与住宅 13=分譲住宅 15=うちマンション 18=うち一戸建
# ----------------------------------------------------------------------
def parse_wareki_year(label):
    s = str(label).replace(" ", "").replace("　", "")
    m = re.match(r"(平成|令和|昭和)(元|\d+)年度", s)
    if m:
        era = m.group(1)
        n = 1 if m.group(2) == "元" else int(m.group(2))
        base = {"平成": 1988, "令和": 2018, "昭和": 1925}[era]
        return base + n
    return None


def build_starts_national():
    wb = openpyxl.load_workbook(
        os.path.join(RAW, "mlit_chakko_starts_by_pref.xlsx"),
        read_only=True, data_only=True)
    ws = wb["jyuu"]
    rows = []
    monthly = {}
    for row in ws.iter_rows(min_row=8, values_only=True):
        label = str(row[0]) if row[0] is not None else ""
        fy = parse_wareki_year(label)
        if fy and num(row[1]) is not None:
            rows.append([fy, "年度",
                         int(num(row[1])),
                         int(num(row[5])) if num(row[5]) is not None else "",
                         int(num(row[9])) if num(row[9]) is not None else "",
                         int(num(row[11])) if num(row[11]) is not None else "",
                         int(num(row[13])) if num(row[13]) is not None else "",
                         int(num(row[15])) if num(row[15]) is not None else "",
                         int(num(row[18])) if num(row[18]) is not None else "",
                         "確定" if fy <= 2024 else "年度途中(未確定)",
                         SRC_CHAKKO + " 時系列表(全国)"])
        # 月次: ラベルが col0 に年, col1 は月ではなく戸数。月次行は col0 に "R7年 1月" 形式
        m = re.match(r"\s*([SHR])(\d+)年\s*(\d+)月", label)
        if m:
            era = {"S": 1925, "H": 1988, "R": 2018}[m.group(1)]
            y = era + int(m.group(2))
            mo = int(m.group(3))
            if num(row[1]) is not None:
                monthly[(y, mo)] = int(num(row[1]))
    wb.close()
    # 暦年集計 (12か月そろった年のみ)
    for y in sorted({y for y, _ in monthly}):
        vals = [monthly.get((y, m)) for m in range(1, 13)]
        if all(v is not None for v in vals):
            rows.append([y, "暦年", sum(vals), "", "", "", "", "", "",
                         "月次12か月の自前集計",
                         SRC_CHAKKO + " 時系列表(全国・月次)より集計"])
    rows.sort(key=lambda r: (r[0], r[1]))
    return write_csv("starts_national_annual.csv",
                     ["year", "year_type", "total_starts", "owner_occupied",
                      "for_rent", "company_housing", "for_sale", "of_which_condo",
                      "of_which_detached_fs", "status", "source"], rows)


if __name__ == "__main__":
    build_snapshot()
    build_trend_national()
    build_dwellings_pref()
    build_structure_pref()
    build_composition()
    build_households_snapshot()
    build_land_rate_pref()
    build_land_timeseries()
    build_starts_pref()
    build_starts_national()
    print("done")
