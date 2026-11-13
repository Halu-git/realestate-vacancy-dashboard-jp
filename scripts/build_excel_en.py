#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.6 "Real Estate & Vacancy Dashboard" — English edition builder.

Loads the Japanese edition (output/realestate_vacancy_dashboard.xlsx, built by
build_excel.py) and translates every label in place — sheet names, headings,
KPI labels, table headers, chart titles, source lines and notes — then saves
output/realestate_vacancy_dashboard_en.xlsx. Data values are never touched,
so the English edition cannot drift numerically from the Japanese one.
Run build_excel.py first, then this script.
"""
import pathlib

from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "output" / "realestate_vacancy_dashboard.xlsx"
OUTF = ROOT / "output" / "realestate_vacancy_dashboard_en.xlsx"

SHEET_MAP = {
    "表紙・使い方": "Cover & Guide",
    "D1_KPI概要": "D1_KPI Overview",
    "D2_都道府県別_空き家": "D2_Vacancy by Prefecture",
    "D3_地価・着工": "D3_Land Prices & Starts",
    "D4_推移": "D4_Trends",
    "data_空き家_都道府県": "data_Vacancy_by_Pref",
    "data_推移_全国": "data_National_Trend",
    "data_地価": "data_Land_Prices",
    "data_着工": "data_Housing_Starts",
    "data_世帯": "data_Households",
    "検証ログ": "Validation Log",
    "出典・注記": "Sources & Notes",
}
PREF_EN = {
    "北海道": "Hokkaido", "青森県": "Aomori", "岩手県": "Iwate", "宮城県": "Miyagi",
    "秋田県": "Akita", "山形県": "Yamagata", "福島県": "Fukushima", "茨城県": "Ibaraki",
    "栃木県": "Tochigi", "群馬県": "Gunma", "埼玉県": "Saitama", "千葉県": "Chiba",
    "東京都": "Tokyo", "神奈川県": "Kanagawa", "新潟県": "Niigata", "富山県": "Toyama",
    "石川県": "Ishikawa", "福井県": "Fukui", "山梨県": "Yamanashi", "長野県": "Nagano",
    "岐阜県": "Gifu", "静岡県": "Shizuoka", "愛知県": "Aichi", "三重県": "Mie",
    "滋賀県": "Shiga", "京都府": "Kyoto", "大阪府": "Osaka", "兵庫県": "Hyogo",
    "奈良県": "Nara", "和歌山県": "Wakayama", "鳥取県": "Tottori", "島根県": "Shimane",
    "岡山県": "Okayama", "広島県": "Hiroshima", "山口県": "Yamaguchi", "徳島県": "Tokushima",
    "香川県": "Kagawa", "愛媛県": "Ehime", "高知県": "Kochi", "福岡県": "Fukuoka",
    "佐賀県": "Saga", "長崎県": "Nagasaki", "熊本県": "Kumamoto", "大分県": "Oita",
    "宮崎県": "Miyazaki", "鹿児島県": "Kagoshima", "沖縄県": "Okinawa",
    "全国": "Japan (total)",
}
TERM_EN = {
    "住宅地": "Residential", "商業地": "Commercial", "工業地": "Industrial",
    "全用途": "All uses", "総数": "Total", "総計": "Total",
    "一戸建": "Detached houses", "長屋建": "Tenement-houses",
    "共同住宅": "Apartments", "その他": "Other",
    "賃貸用": "For rent", "売却用": "For sale", "二次的住宅": "Second dwellings",
    "賃貸・売却用・二次的住宅を除く空き家": "Other vacant dwellings",
    "賃貸・売却用及び二次的住宅を除く空き家": "Other vacant dwellings",
    "賃貸用の空き家": "Vacant dwellings for rent",
    "売却用の空き家": "Vacant dwellings for sale",
    "持家": "Owner-occupied", "貸家": "For rent", "分譲": "For sale",
    "分譲住宅": "For-sale housing", "年度": "Fiscal year", "暦年": "Calendar year",
    "都道府県別・暦年": "By prefecture (calendar year)",
    "全国推移": "National trend",
    "変動率（第5表）": "Change rate (Table 5)",
    "平均価格・変動率2025（推移表）": "Avg. price & rate 2025 (time-series table)",
}
STRING_MAP = {}
STRING_MAP.update({
# ---- cover
"不動産・空き家ダッシュボード（都道府県レベル）": "Real Estate & Vacancy Dashboard (Prefecture Level)",
"都道府県レベルで「空き家ストック × 地価 × 住宅着工フロー」を見る市場ダッシュボードのサンプルです。": "A sample market dashboard viewing 'vacant-dwelling stock x land prices x housing-start flows' at prefecture level.",
"単一エリアの出店診断（サンプル第1弾）とは異なり、都道府県単位の広域市場概観に特化しています。": "Unlike the single-area site-selection diagnosis (Sample No.1), this dashboard focuses on a wide-area market overview at prefecture level.",
"対象読者： 不動産事業者、自治体の空き家対策担当、住宅関連事業者": "Intended readers: real estate businesses, municipal vacant-house policy officers, housing-related businesses",
"": "[IMPORTANT: Freshness limitation]",
"空き家の主データである住宅・土地統計調査は5年周期の調査で、最新は2023年（令和5年）10月1日現在です。": "The main data source for vacant dwellings, the Housing and Land Survey, is conducted every five years; the latest is as of 1 October 2023.",
"2024年以降の空き家動向はこの統計では追えません（次回調査は2028年予定）。直近の市場動向は、": "Vacancy trends from 2024 onward cannot be tracked with this survey (the next survey is planned for 2028). Recent market movements are",
"毎年更新される地価公示（2025年）と建築着工統計（2025年）で補完する設計にしています。": "covered instead by the annually updated Official Land Price survey (2025) and Building Starts statistics (2025).",
"シート構成": "Sheet guide",
"D1_KPI概要 ― 空き家・地価・着工の主要指標（最新値）": "D1_KPI Overview - Key indicators for vacancy, land prices and starts (latest values)",
"D2_都道府県別_空き家 ― 47都道府県の空き家数・空き家率・建て方別構成": "D2_Vacancy by Prefecture - Vacant dwellings, vacancy rates and structure mix for all 47 prefectures",
"D3_地価・着工 ― 都道府県別の地価変動率と住宅着工戸数": "D3_Land Prices & Starts - Land-price change rates and housing starts by prefecture",
"D4_推移 ― 住宅ストックと世帯の長期推移、空き家の種類別推移、地価・着工の推移": "D4_Trends - Long-term trends in housing stock and households, vacancy by type, land prices and starts",
"data_* ― 整形済みデータ（出典列つき）／検証ログ ― 公表値との突合結果／出典・注記 ― 定義・採用理由・不採用理由": "data_* - Cleaned data (with source columns) / Validation Log - Cross-checks against published figures / Sources & Notes - Definitions and reasons for including/excluding data",
"使い方： D1→D2→D3→D4 の順に読むと「現状（ストック）→ 地域差 → 直近の動き（価格・フロー）→ 長期トレンド」が追えます。": "How to read: follow D1 -> D2 -> D3 -> D4 to move from the current stock, to regional differences, to recent price/flow movements, to long-term trends.",
"数値の出典は各シート下部と「出典・注記」に明記しています。率は原則として本ダッシュボードの計算値です。": "Sources are stated at the bottom of each sheet and in 'Sources & Notes'. Rates are calculated values in this dashboard unless noted otherwise.",
# ---- D1
"D1 KPI概要 ― 空き家ストック、地価、着工フロー（最新値）": "D1 KPI Overview - Vacancy stock, land prices and housing-start flows (latest)",
"空き家数（2023年）": "Vacant dwellings (2023)",
"過去最多（2018年 8,488,600戸から増加）": "Record high (up from 8,488,600 in 2018)",
"空き家率（2023年）": "Vacancy rate (2023)",
"％（公表値13.8％と一致・計算値）": "% (calculated; matches the published 13.8%)",
"「その他」空き家（2023年）": "'Other' vacant dwellings (2023)",
"賃貸・売却用・二次的住宅を除く空き家": "Vacant dwellings excluding for rent / for sale / second dwellings",
"総住宅数（2023年）": "Total dwellings (2023)",
"総世帯数 56,215,300世帯": "Total households: 56,215,300",
"地価公示 住宅地 変動率（2025年）": "Official Land Prices: residential change rate (2025)",
"％（全国平均・対前年）": "% (national average, YoY)",
"新設住宅着工戸数（2025年・暦年）": "New housing starts (2025, calendar year)",
"前年比 -6.5％": "YoY -6.5%",
"空き家率が最も高いのは徳島県（21.3％）、最も低いのは埼玉県（9.3％）。地価公示2025は全用途 +2.7％、商業地 +3.9％と上昇が続く。住宅地の平均価格は 137,100円/m²。": "Tokushima has the highest vacancy rate (21.3%) and Saitama the lowest (9.3%). Official Land Prices 2025 rose +2.7% for all uses and +3.9% for commercial land. The average residential land price is JPY 137,100/m2.",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報）／国土交通省 地価公示（令和7年）／国土交通省 建築着工統計調査（住宅着工統計）": "Sources: Statistics Bureau, Housing and Land Survey (2023, final figures) / MLIT, Official Land Prices (2025) / MLIT, Building Starts Survey (Housing Starts Statistics)",
"空き家数と空き家率の推移（全国、1978〜2023年）": "Vacant dwellings and vacancy rate (Japan, 1978-2023)",
"年": "Year",
"空き家数（戸）": "Vacant dwellings",
"空き家率（％・計算値）": "Vacancy rate (%, calculated)",
"空き家数（棒）と空き家率（線）の推移": "Vacant dwellings (bars) and vacancy rate (line)",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第3表・第4表。空き家率は空き家数÷総住宅数の計算値。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Tables 3 and 4. The vacancy rate is calculated as vacant dwellings / total dwellings.",
"空き家の種類別内訳（全国、2023年）": "Vacant dwellings by type (Japan, 2023)",
"種類": "Type",
"戸数": "Dwellings",
"構成比（％）": "Share (%)",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第5表。「その他」は表記上「賃貸・売却用及び二次的住宅を除く空き家」。内訳合計と総数は原表の丸めで最大100戸差が生じる。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Table 5. 'Other' denotes vacant dwellings excluding those for rent, for sale and as second dwellings. Component sums may differ from the total by up to 100 dwellings due to rounding in the source tables.",
# ---- D2
"D2 都道府県別 空き家 ― 空き家数・空き家率・建て方別（2023年）": "D2 Vacancy by Prefecture - Counts, rates and structure mix (2023)",
"都道府県": "Prefecture",
"総住宅数": "Total dwellings",
"世帯数": "Households",
"空き家数": "Vacant dwellings",
"空き家率（％）": "Vacancy rate (%)",
"その他空き家": "Other vacant dwellings",
"その他率（％）": "Other share (%)",
"賃貸用空き家": "Vacant for rent",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第1-2表・第2-2表。空き家率・その他率は計算値。表は空き家率の降順。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Tables 1-2 and 2-2. Vacancy and 'other' rates are calculated values. Sorted by vacancy rate, descending.",
"空き家率 上位10県・下位10県（％）": "Vacancy rate: top 10 and bottom 10 prefectures (%)",
"空き家率 上位10県・下位10県（2023年）": "Vacancy rate: top 10 / bottom 10 prefectures (2023)",
"建て方別 空き家数（全国、2023年）": "Vacant dwellings by structure type (Japan, 2023)",
"建て方": "Structure type",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第4表。空き家の過半は共同住宅だが、「その他」空き家は一戸建に偏る（D1の種類別内訳・data参照）。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Table 4. Apartments account for over half of all vacant dwellings, while 'other' vacant dwellings are concentrated in detached houses (see the D1 breakdown and data sheets).",
# ---- D3
"D3 地価・着工 ― 都道府県別の地価変動率（2025年）と住宅着工戸数": "D3 Land Prices & Starts - Land-price change rates (2025) and housing starts by prefecture",
"住宅地 変動率2025（％）": "Residential rate 2025 (%)",
"商業地 変動率2025（％）": "Commercial rate 2025 (%)",
"住宅地 平均価格2025（円/m²）": "Residential avg. price 2025 (JPY/m2)",
"出典： 国土交通省 地価公示（令和7年） 第5表・価格/変動率推移表。平均価格は標準地の平均で、取引価格ではない。表は住宅地変動率の降順。": "Source: MLIT, Official Land Prices (2025), Table 5 and price/rate time-series tables. Average prices are averages of standard land plots, not transaction prices. Sorted by residential change rate, descending.",
"住宅地 変動率 上位15県（2025年、％）": "Residential change rate: top 15 prefectures (2025, %)",
"住宅地変動率（％）": "Residential change rate (%)",
"地価公示 住宅地変動率 上位15県（2025年）": "Official Land Prices: top 15 prefectures by residential change rate (2025)",
"新設住宅着工戸数（全国、暦年 2014〜2025年）": "New housing starts (Japan, calendar years 2014-2025)",
"着工戸数": "Housing starts",
"前年比（％）": "YoY (%)",
"新設住宅着工戸数の推移（全国・暦年）": "New housing starts (Japan, calendar year)",
"出典： 国土交通省 建築着工統計調査（住宅着工統計） 都道府県別戸数 時系列（年計）。2025年は建築基準法改正（4月）を控えた駆け込みと反動を含む。": "Source: MLIT, Building Starts Survey, starts by prefecture (annual totals). CY2025 includes a rush ahead of the April 2025 Building Standards Act revision and the subsequent pullback.",
# ---- D4
"D4 推移 ― ストック・価格・フローの長期推移": "D4 Trends - Long-term trends in stock, prices and flows",
"住宅数と世帯数（全国、1958〜2023年）": "Dwellings and households (Japan, 1958-2023)",
"総世帯数": "Total households",
"総住宅数と総世帯数の推移": "Total dwellings vs total households",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第3表。住宅数が世帯数を上回り続けており、その差が空き家の受け皿になっている。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Table 3. Dwellings have consistently outnumbered households, and the gap is what vacancy fills.",
"空き家の種類別推移（全国、1983〜2023年）": "Vacant dwellings by type (Japan, 1983-2023)",
"その他（賃貸・売却・二次を除く）": "Other (excl. rent/sale/second)",
"空き家の種類別推移（積み上げ）": "Vacant dwellings by type (stacked)",
"出典： 総務省統計局 住宅・土地統計調査（令和5年・基本集計 確報） 第4表。売却用は2003年から分離集計（それ以前は賃貸用に含まれる）。「その他」の増加が空き家問題の中核。": "Source: Statistics Bureau, Housing and Land Survey (2023, final), Table 4. 'For sale' is tabulated separately only from 2003 (included in 'for rent' before that). Growth in the 'other' category is the core of Japan's vacancy problem.",
"地価公示 平均価格の推移（全国、1975〜2025年）": "Official Land Prices: average price trend (Japan, 1975-2025)",
"住宅地 平均価格（円/m²）": "Residential avg. price (JPY/m2)",
"商業地 平均価格（円/m²）": "Commercial avg. price (JPY/m2)",
"出典： 国土交通省 地価公示（令和7年） 価格推移表（5年刻み＋2025年）。完全な年次系列は data_地価 を参照。": "Source: MLIT, Official Land Prices (2025), price time-series table (every 5 years plus 2025). See data_Land_Prices for the full annual series.",
"新設住宅着工戸数（全国、年度 1994〜2024年度）": "New housing starts (Japan, fiscal years 1994-2024)",
"新設住宅着工戸数（年度・利用関係別は表参照）": "New housing starts (fiscal year; see table for tenure breakdown)",
"出典： 国土交通省 建築着工統計調査（住宅着工統計） 時系列表（全国）。2025年度は年度途中のため未掲載（値は data_着工 に未確定として収録）。": "Source: MLIT, Building Starts Survey, time-series table (national). FY2025 is not shown because the fiscal year is still in progress (recorded as provisional in data_Housing_Starts).",
})
STRING_MAP.update({
    "【鮮度の限界について（重要）】": "[IMPORTANT: Freshness limitation]",
    "・【鮮度限界】住調は5年周期で最新は2023年10月1日現在。2024年以降の空き家動向は追えない（次回2028年調査予定）。地価・着工は年次で補完しているが、空き家そのものの代替指標ではない。": "- [Freshness limitation] The Housing and Land Survey is five-yearly; the latest is as of 1 October 2023. Vacancy trends from 2024 onward cannot be tracked (next survey planned for 2028). Land prices and starts are annual complements, not substitutes for vacancy itself.",
    # ---- data sheets
"データ": "Dataset",
"値1": "Value 1",
"値2": "Value 2",
"出典": "Source",
"用途": "Land use",
"都道府県/年種別": "Prefecture / year type",
"pref_code/年": "pref_code / year",
"戸数/総計": "Units / total",
"世帯総数（第2-2表・2023）": "Households total (Table 2-2, 2023)",
"1世帯当たり人員（2023）": "Persons per household (2023)",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第1-2表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Table 1-2",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第3表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Table 3",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第3表・第4表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Tables 3 and 4",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第2-2表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Table 2-2",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第4表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Table 4",
"総務省統計局 住宅・土地統計調査 令和5年(2023) 基本集計 第5表": "Statistics Bureau, Housing and Land Survey 2023, Basic Tabulation, Table 5",
"国土交通省 地価公示 令和7年(2025) 第5表": "MLIT, Official Land Prices 2025, Table 5",
"国土交通省 地価公示 令和7年(2025) 価格・変動率推移表": "MLIT, Official Land Prices 2025, price/rate time-series tables",
"国土交通省 建築着工統計調査(住宅着工統計) 都道府県別戸数 時系列(年計)": "MLIT, Building Starts Survey (Housing Starts Statistics), starts by prefecture, annual time series",
"国土交通省 建築着工統計調査(住宅着工統計) 時系列表(全国)（確定）": "MLIT, Building Starts Survey, time-series table (national) (final)",
"国土交通省 建築着工統計調査(住宅着工統計) 時系列表(全国)（年度途中(未確定)）": "MLIT, Building Starts Survey, time-series table (national) (fiscal year in progress, provisional)",
"国土交通省 建築着工統計調査(住宅着工統計) 時系列表(全国・月次)より集計（月次12か月の自前集計）": "Aggregated from the MLIT Building Starts Survey monthly time series (national); own aggregation of 12 months",
# ---- validation log
"検証ログ ― 公表値との突合（2026-11-13 実施）": "Validation Log - Cross-checks against published figures (run 2026-11-13)",
"項目": "Item",
"本ダッシュボードの値": "Value in this dashboard",
"突合相手": "Reference",
"結果": "Result",
"備考": "Notes",
"一致": "Match",
"完全一致": "Exact match",
"全件一致": "All matched",
"整合": "Consistent",
"採用": "Used",
"同上": "Same as above",
"空き家総数（2023・全国）": "Total vacant dwellings (2023, Japan)",
"9,001,600戸": "9,001,600 dwellings",
"900万2千戸（基本集計 確報・結果概要）": "9,002 thousand dwellings (Basic Tabulation, final results summary)",
"公表は千戸単位に丸め": "Published figure is rounded to thousands",
"空き家率（2023・全国）": "Vacancy rate (2023, Japan)",
"13.84％（計算値）": "13.84% (calculated)",
"13.8％（同上）": "13.8% (same source)",
"小数第1位まで一致": "Matches to one decimal place",
"賃貸・売却用及び二次的住宅を除く空き家（2023）": "Vacant dwellings excl. rent/sale/second (2023)",
"3,856,000戸": "3,856,000 dwellings",
"385万6千戸（同上）": "3,856 thousand dwellings (same source)",
"総住宅数（2023・全国）": "Total dwellings (2023, Japan)",
"65,046,700戸": "65,046,700 dwellings",
"6,502万戸（同上）": "65.02 million dwellings (same source)",
"空き家総数（2018・全国）": "Total vacant dwellings (2018, Japan)",
"8,488,600戸": "8,488,600 dwellings",
"848万9千戸（同上）": "8,489 thousand dwellings (same source)",
"総世帯数（2023・全国）": "Total households (2023, Japan)",
"56,215,300世帯": "56,215,300 households",
"第3表と第2-2表の二源が完全一致": "Tables 3 and 2-2 agree exactly (two sources)",
"定義： 住宅に居住する世帯の総数": "Definition: households living in dwellings",
"空き家 種類別内訳の合計（2023）": "Sum of vacancy types (2023)",
"9,001,500戸": "9,001,500 dwellings",
"空き家総数 9,001,600戸": "Vacant total: 9,001,600 dwellings",
"差100戸": "Difference: 100",
"原表の丸めによる差（各値は推計・丸め済み）": "Difference due to rounding in the source tables (values are estimates, pre-rounded)",
"建て方別合計 vs 空き家総数（都道府県別）": "Structure-type sum vs vacant total (by prefecture)",
"差は最大100戸": "Maximum difference: 100 dwellings",
"第4表と第1-2表の照合": "Cross-check of Table 4 against Table 1-2",
"原表の丸めの範囲内": "Within source-table rounding",
"地価公示2025 全国変動率": "Official Land Prices 2025: national change rates",
"住宅地2.1／商業地3.9／工業地4.8／全用途2.7％": "Residential 2.1 / Commercial 3.9 / Industrial 4.8 / All uses 2.7 (%)",
"第5表と価格・変動率推移表の二源": "Two sources: Table 5 and the price/rate time-series tables",
"全用途2.7％は公表値とも一致": "All-uses 2.7% also matches the published headline figure",
"地価公示2025 都道府県別変動率": "Official Land Prices 2025: change rates by prefecture",
"141件を二源照合": "141 entries cross-checked between two sources",
"第5表 vs 推移表": "Table 5 vs time-series tables",
"着工 暦年合計（2023）": "Housing starts, calendar-year total (2023)",
"着工 暦年合計（2024）": "Housing starts, calendar-year total (2024)",
"着工 暦年合計（2025）": "Housing starts, calendar-year total (2025)",
"819,623戸": "819,623 units",
"792,195戸": "792,195 units",
"740,667戸": "740,667 units",
"都道府県別年計表 vs 全国月次の自前集計": "Prefecture annual table vs own aggregation of national monthly figures",
"着工 年度計（2024年度）": "Housing starts, fiscal-year total (FY2024)",
"816,388戸": "816,388 units",
"時系列表（全国）": "Time-series table (national)",
"2025年度分 711,171戸は年度途中の未確定値として区別": "The FY2025 figure of 711,171 units is flagged separately as provisional (fiscal year in progress)",
"突合相手の公表値は、総務省「令和5年住宅・土地統計調査 住宅及び世帯に関する基本集計（確報集計）結果」（2024-09-25公表）等の一次資料による。": "Reference figures are taken from primary sources, including the Statistics Bureau's 'Housing and Land Survey 2023: Basic Tabulation (final) results' (published 25 Sep 2024).",
# ---- sources & notes
"出典・注記 ― 定義、採用理由・不採用理由、限界": "Sources & Notes - Definitions, reasons for including/excluding data, limitations",
"採用したデータと採用理由": "Datasets used, and why",
"採用しなかったデータと不採用理由": "Datasets not used, and why",
"定義・利用上の注意": "Definitions and notes on use",
"住宅・土地統計調査（総務省統計局、令和5年・基本集計 確報）": "Housing and Land Survey (Statistics Bureau, 2023, Basic Tabulation final figures)",
"空き家数・空き家率の公式定義を持つ唯一の専門調査。都道府県別・空き家の種類別・建て方別まで一貫した定義で比較できるため、本ダッシュボードの主データとして採用。速報ではなく確報を用いた。": "The only specialised survey with official definitions of vacant dwellings and the vacancy rate. It allows consistent comparison by prefecture, vacancy type and structure type, so it is the main dataset of this dashboard. Final figures are used rather than preliminary ones.",
"地価公示（国土交通省、令和7年）": "Official Land Prices (MLIT, 2025)",
"公的な地価指標で毎年公表される。都道府県別・用途別の平均価格と対前年変動率が揃い、5年周期の住調を年次で補完できるため採用。第5表と推移表の二源で相互検証できることも採用理由。": "An official land-price indicator published every year. Average prices and year-on-year change rates by prefecture and land use complement the five-yearly Housing and Land Survey on an annual cycle. The ability to cross-validate between Table 5 and the time-series tables is another reason for adoption.",
"建築着工統計調査／住宅着工統計（国土交通省）": "Building Starts Survey / Housing Starts Statistics (MLIT)",
"住宅フローの公式統計。都道府県別・暦年の着工戸数（e-Stat時系列表）が長期で揃い、全国の年度・月次系列と相互検証できるため採用。": "The official statistics on housing flows. Calendar-year starts by prefecture (e-Stat time-series tables) are available over a long period and can be cross-validated against the national fiscal-year and monthly series.",
"世帯数（住宅・土地統計調査 第2-2表・第3表）": "Households (Housing and Land Survey, Tables 2-2 and 3)",
"空き家（ストック）の需要側の分母として、住宅総数との対比に必要。住調の中で定義が揃うため、外部の人口統計ではなく同一調査から取得した。": "Needed as the demand-side counterpart to the dwelling stock. Definitions are consistent within the Housing and Land Survey, so the figures are taken from the same survey rather than from external population statistics.",
"住宅・土地統計調査 速報集計（住宅数概数集計）": "Housing and Land Survey, preliminary tabulation (approximate dwelling counts)",
"早期公表用の概数で、確報で値が更新される。確報（基本集計）が公表済みのため不採用。": "Approximate figures released early and superseded by the final tabulation. Not used because the final figures are already published.",
"住宅・土地統計調査 市区町村別結果": "Housing and Land Survey, municipality-level results",
"本サンプルは都道府県粒度に限定したため不採用。市区町村粒度は自治体・事業者の買い手価値が高いため、商品化時は市区町村別レポートとして有料オプションで提供する想定。": "Not used because this sample is limited to prefecture level. Municipality-level data has high buyer value for local governments and businesses, so in commercial work it is planned as a paid optional extra (municipality-level report).",
"建築着工統計 構造別・建て方別・利用関係別の年別詳細ファイル": "Building Starts Survey, detailed annual files by structure, building type and tenure",
"粒度（年×都道府県×構造×建て方×利用関係）が本ダッシュボードの粒度を大きく上回り、主要指標は時系列表で代替できるため本編では不採用。商品化時は構造別・利用関係別の深掘り分析として有料オプション提供の想定。": "Their granularity (year x prefecture x structure x building type x tenure) far exceeds this dashboard's, and the key indicators are covered by the time-series tables, so they are not used in the main product. In commercial work they are planned as a paid optional extra for deeper structure/tenure analysis.",
"地価公示 市区町村別の平均価格・変動率表": "Official Land Prices, municipality-level average price / change-rate tables",
"都道府県粒度に限定したため不採用。商品化時は市区町村別レポートとして有料オプション提供の想定。": "Not used because this sample is limited to prefecture level. In commercial work, planned as a paid optional extra (municipality-level report).",
"地価公示 地点別（標準地）生データ": "Official Land Prices, plot-level (standard land) raw data",
"都道府県集計は公表表（第5表・推移表）で代替できるため不採用。地点レベルの立地分析は商品化時の有料オプション想定。": "Prefecture aggregates are covered by the published tables (Table 5 and the time-series tables), so the raw plot data is not used. Plot-level location analysis is planned as a paid optional extra in commercial work.",
"国勢調査の住宅・世帯項目": "Population Census, housing and household items",
"空き家の種類別内訳は国勢調査では取れず、空き家の定義は住調が専門であるため不採用。": "The Census does not provide the breakdown of vacant dwellings by type, and vacancy definitions are the speciality of the Housing and Land Survey, so it is not used.",
"人口動態（住民基本台帳・人口推計）": "Population dynamics (Basic Resident Register / population estimates)",
"空き家率の公式分母は住宅総数であり、需要側は世帯数で代表できるため本編では不採用。商品化時は人口動態とのクロス分析を有料オプション想定。": "The official denominator of the vacancy rate is the total dwelling stock, and demand is represented by household counts, so population dynamics are not used in the main product. Cross-analysis with population dynamics is planned as a paid optional extra in commercial work.",
"建築着工統計 都道府県別 月次": "Building Starts Survey, monthly starts by prefecture",
"年次に統一した方が住調（5年）・地価（年次）と粒度が揃うため本編では不採用。全国の月次は暦年集計の検証にのみ使用した。": "Annual figures align better with the five-yearly Housing and Land Survey and the annual land prices, so monthly prefecture figures are not used in the main product. National monthly figures were used only to validate the calendar-year aggregation.",
"・「空き家」（住調の定義）： 居住世帯のいない住宅のうち、一時現在者のみの住宅と建築中の住宅を除いたもの。種類は 賃貸用／売却用／二次的住宅／その他（賃貸・売却用及び二次的住宅を除く空き家＝長期不在・取壊し予定など）。": "- 'Vacant dwelling' (Housing and Land Survey definition): a dwelling with no resident household, excluding dwellings with only temporary occupants and dwellings under construction. Types: for rent / for sale / second dwellings / other (long-term absence, scheduled demolition, etc.).",
"・空き家率＝空き家数÷総住宅数（本ダッシュボードの計算値）。公表値13.8％と小数第1位で一致する。": "- Vacancy rate = vacant dwellings / total dwellings (calculated in this dashboard). It matches the published 13.8% to one decimal place.",
"・住調は5年周期で最新は2023年10月1日現在。2024年以降の空き家動向は追えない（次回2028年調査予定）。地価・着工は年次で補完しているが、空き家そのものの代替指標ではない。": "- [Freshness limitation] The Housing and Land Survey is five-yearly; the latest is as of 1 October 2023. Vacancy trends from 2024 onward cannot be tracked (next survey planned for 2028). Land prices and starts are annual complements, not substitutes for vacancy itself.",
"・地価公示の平均価格は標準地の平均価格であり、実際の取引価格とは異なる。変動率は継続地点の対前年変動率の平均。": "- Official Land Price averages are averages over standard land plots and differ from actual transaction prices. Change rates are averages over continuing plots.",
"・着工統計の2025年度（年度計）は年度途中の未確定値。暦年2025年には建築基準法改正（2025年4月）を控えた駆け込み（3月の急増）とその反動が含まれる。": "- The FY2025 starts total is provisional (fiscal year in progress). CY2025 includes a rush ahead of the April 2025 Building Standards Act revision (a March spike) and the subsequent pullback.",
"・住調の数値は標本調査に基づく推計値で、原表の丸めにより内訳の合計が総数と一致しない場合がある（最大100戸程度）。": "- Housing and Land Survey figures are estimates from a sample survey; due to rounding in the source tables, component sums may differ from totals (by up to about 100 dwellings).",
})

LABEL_MAPS = (PREF_EN, TERM_EN)


def translate_text(s):
    if not s:
        return s
    if s in STRING_MAP:
        return STRING_MAP[s]
    for m in LABEL_MAPS:
        if s in m:
            return m[s]
    return s


def _has_jp(s):
    return any(ord(ch) > 0x2000 for ch in s)


def main():
    wb = load_workbook(SRC)
    for old, new in SHEET_MAP.items():
        wb[old].title = new
    leftovers = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    c.value = translate_text(c.value)
                    if _has_jp(c.value):
                        leftovers.add(c.value)
        for chart in getattr(ws, "_charts", []):
            try:
                for p in chart.title.tx.rich.p:
                    for r in p.r:
                        if r.t:
                            r.t = translate_text(r.t)
                            if _has_jp(r.t):
                                leftovers.add(r.t)
            except Exception:
                pass
            try:
                for ax in (chart.x_axis, chart.y_axis):
                    if ax.title is not None:
                        for p in ax.title.tx.rich.p:
                            for r in p.r:
                                if r.t:
                                    r.t = translate_text(r.t)
            except Exception:
                pass
    wb.save(OUTF)
    print("saved:", OUTF)
    if leftovers:
        print("UNTRANSLATED strings remaining:", len(leftovers))
        for s in sorted(leftovers):
            print("  ", repr(s))
    else:
        print("no untranslated strings")


if __name__ == "__main__":
    main()
