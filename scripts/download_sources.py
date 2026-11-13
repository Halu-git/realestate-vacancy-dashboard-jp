#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第6弾「不動産・空き家ダッシュボード」— 公開データ取得スクリプト

すべての出典は公開ファイルのみ（e-Stat APIのアプリケーションIDは不使用）。
取得日: 2026-11-13。

出典と data/raw/ ファイルの対応:
  S1 住宅・土地統計調査 (総務省統計局、令和5年=2023年、基本集計 確報)
     e-Stat「住宅・土地統計調査」統計表一覧から各表の Excel を取得:
       estat_hls_t1-2_dwellings_by_pref.xlsx        第1-2表 居住世帯の有無別住宅数
       estat_hls_t2-1_households_national.xlsx      第2-1表 世帯数・世帯人員 (全国)
       estat_hls_t2-2_households_by_pref.xlsx       第2-2表 世帯数・世帯人員
       estat_hls_t3_dwellings_households_by_pref.xlsx 第3表 住宅数及び世帯数 (時系列)
       estat_hls_t4_vacancy_by_type_pref.xlsx       第4表 空き家の種類・建て方別空き家数
       estat_hls_t5_vacancy_composition_national.xlsx 第5表 空き家の種類・建て方・構造 (全国)
     結果概要 (突合用 PDF):
       kihon_gaiyou.pdf  「住宅及び世帯に関する基本集計 (確報集計) 結果」2024-09-25 公表
       g_kekka.pdf       「住宅数概数集計 (速報集計) 結果」2024-04-30 公表 (速報は不採用)
  S2 地価公示 (国土交通省、令和7年=2025年)
       mlit_kouji2025_t5_rate_by_pref.xlsx   第5表 都道府県別・用途別 対前年平均変動率
       mlit_kouji_timeseries_rate_price.xlsx 価格・変動率推移表 (長期時系列)
       lic_kouji2025_pref_table.pdf         土地情報センター 都道府県市区町村別表
                                            (都道府県集計の裏取り用。市区町村部は不採用)
  S3 建築着工統計調査 / 住宅着工統計 (国土交通省)
       mlit_chakko_starts_by_pref.xlsx  時系列表 (全国: 年度計・月次、利用関係別。
                                        ファイル名に反し中身は全国系列 + 非居住建築物)
       estat_chakko_starts_by_pref_annual.xls
                                        e-Stat 時系列表「都道府県別 新設住宅の戸数・
                                        床面積の合計 (年計)」※本スクリプトで取得可能
       estat_chakko_starts_by_use_annual.xls
                                        e-Stat 時系列表「利用関係別 新設住宅 (年計)」
                                        ※同上 (暦年の裏取り用)
     不採用 (取得したが未使用): e-Stat 構造別・建て方別・利用関係別 年別詳細
       (statInfId=000040405974、約4.2MB)。理由は Excel「出典・注記」参照。

このスクリプトが自動取得するのは URL が確定している e-Stat の2ファイルのみ。
それ以外は上記の対応表を手がかりに各ポータルから手動で取得する (URL を推測しない)。
"""
import pathlib
import urllib.request

RAW = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
ESTAT_DL = "https://www.e-stat.go.jp/stat-search/file-download?statInfId={}&fileKind=0"

FILES = [
    (ESTAT_DL.format("000040405972"), "estat_chakko_starts_by_pref_annual.xls"),
    (ESTAT_DL.format("000040405973"), "estat_chakko_starts_by_use_annual.xls"),
]


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    for url, name in FILES:
        dest = RAW / name
        print(f"GET {url} -> {dest}")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as r, open(dest, "wb") as f:
            f.write(r.read())
    print("done")


if __name__ == "__main__":
    main()
