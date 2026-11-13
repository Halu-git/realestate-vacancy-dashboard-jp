# 不動産・空き家ダッシュボード（都道府県レベル）

都道府県レベルで「空き家ストック × 地価 × 住宅着工フロー」を見る市場ダッシュボードのサンプルです。Excel完成品（日本語版・英語版）、Power BI設計パック、整形済みCSVを含みます。単一エリアの出店診断ではなく、広域の市場概観に特化しています。

## 主な数値

- 空き家数（2023年）： **9,001,600戸**（過去最多）／ 空き家率 **13.8％**（計算値13.84％、公表値と一致）
- 「その他」空き家（賃貸・売却用・二次的住宅を除く）： **3,856,000戸**（総住宅数の5.9％）
- 総住宅数 65,046,700戸 ／ 総世帯数 56,215,300世帯（2023年）
- 空き家率の最高は徳島県・和歌山県（21.3％）、最低は埼玉県（9.3％）
- 地価公示2025： 住宅地 **+2.1％**、商業地 **+3.9％**、全用途 **+2.7％**（全国平均変動率）／ 住宅地平均価格 137,100円/m²
- 新設住宅着工戸数： 暦年2025年 **740,667戸**（前年比 −6.5％）／ 2024年度 816,388戸

## 構成

- `excel/realestate_vacancy_dashboard.xlsx` ― Excel完成品（12シート、ネイティブグラフ7点）
- `excel/realestate_vacancy_dashboard_en.xlsx` ― 英語版Excel（数値は日本語版と同一・全セル突合済み）
- `powerbi/` ― Power BI設計パック（model.md / measures.dax / pages.md、4ページ設計）
- `data/processed/` ― tidy形式CSV 11ファイル（出典列つき）
- `scripts/` ― 取得・整形・Excel生成スクリプト（download_sources.py / build_tidy.py / build_excel.py / build_excel_en.py）
- `design.md` ― 設計メモ（目的・採用理由・検証結果・限界）
- 生データ（約22MB）は容量の都合で同梱していません。取得元は `scripts/download_sources.py` の対応表を参照してください

## データソースと採用理由

| データ | 採用理由 |
|---|---|
| 総務省統計局「住宅・土地統計調査」令和5年 基本集計（確報） | 空き家数・空き家率の公式定義を持つ唯一の専門調査。都道府県別・空き家の種類別・建て方別まで一貫した定義で比較できる。速報ではなく確報を採用 |
| 国土交通省「地価公示」令和7年 | 公的な地価指標で毎年公表される。都道府県別・用途別の平均価格と変動率が揃い、5年周期の住調を年次で補完できる。第5表と推移表の二源で相互検証できる |
| 国土交通省「建築着工統計調査」 | 住宅フローの公式統計。都道府県別・暦年の着工戸数が長期で揃い、全国の年度・月次系列と相互検証できる |
| 世帯数（住宅・土地統計調査 第2-2表・第3表） | ストックの需要側の分母。定義が揃う同一調査から取得 |

取得日： いずれも 2026-11-13。いずれも公開ファイルのみを使用（APIのアプリケーションID不使用）。

### 鮮度に関する重要な注意

- 住宅・土地統計調査は**5年周期**で、空き家の最新値は **2023年10月1日現在** です。2024年以降の空き家動向はこの統計では追えません（次回調査は2028年予定）
- 地価・着工は2025年まで年次で補完していますが、空き家そのものの代替指標ではありません

### 定義の注意

- 「空き家」＝ 居住世帯のいない住宅のうち、一時現在者のみの住宅と建築中の住宅を除いたもの。種類は 賃貸用／売却用／二次的住宅／その他（賃貸・売却用及び二次的住宅を除く空き家）
- 空き家率＝ 空き家数÷総住宅数（本ダッシュボードの計算値。公表値13.8％と小数第1位で一致）
- 地価公示の平均価格は標準地の平均であり、実際の取引価格ではありません

## 採用しなかったデータ

- **住宅・土地統計調査 速報集計**： 早期公表用の概数で確報で値が更新されるため不採用（確報を採用）
- **住宅・土地統計調査 市区町村別結果／地価公示 市区町村別表**： 本サンプルは都道府県粒度に限定したため不採用。市区町村粒度は自治体・事業者の買い手価値が高いため、商品化時は市区町村別レポートとして有料オプションで提供する想定です
- **建築着工統計 構造別・建て方別・利用関係別の年別詳細ファイル**： 粒度が本ダッシュボードを大きく上回り、主要指標は時系列表で代替できるため本編では不採用。商品化時は深掘り分析として有料オプション提供の想定です
- **地価公示 地点別（標準地）生データ**： 都道府県集計は公表表で代替できるため不採用
- **国勢調査の住宅・世帯項目**： 空き家の種類別内訳が取れず、空き家の定義は住調が専門であるため不採用
- **人口動態（住民基本台帳・人口推計）**： 空き家率の公式分母は住宅総数であり、需要側は世帯数で代表できるため本編では不採用。商品化時は人口動態とのクロス分析を有料オプション想定です

## 検証

公表値との突合を実施しています。主な結果： 住調2023の全国主要値（空き家総数・率・その他空き家・総住宅数）は確報の公表値とすべて一致／地価公示2025の都道府県別変動率は2つの公表表の間で141件全件一致／着工の暦年合計は都道府県別表と全国月次の自前集計が3年分とも完全一致。既知の差異は、空き家の種類別内訳の合計が総数と100戸ずれる点のみで、原表の丸めによるものです。詳細は `design.md` と Excel の「検証ログ」シートを参照してください。

## ライセンス・出典の扱い

データの出典は各統計の公表機関（総務省統計局・国土交通省）に帰属します。政府統計のデータは各機関の利用条件に従ってください。本リポジトリのコードとドキュメントはサンプルとして公開しています。

## English summary

**Real Estate & Vacancy Dashboard — Japan, by prefecture.** A sample market dashboard combining three official statistics systems at prefecture level: vacant-dwelling *stock* (Housing and Land Survey 2023), *land prices* (Official Land Prices 2025) and housing-start *flows* (Building Starts Statistics). Includes a finished Excel dashboard in Japanese and English (identical numbers, verified cell by cell), a Power BI design pack (data model, DAX measures, 4 page layouts), and tidy CSVs with source columns.

Key figures: 9,001,600 vacant dwellings in 2023 — a record high — for a vacancy rate of 13.8%; 3,856,000 of them fall in the "other" category (excluding dwellings for rent, for sale and second dwellings), concentrated in detached houses. Official Land Prices rose in 2025 (residential +2.1%, commercial +3.9%), while new housing starts fell to 740,667 units in CY2025 (−6.5% YoY).

**Freshness limitation:** the Housing and Land Survey is conducted every five years, so vacancy data is current as of 1 October 2023; the next survey is planned for 2028. Land prices and starts are updated annually and complement — but do not substitute for — the vacancy figures. All headline values were cross-checked against published official figures; checks and reasons for including/excluding each dataset are documented in the Excel workbook ("Validation Log" and "Sources & Notes" sheets) and in `design.md`.
