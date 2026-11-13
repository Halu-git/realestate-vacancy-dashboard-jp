# Power BI モデル設計 ― 不動産・空き家ダッシュボード

入力： `data/processed/` の tidy CSV 群（11ファイル）。Excel版と同じCSVをそのまま使う。

## テーブル構成（調査系統ごとにレイヤーを分ける）

### S1 住宅・土地統計調査（主レイヤー・ストック）
| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactVacancySnapshot | vacancy_snapshot_pref_2023.csv | 都道府県（全国含む48行）×2023年 | total_dwellings, occupied_dwellings, vacant_total, vacant_other, vacant_for_rent, vacant_for_sale, vacant_second, vacant_villa, under_construction, vacancy_rate_pct_calc, vacant_other_rate_pct_calc |
| FactVacancyTrendNat | vacancy_trend_national.csv | 年（1958〜2023、調査年のみ14行）×全国 | total_dwellings, households, vacant_total, 種類別4列, vacancy_rate_pct_calc |
| FactVacancyStructure | vacancy_by_structure_pref_2023.csv | 都道府県×建て方（5区分）×2023年 | vacant_dwellings |
| FactVacancyComposition | vacancy_composition_national.csv | 空き家種類（5区分）×全国×2023年 | total, detached, tenement, apartment, other_structure |
| FactDwellingsPref | dwellings_households_pref.csv | 都道府県×年（2003/2008/2013/2018/2023） | total_dwellings, households |
| FactHouseholdsSnapshot | households_snapshot_pref_2023.csv | 都道府県×2023年 | households_total, principal_households, household_members, persons_per_household_calc |

### S2 地価公示（価格レイヤー・年次更新）
| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactLandRatePref | land_price_rate_pref_2025.csv | 都道府県×用途（住宅地/商業地/工業地） | rate_2024_pct, rate_2025_pct, survey_points_2025 |
| FactLandSnapshot | land_price_snapshot_2025.csv | 都道府県×用途（全用途含む4区分） | avg_price_2025_yen_m2, rate_2025_pct |
| FactLandTrendNat | land_price_trend_national.csv | 年（1975〜2025）×用途×全国 | avg_price_yen_m2, rate_pct |

### S3 建築着工統計（フローレイヤー・年次更新）
| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactStartsPref | starts_pref_annual.csv | 都道府県×暦年（2014〜2025） | new_housing_starts |
| FactStartsNational | starts_national_annual.csv | 年×年種別（年度/暦年） | total_starts, owner_occupied, for_rent, company_housing, for_sale, of_which_condo, of_which_detached_fs, status |

## ディメンション

- DimPrefecture： 全国＋47都道府県（pref_code 順に並び替え。「全国」は is_total フラグを持たせ、ランキング系ビジュアルでは除外する）
- DimYearVacancy： 住調の調査年（1958〜2023 の14年。5年周期で等間隔でない点に注意）
- DimYear： 暦年（地価 1975〜2025、着工 2014〜2025 の共通軸として使う場合のみ）
- DimLandUse： 住宅地／商業地／工業地／全用途
- DimStructure： 一戸建／長屋建／共同住宅／その他（「総数」はディメンションに含めず、メジャー側で分離する）
- DimVacancyType： 総数／賃貸・売却用及び二次的住宅を除く空き家／賃貸用／売却用／二次的住宅

## リレーション方針

- S1・S2・S3 は**調査系統が異なるため、ファクト同士を直接リレーションしない**。DimPrefecture のみ共通ディメンションとして各ファクトへ単方向フィルタする
- 空き家率・その他率（*_pct_calc）は原表の値ではなく計算値。率同士を SUM しないこと（メジャーで再計算する）
- FactStartsNational の年度と暦年は year_type で区別された別系列。同一ビジュアルに混在させない。status =「年度途中(未確定)」の行（FY2025）は既定で除外するフィルタをページ側に置く

## 注意（実装時に落とし穴になる点）

- **鮮度が系統ごとに違う**： S1 は2023年10月1日現在（5年周期）、S2・S3 は2025年まで年次更新。ページ上で「いつの時点の数字か」を必ず明示する
- FactVacancySnapshot の種類別内訳（vacant_other + vacant_for_rent + vacant_for_sale + vacant_second）は原表の丸めで vacant_total と最大100戸ずれる。合計チェックでエラー扱いしない
- FactLandRate の工業地は地点数が少ない県があり、欠測（空欄）がある。0として扱わない
- 地価の平均価格は「標準地の平均」であって取引価格ではない。ビジュアルの注記に必須
- FactStartsPref の2023年の全国値（819,623戸）は FactStartsNational の暦年2023と一致する（突合済み）。どちらを正とするかはページ単位で統一する
