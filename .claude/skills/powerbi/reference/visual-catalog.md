# Visual catalog (v1)

| type | roles (required → optional) | use it for |
|---|---|---|
| `card` | Values(1) | one KPI number (총 매출, 앱 수) |
| `multiRowCard` | Values(n) | several small KPIs in one box |
| `gauge` | Y(1) → MinValue, MaxValue, TargetValue | progress toward a target |
| `lineChart` / `areaChart` | Category(1), Y(n) → Series(1) | trend over time (Category = Calendar.Month / YearMonth / Year) |
| `clusteredColumnChart` / `stackedColumnChart` | Category(1), Y(n) → Series(1) | compare few categories; stacked for composition |
| `clusteredBarChart` / `stackedBarChart` | Category(1), Y(n) → Series(1) | ranking of many categories (add `sort` desc + `topN`) |
| `pieChart` / `donutChart` | Category(1), Y(1) | share of ≤ 6 categories |
| `scatterChart` | Category(1), X(1), Y(1) → Size(1) | correlation; Category = the entity (product, app) |
| `tableEx` | Values(n) | detail rows |
| `pivotTable` | Rows(n), Values(n) → Columns(n) | matrix / cross-tab |
| `slicer` | Values(1) | user filter (category, year) |
| `textbox` | — (`text`, `fontSize`) | page title / notes |

Rules of thumb:
- Y / Values / X / Size take **measures or aggregated columns** (`sum:`), Category / Series / Rows / Columns / slicer take **columns**.
- Page layout default: row 0 textbox title `[0,0,12,1]`; row 1 KPI cards `[0,1,3,2] [3,1,3,2] [6,1,3,2] [9,1,3,2]`; rows 3–7 two charts `[0,3,6,5] [6,3,6,5]` or one wide chart + slicers.
- Max ~8 visuals per page, 1–4 pages. Use `topN` for bar charts over > 15 categories.
- Korean titles for every visual unless the user writes in English.
