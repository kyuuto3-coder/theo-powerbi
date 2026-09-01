# DAX measure recipes

Replace `T`, `[Amount]`, `[Qty]`, `[Cost]` with model names. `Calendar` exists when `autoDateTable` created it.

| purpose | dax |
|---|---|
| total | `SUM(T[Amount])` |
| row count | `COUNTROWS(T)` |
| distinct count | `DISTINCTCOUNT(T[CustomerId])` |
| average | `AVERAGE(T[Rating])` |
| profit | `[Revenue] - [Cost]` (measures referenced with `[ ]`) |
| margin % | `DIVIDE([Profit], [Revenue])` → format `0.0%` |
| share of total | `DIVIDE([Revenue], CALCULATE([Revenue], ALL(T)))` → `0.0%` |
| average per row | `DIVIDE([Revenue], [Row Count])` |
| year-to-date | `TOTALYTD([Revenue], Calendar[Date])` |
| previous year | `CALCULATE([Revenue], SAMEPERIODLASTYEAR(Calendar[Date]))` |
| YoY % | `DIVIDE([Revenue] - [Revenue PY], [Revenue PY])` → `0.0%` |
| running total | `CALCULATE([Revenue], FILTER(ALL(Calendar[Date]), Calendar[Date] <= MAX(Calendar[Date])))` |
| top-N rank | `RANKX(ALL(T[Product]), [Revenue])` |
| count where | `CALCULATE(COUNTROWS(T), T[Status] = "Done")` |
| conditional total | `CALCULATE([Revenue], T[Price] > 0)` |

Format strings: integers `#,0`; currency `$#,0` or `₩#,0`; decimals `0.00`; percent `0.0%`.
Multi-line DAX is fine in the spec (`"dax": "VAR x = ...\nRETURN ..."`).
