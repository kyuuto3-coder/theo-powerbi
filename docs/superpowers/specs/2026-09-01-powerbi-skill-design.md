# Power BI Skill — Design Spec

**Date:** 2026-09-01
**Repo:** https://github.com/kyuuto3-coder/theo-powerbi (this folder is the repo root)
**Status:** approved design, pending implementation plan

## 1. Goal

A Claude Code project skill that lets a non-technical coworker, working in **Claude Code Desktop on Windows**:

1. drop one or more `.csv` / `.xlsx` files into `rawdata/`,
2. describe the dashboard they want (text, a mockup screenshot, or preferences),
3. receive a ready-to-open **Power BI Project (`.pbip`)** that they open in Power BI Desktop and save as `.pbix`.

Hard constraints:

- **Minimal tokens** — coworkers are on a Pro plan. Claude must never read raw data files or generated Power BI files; it reads a compact profile and writes one small spec.
- **Zero install** — only Windows PowerShell 5.1 (built into Windows) and Power BI Desktop. No Python, no Node, no MCP server.
- **Korean-facing** — Claude talks to coworkers in Korean; skill internals (SKILL.md, scripts, reference docs) are English.
- **Deterministic output** — a PowerShell builder, not Claude, emits every PBIR/TMDL file.

The existing `../worldwideimporters` PBIP and `../wwi-model-tmdl` are **format references only** (PBIR v2.0.0 report definition, TMDL semantic model). They are not part of the repo.

## 2. Repository layout

```
theo-powerbi/                       ← coworkers clone or download ZIP, then open in Claude Code Desktop
├── README.md                       ← Korean 5-step guide + prerequisites
├── CLAUDE.md                       ← 3 lines: use the powerbi skill for any data/dashboard request; speak Korean
├── .gitignore                      ← rawdata/*  output/*  (keep .gitkeep in both)
├── rawdata/                        ← coworker input (.csv / .xlsx)
├── output/                         ← generated: output/<Name>/report-spec.json, output/<Name>/<Name>.pbip, …
├── docs/superpowers/specs/         ← this document (+ plans)
└── .claude/skills/powerbi/
    ├── SKILL.md                    ← workflow + token rules (English, < 150 lines)
    ├── scripts/
    │   ├── profile-data.ps1        ← rawdata/ → compact text profile
    │   ├── build-pbip.ps1          ← report-spec.json → validated PBIP project
    │   ├── build-guide.ps1         ← report-spec.json → manual-guide.md (section 10)
    │   ├── lib/*.ps1               ← shared: io, xlsx, csvread, infer, spec, tmdl, pbir, guide
    │   └── tests/                  ← run-tests.ps1 harness, *.test.ps1, fixtures/
    ├── templates/
    │   ├── report.json             ← copied from the WWI reference (theme + report settings)
    │   └── theme/Fluent2-CY26SU08.json ← base theme copied from the WWI reference
    │   (visual skeletons live in code: the visual catalog in scripts/lib/spec.ps1 / pbir.ps1)
    ├── reference/
    │   ├── spec-schema.md          ← the report-spec.json contract (section 5), with a complete example
    │   ├── visual-catalog.md       ← which visual answers which question; roles each accepts
    │   └── dax-patterns.md         ← measure recipes (totals, averages, % share, YoY, running total, top-N)
    └── examples/app-data-spec.json ← worked example spec for a flat table (App Data.xlsx shape)
```

Claude reads `SKILL.md` always, `reference/*.md` only when needed (SKILL.md says which situation triggers which file), and **never** reads `templates/`.

## 3. Coworker workflow (README.md, Korean)

1. Prerequisites: Power BI Desktop (Store or installer). In *옵션 → 미리 보기 기능*, enable **"Power BI 프로젝트(.pbip) 저장 옵션"** and **"향상된 보고서 형식(PBIR)으로 보고서 저장"** if present (versions where these are GA have no toggle).
2. Put data files in `rawdata/`.
3. Open the folder in Claude Code Desktop, describe the dashboard (attach a mockup image if you have one).
4. Claude produces `output/<Name>/<Name>.pbip`. Double-click it, click **새로 고침**, then **파일 → 다른 이름으로 저장 → .pbix**.
5. Iterate with Claude ("2페이지의 차트를 막대로 바꿔줘"). **Finish iterating with Claude before editing in Power BI Desktop** — a rebuild overwrites the generated project.
6. If Power BI shows an error on open/refresh, paste the error text to Claude.
7. **Changing the data later.** `rawdata/` is the coworker's own folder — the scripts only read it, never write or delete there, and git ignores it. Replacing a file with a newer version that keeps the **same file name, sheet name and columns** needs no Claude involvement: open the `.pbix`/`.pbip` and click **새로 고침**. Adding files, renaming a file/sheet, or adding/removing columns → tell Claude ("데이터 파일을 바꿨어요"); Claude re-runs the profiler and rebuilds (or edits the spec's `tables[]`). The generated M query keeps only the columns listed in the spec, so extra new columns are ignored until the spec is updated.

## 4. Claude workflow (SKILL.md)

```
1. PROFILE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1
             → prints the profile (section 6). Never open rawdata/ files with Read/cat.
2. UNDERSTAND
             - If the user attached an image: view it once, extract page list, visual types, rough layout, KPIs.
             - Ask at most 2 clarifying questions, in Korean, only when the profile cannot resolve
               an ambiguity that changes the result (e.g. two candidate revenue columns).
               Otherwise choose sensible defaults and state assumptions in the final message.
3. DESIGN    Choose schema mode (flat | star, section 7), measures (5–10), pages (1–4), visuals (≤ 8 per page).
             Consult reference/visual-catalog.md for role names; reference/dax-patterns.md for DAX only if
             a non-trivial measure is needed.
4. WRITE     output/<Name>/report-spec.json (Write tool). Target ≤ 200 lines.
5. BUILD     powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec output/<Name>/report-spec.json
             On validation error: fix the spec with a minimal Edit, rebuild. Do not rewrite the whole spec.
6. HANDOFF   Korean summary: what was built (pages/visuals/measures), assumptions, how to open, how to iterate.
7. GUIDE?    Ask in Korean whether they also want a manual how-to guide (section 10). If yes:
             powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1 -Spec output/<Name>/report-spec.json
             then point to output/<Name>/manual-guide.md. Do not read the guide back.
```

Token rules (verbatim in SKILL.md):

- Never read files in `rawdata/`, `output/*/`, or `templates/`. The scripts own them.
- Do not re-read the spec after writing it; use `Edit` for changes.
- Read `reference/*.md` only when the trigger listed in SKILL.md applies.
- Keep clarifying questions to ≤ 2, and only when truly ambiguous.
- On a Power BI error pasted by the user, reason from the error text and the spec; do not regenerate from scratch.
- Never run the profiler more than once per session unless the user says `rawdata/` changed (then re-run it — never guess the new columns).

## 5. `report-spec.json` contract

```jsonc
{
  "name": "AppStoreDashboard",          // folder + .pbip name; letters/digits/_/- only
  "locale": "ko-KR",                    // model culture
  "autoDateTable": true,                // build a Calendar table if any date column exists and no date dim
  "tables": [
    {
      "name": "Apps",                   // model table name; no "." allowed
      "file": "App Data.xlsx",          // relative to rawdata/
      "sheet": "Data",                  // xlsx only; null for csv
      "headerRow": 1,                   // 1-based row in the sheet/csv holding headers (default 1)
      "encoding": 65001,                // csv only: 65001 (UTF-8) or 949 (CP949); from the profile
      "delimiter": ",",                 // csv only (default ",")
      "columns": [                      // ONLY listed columns are loaded (others dropped in M)
        { "name": "id", "type": "int64", "key": true, "hidden": true },
        { "name": "track_name", "rename": "앱 이름", "type": "text", "description": "Track Name" },
        { "name": "price", "type": "decimal", "format": "$#,0.00" },
        { "name": "user_rating", "type": "double", "summarizeBy": "average" },
        { "name": "release_date", "type": "date" }
      ]
    }
  ],
  "relationships": [                    // star mode only; omit or [] for flat
    { "from": "Sales.CityKey", "to": "City.CityKey", "active": true, "crossFilter": "single" }
  ],
  "measures": [
    { "table": "Apps", "name": "앱 수", "dax": "COUNTROWS(Apps)", "format": "#,0" },
    { "table": "Apps", "name": "평균 평점", "dax": "AVERAGE(Apps[user_rating])", "format": "0.00" }
  ],
  "pages": [
    {
      "name": "개요",
      "filters": [ { "field": "Apps.prime_genre", "in": ["Games", "Education"] } ],   // optional page filter
      "visuals": [
        { "type": "textbox", "text": "앱스토어 대시보드", "grid": [0,0,12,1], "fontSize": 24 },
        { "type": "card", "title": "앱 수", "fields": { "Values": ["Apps.앱 수"] }, "grid": [0,1,3,2] },
        { "type": "clusteredBarChart", "title": "장르별 앱 수",
          "fields": { "Category": ["Apps.prime_genre"], "Y": ["Apps.앱 수"] },
          "sort": { "field": "Apps.앱 수", "direction": "desc" },
          "topN": { "n": 10, "by": "Apps.앱 수" },
          "grid": [3,1,5,4] },
        { "type": "scatterChart", "fields": { "Category": ["Apps.앱 이름"], "X": ["avg:Apps.price"], "Y": ["Apps.평균 평점"], "Size": ["sum:Apps.rating_count_tot"] }, "grid": [8,1,4,4] },
        { "type": "slicer", "fields": { "Values": ["Apps.prime_genre"] }, "grid": [0,3,3,2] }
      ]
    }
  ]
}
```

Rules:

- **Field reference** = `Table.Name` (split on the first `.`; table names may not contain `.`). The builder resolves `Name` against measures of that table first, then columns (post-`rename` names). Aggregation prefix `sum: | avg: | min: | max: | count: | countDistinct:` is allowed on columns only.
- **Types**: `int64 | double | decimal | text | date | datetime | boolean` → TMDL `dataType` and M `TransformColumnTypes` type. `summarizeBy` defaults: numeric → `sum`, `key`/text/date → `none`.
- **grid** = `[col,row,w,h]` on a 12 × 8 grid of a 1920 × 1080 page, 40 px margin, 20 px gutter (cell 135 × 107.5 px, rounded). `x = 40 + col·155`, `y = 40 + row·127.5`, `w = cw·135 + (cw−1)·20`, `h = rh·107.5 + (rh−1)·20`.
- **Visual roles** per type (PBIR `queryState` keys):

| type | required | optional |
|---|---|---|
| card | Values (1) | |
| multiRowCard | Values | |
| gauge | Y (1) | MinValue, MaxValue, TargetValue |
| lineChart, areaChart | Category, Y | Series |
| clusteredColumnChart, stackedColumnChart, clusteredBarChart, stackedBarChart | Category, Y | Series |
| pieChart, donutChart | Category, Y | |
| scatterChart | Category, X, Y | Size |
| tableEx | Values | |
| pivotTable | Rows, Values | Columns |
| slicer | Values (1) | |
| textbox | — (`text` required) | fontSize |

- `sort`, `topN`, `filters` are optional. `topN.by` must be a measure or aggregated column ref; `filters[].in` values are literals matched against a column.
- `autoDateTable`: when true and a `date`/`datetime` column exists in a table that has no relationship to a date-typed dimension, the builder adds a calculated table `Calendar` (`CALENDAR(MIN(all date columns), MAX(...))` with `Year`, `Quarter`, `Month` (sortable by `MonthNo`), `Date`), marks it as the date table, and relates it to the first date column of each such table (additional date columns get inactive relationships). Claude may then reference `Calendar.Year` / `Calendar.Month` in visuals.

## 6. `profile-data.ps1`

**Input:** `-DataFolder` (default `rawdata`). **Output:** plain text to stdout, ≤ ~60 lines for typical inputs (hard cap: 25 columns per table shown, rest summarised as "+N more columns").

```
PROFILE rawdata/  (2 files)
FILE "App Data.xlsx"
  SHEET "Data"  rows=7197  header_row=1  first_col=B  → table candidate
    col  name             type    nulls  distinct  sample                     note
    A    (unnamed)        int64   0%     7197      0, 1, 2                    index? drop
    B    id               int64   0%     7197      281656475, 281796108       KEY
    C    track_name       text    0%     7195      "PAC-MAN Premium"
    E    price            decimal 0%     37        3.99, 0                    min=0 max=299.99
    ...
  SHEET "Data Defintion"  rows=10  → dictionary (not loaded)
    id=App ID; track_name=Track Name; size_bytes=App Size in Bytes; ...
FILE "sales.csv"  encoding=949  delimiter=","  rows=228265  header_row=1  → table candidate
    ...
KEY MATCHES (star candidates): sales.City Key ↔ city.City Key (100% of 12 distinct)   ← only when ≥ 2 tables
```

Detection rules:

- Header row = first row where ≥ 60 % of non-empty cells are text and the following row has a different type mix; `first_col` = first column with a non-empty header.
- Type inference over the first 10,000 rows (whole table for typical files; the summary size does not depend on row count): int64 (all integers), decimal (numeric with ≤ 4 decimals, or currency-like format), double (other numeric), date/datetime (xlsx: cell style numFmt is a date format; csv: ≥ 95 % of samples parse as date), boolean (true/false/0/1 only with header hinting), else text.
- CSV encoding: BOM → UTF-8; else try UTF-8 strict decode on the first 64 KB; on failure → 949 (CP949).
- Dictionary sheet: ≤ 3 non-empty columns, ≤ 60 rows, header contains name/description-like words (`name|field|column|컬럼|항목` and `desc|설명|meaning`), or > 70 % of first-column values equal a header in another sheet.
- Key candidate: distinct == non-null == sampled rows (`KEY` when the sample covers the whole table, `KEY?` otherwise). Star candidate: for every pair of tables, a column whose normalized name (letters/digits only, case-insensitive) matches a column that is unique on the other side. Claude confirms the match from the profile; no containment scan is done (keeps the profiler fast on 100k+ row files).
- xlsx is read via `System.IO.Compression` + XML (no Excel COM, no modules). Shared strings and styles are parsed once per workbook.

## 6b. Multi-file capabilities (added 2026-09-01 after review)

- **Value-based key matching**: `KEY MATCHES` lists `many.col -> one.key (NN% of sampled values found)` using canonical value overlap, so differently named keys (`cust_no` → `고객ID`) are found. Different-name matches require both names to look like ids or a key with ≥ 20 distinct values (guards against tiny code ranges). Unique-key columns of tables larger than the sample are re-read in full (up to 400k rows) so containment is exact.
- **Guard rails for different-name matches** (tuned on the 350k-row WWI set): fact→dimension direction (many side has ≥ rows), ≥ 3 distinct values, numeric many-side spans ≥ 50% of the key's full range, dense 1..N keys require an id-like many-side name, a column with a same-name key gets no other guesses, and subject words must be compatible by prefix (`cust` ~ `customer`; Korean vs English is not judged). Value-strong but name-unrelated pairs are listed under `POSSIBLE MATCHES` for the user to confirm. Sample is 10,000 rows (summary size is independent of rows).
- **Composite keys**: when a table has no single unique column, the profiler searches 2–3 column combinations among low-cardinality dimension-like columns and prints `KEY (composite): A + B`; matches against other tables' tuples are reported as `composite key`. The spec accepts arrays on both sides of a relationship; the builder adds hidden text columns `_key_<parts>` (M `Text.Combine`) on both tables and relates them (Power BI relationships are single-column).
- **Same-structure files**: `GROUPS` lists files with identical headers and a `filePattern`; the spec's `filePattern` (one `*`) makes the builder emit `Folder.Files(DataFolder)` + `Table.Combine`, so new files matching the pattern load on Refresh.
- Column cap raised from 25 to 200 per table.

## 7. Schema modes

- **flat** — exactly one table candidate: no relationships, `key` on the detected key, measures = COUNTROWS, SUM/AVERAGE of numeric non-key columns as appropriate, `autoDateTable` if a date column exists.
- **star** — ≥ 2 table candidates with key matches: the largest table(s) with numeric measures become facts; tables with a unique key become dimensions; relationships many→one, single-direction, first date column of the fact → date dimension if one exists, else `Calendar`.

Claude decides the mode from the profile; the builder does not infer anything — it only validates and emits.

## 8. `build-pbip.ps1`

**Input:** `-Spec <path>` (`-Force` to overwrite an existing output project; default: overwrite `definition/` folders but keep `.pbi/` caches). **Output:** `output/<Name>/<Name>.pbip`, `<Name>.Report/`, `<Name>.SemanticModel/`.

Validation (all errors collected, then printed as `ERROR <n>: <message> — spec path: pages[0].visuals[2].fields.Y[0]`, exit code 1):

1. JSON parses; required top-level keys present; `name` matches `^[A-Za-z0-9_-]+$`.
2. Every `tables[].file` exists under `rawdata/`; `sheet` given for xlsx; table names unique, no `.`.
3. Column names unique per table (after rename); types in the allowed set.
4. Relationship endpoints resolve to columns; no duplicate pairs.
5. Measure names unique per table and not colliding with column names; `dax` non-empty (no DAX parsing).
6. Visual `type` in the catalog; required roles present; role cardinality respected; every field ref resolves; aggregation prefixes only on columns; `grid` inside 12 × 8 and no negative sizes (overlaps are allowed but reported as `WARN`).
7. `sort.field`, `topN.by`, `filters[].field` resolve.

Emission:

- **Semantic model (TMDL)** — `definition.pbism`, `definition/database.tmdl`, `model.tmdl` (culture from `locale`, `PBI_QueryOrder`, `ref table` lines, `__PBI_TimeIntelligenceEnabled = 0` when `autoDateTable` is used), `relationships.tmdl`, `tables/<Table>.tmdl` with columns (`dataType`, `formatString`, `summarizeBy`, `isHidden`, `isKey`, `description`, `lineageTag` GUIDs, `sortByColumn`), measures, and an `m` partition whose M query is generated from the table's source:
  - csv: `Csv.Document(File.Contents("<abs path>"), [Delimiter=…, Encoding=…, QuoteStyle=QuoteStyle.Csv])` → `Table.Skip(headerRow−1)` → `Table.PromoteHeaders` → `Table.SelectColumns(listed source names)` → `Table.RenameColumns` → `Table.TransformColumnTypes`.
  - xlsx: `Excel.Workbook(File.Contents("<abs path>"), null, true){[Item="<sheet>",Kind="Sheet"]}[Data]` → same chain.
  - `<abs path>` is the resolved absolute path of `rawdata/<file>` **on the machine running the builder**, written verbatim (M string literals do not escape backslashes; only `"` is doubled). On macOS runs (dev only) the POSIX path is written as-is and a `WARN` says the project must be rebuilt on Windows before opening.
  - `Calendar` calculated table when `autoDateTable` triggers.
- **Report (PBIR)** — `definition.pbir` (byPath to the model), `definition/version.json`, `report.json` (theme, settings copied from the reference), `pages/pages.json` (`pageOrder`, `activePageName` = first page), `pages/<id>/page.json` (1920 × 1080, `FitToPage`), `pages/<id>/visuals/<id>/visual.json` from `templates/visuals/<type>.json` with `position`, `queryState` projections (`Column` vs `Measure` vs `Aggregation` expressions), `sortDefinition`, `filterConfig` (topN / in), `objects.title` (title text), textbox `paragraphs`. IDs are 20-hex-char random strings; `tabOrder` increments by 1000 per visual.
- **Project** — `<Name>.pbip` referencing `<Name>.Report`; `.gitignore` for `.pbi/`.
- **File encoding** — UTF-8 without BOM, CRLF line endings (matches Power BI Desktop's own output). Written with `[IO.File]::WriteAllText` + `UTF8Encoding($false)`. `ConvertTo-Json -Depth 32`; Korean text may be `\uXXXX`-escaped in JSON (valid); TMDL is written raw UTF-8.

Exit code 0 prints one line: `BUILT output/<Name>/<Name>.pbip  (tables=N measures=N pages=N visuals=N)` plus any `WARN` lines.

## 9. Testing

- `scripts/test-build.ps1` — runs the profiler on a fixture folder (`scripts/fixtures/`: a 30-row csv in CP949, a 2-sheet xlsx with offset header + dictionary sheet), builds `examples/app-data-spec.json`, and asserts: exit 0, expected files exist, every `.json` parses, every visual's `queryState` keys equal the spec roles, the M query contains the absolute path, the generated TMDL has one `partition` per table and one `measure` per spec measure. Also asserts that a deliberately broken spec (unknown column, bad role, bad type) yields exit 1 with three `ERROR` lines.
- Runs under **Windows PowerShell 5.1** on the Parallels VM via SSH (the real target), with the repo reached through `\\Mac\Home\...`. Also runs under PowerShell 7 on macOS if installed (nice-to-have, not required).
- **End-to-end**: open the generated `.pbip` in Power BI Desktop on the VM (`Start-Process`), then connect the `powerbi-modeling` MCP to the running instance and assert the model exposes the spec's tables, measures, and relationships, and that a DAX `EVALUATE` of a measure returns a value. This is the acceptance test for "opens without error".
- Structural diff of the generated report tree against the WWI reference (same file set per page/visual, same `$schema` versions).

## 10. Manual guideline (`build-guide.ps1`)

After the hand-off (workflow step 6) Claude asks, in Korean, whether the coworker also wants a guide for building the same report **by hand** in Power BI Desktop. If yes, Claude runs `build-guide.ps1 -Spec output/<Name>/report-spec.json`, which writes `output/<Name>/manual-guide.md` deterministically from the spec — Claude does not author the guide (zero extra tokens beyond one command and a one-line answer).

Guide rules:

- **Prose in Korean; every menu, ribbon tab, pane, button, visual name, field-well name, data type and dialog label in strict English**, exactly as Power BI Desktop's English UI shows them (e.g. **Home > Get Data > Excel Workbook**, **Transform Data**, **Use First Row as Headers**, **Choose Columns**, **Close & Apply**, **Model view**, **Manage relationships**, **New measure**, **Mark as date table**, **Visualizations** pane, **Clustered bar chart**, **X-axis / Y-axis / Legend / Values / Rows / Columns / Field / Data**, **Filters** pane → **Top N**, **Format visual > General > Title**).
- Sections, generated from the spec in this order: 0 준비 (files, Power BI Desktop, canvas size), 1 데이터 가져오기 (one subsection per table: Get Data → Navigator → Power Query steps: Remove Top Rows when `headerRow` > 1, Use First Row as Headers, Choose Columns listing the spec columns, renames, data types with the Power Query type names `Whole Number / Decimal Number / Fixed decimal number / Text / Date / Date/Time / True/False`, query rename, Close & Apply; CSV adds **File Origin** `65001: Unicode (UTF-8)` or `949: Korean` and **Delimiter**), 2 모델 (hide columns, Summarization, key, relationships via **Manage relationships > New** with cardinality/direction/active, Calendar table via **New table** with the DAX pasted in a code block, **Mark as date table**, **Sort by column**), 3 측정값 (one **New measure** per measure with the DAX in a code block and its **Format**), 4 보고서 페이지 (per page: rename page, page filters; per visual: the Visualizations-pane name, each field well and what to drag there, **Sort axis** direction, **Top N** filter setup, title text, and the exact **Size** and **Position** from the grid math so the layout matches the generated report), 5 저장 (**File > Save As** `.pbix`).
- Field-well names per visual (English UI): card → Data; multi-row card → Fields; gauge → Value / Minimum value / Maximum value / Target value; line/area/column charts → X-axis / Y-axis / Legend; bar charts → Y-axis (category) / X-axis (value) / Legend; pie/donut → Legend / Values; scatter → Values (details) / X Axis / Y Axis / Size; table → Columns; matrix → Rows / Columns / Values; slicer → Field; text box → **Insert > Text box**.

## 11. Out of scope (v1)

Maps and custom visuals; bookmarks, drillthrough, tooltips pages; row-level security; DirectQuery/database sources; editing an existing `.pbix`/`.pbip`; publishing to Fabric/Service; incremental refresh; Power Query transformations beyond column select/rename/type; multiple files merged into one table (folder sources).

## 12. Distribution

The repo root is the workspace. Coworkers: **Code → Download ZIP** (or `git clone`), extract, open the folder in Claude Code Desktop. Updates are pulled by re-downloading; `rawdata/` and `output/` are gitignored so their work is never overwritten. `.mcp.json` and `.claude/settings.local.json` are not committed.

## Acceptance record

- **2026-09-01** — 52 tests pass on Windows PowerShell 5.1 (Parallels VM). `examples/app-data-spec.json` built against the real `rawdata/App Data.xlsx` (7,197 rows; 1 table, 5 measures, 2 pages, 12 visuals) and the star fixture (3 tables incl. Calendar, 2 relationships) both load without errors through the Tabular Object Model (`powerbi-modeling` MCP `ConnectFolder`). Opening the `.pbip` in Power BI Desktop's GUI could not be automated from the SSH session (no interactive desktop); to be confirmed by hand on Windows.
- Implementation notes: every `.ps1` is UTF-8 **with BOM** (Windows PowerShell 5.1 reads BOM-less scripts in the ANSI code page and corrupts Korean literals); PowerShell variable names are case-insensitive (`$visual` vs `$Visual` collided once); visual skeletons live in the in-code catalog rather than template files.
