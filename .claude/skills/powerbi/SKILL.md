---
name: powerbi
description: Build a Power BI Desktop project (.pbip) from the CSV/XLSX files in rawdata/ based on the user's description, preferences or mockup image. Use for ANY request about dashboards, reports, charts, KPIs, data analysis, visualization, or Power BI. Talk to the user in Korean.
---

# Power BI report builder

You turn `rawdata/*.csv|xlsx` + the user's wishes into `output/<Name>/<Name>.pbip`. Two PowerShell scripts do the heavy lifting; you only read a compact profile and write one small spec. **The user is non-technical and on a limited token budget: follow the token rules exactly.** This skill is written to be followed literally by Claude Sonnet: when a step says run a command, run exactly that command; when it says decide, decide with the defaults given here rather than asking.

## Language rule (hard)
- **Chat with the user in Korean.** Everything *inside* the report must be **English**: table names, column `rename`s, measure names, page names, visual titles, textbox text, descriptions. Power BI Desktop's UI is English and the files are shared with English-speaking readers. Korean source headers are fine in `columns[].name` (that is the raw header) — give them an English `rename`.

## Token rules (hard)
- NEVER open files in `rawdata/`, `output/*/`, or `templates/` with Read/cat/head. Scripts read them.
- Run the profiler once per session (again only if the user says the data changed).
- Read `reference/*.md` only when the trigger below applies. Do not read `scripts/`.
- Write the spec once with Write; afterwards change it only with small Edits. Do not re-read it.
- At most 2 clarifying questions, only when the answer changes the result. Otherwise decide, and list assumptions at the end.
- On a Power BI error pasted by the user: reason from the error text + spec; fix with an Edit; rebuild. Never regenerate from scratch.
- Keep replies short. No progress narration.

## Workflow
1. **PROFILE** — run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1`
   Output = files → tables (columns, type, nulls, distinct, samples, `KEY` / `KEY (composite): A + B`), dictionary sheets, csv `encoding`, `header_row`, `GROUPS` (files with identical columns → one table via `filePattern`), and `KEY MATCHES` (`many.col -> one.key (NN% of sampled values found[; names differ][; composite key])`). Exit code 2 = no files → ask the user (Korean) to put files in `rawdata/`.
2. **UNDERSTAND** — read the user's text; view an attached mockup image once and note: pages, visual types, KPIs, layout (colors are ignored; the theme is fixed). Map their words to columns from the profile. If a dictionary sheet exists, use its descriptions for column `description` and for choosing English names.
3. **DESIGN** — decide:
   - *Schema:* one table candidate → **flat** (no relationships). Several tables → **star** from `KEY MATCHES`: every line with ≥ 90% becomes a relationship (`"from": "Fact.col", "to": "Dim.key"`); names differing is fine — the % is the evidence. A `composite key` line becomes an array relationship: `"from": ["Fact.A","Fact.B"], "to": ["Dim.A","Dim.B"]` (the builder creates the hidden key columns). 70–89% or any `POSSIBLE MATCHES` line → ask the user one question ("A의 X가 B의 Y를 가리키나요?"); < 70% → no relationship. Facts = big tables with numeric measures, dims = tables with a KEY. Tables with no matches → independent tables, no relationships.
   - *Multiple files, same columns:* a `GROUPS` line → ONE table with `"filePattern": "orders_*.csv"` (no `file`); new files matching the pattern load on Refresh. Never create one table per monthly file.
   - *Columns:* list only what the report needs + keys. Drop `(unnamed)`/index columns. `hidden: true` for keys. Types from the profile (`decimal` for money, `double` for ratios/ratings). English `rename` for every user-facing column (translate Korean headers, e.g. `매출액` → `Revenue`).
   - *Measures:* 4–10. Always a count (`COUNTROWS`) and totals/averages of the main numeric columns; add ratio measures where sensible. English names always (`App Count`, `Avg Rating`, `Total Revenue`). Trigger → read `reference/dax-patterns.md` only for time intelligence / ranking / running totals.
   - *Pages/visuals:* 1–4 pages, ≤ 8 visuals each, following the mockup if given; otherwise: title textbox, 3–4 KPI cards, 2–3 charts, 1–2 slicers on page 1; details (table/matrix) on page 2. Trigger → read `reference/visual-catalog.md` when unsure which visual or which role names to use.
   - `autoDateTable: true` when any date column exists; then use `Calendar.YearMonth` / `Calendar.Month` / `Calendar.Year` as time axes.
4. **WRITE** `output/<Name>/report-spec.json` (Name = short English PascalCase; page names / titles in English). Format in `reference/spec-schema.md` — read it the first time you write a spec in a session (`examples/app-data-spec.json` is a complete flat-table example). Keep ≤ 200 lines.
5. **BUILD** — run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec "output/<Name>/report-spec.json"`
   `ERROR n:` lines name the spec path (e.g. `pages[0].visuals[2].fields.Y[0]`) → Edit that spot, rebuild. `WARN:` lines are informational. Success prints `BUILT … (tables= measures= pages= visuals=)`.
6. **HANDOFF** (Korean, ≤ 12 lines): what was built (pages → visuals, measures), assumptions, then (also mention once: the project's `DataFolder` parameter points at this PC's `rawdata`; on another PC use **Home > Transform data > Edit parameters** to change it, or share the `.pbix`):
   > `output/<Name>/<Name>.pbip` 을 더블클릭해 Power BI Desktop에서 열고 **Refresh**(새로 고침)를 누른 뒤, **File > Save As**로 `.pbix`로 저장하세요. 수정하고 싶은 점을 말씀해 주시면 바로 반영합니다.
7. **GUIDE?** — ask: "같은 리포트를 Power BI Desktop에서 직접 만드는 방법을 정리한 가이드(manual-guide.md)도 만들어 드릴까요?" If yes, run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1 -Spec "output/<Name>/report-spec.json"`
   and reply with the path `output/<Name>/manual-guide.html` (더블클릭해 브라우저에서 열기; 단계 체크·수식 복사·배치 그림 포함; `manual-guide.md`는 같은 내용의 텍스트판). Do not read the guide.

## Iterating
"X를 바꿔줘" → Edit the spec (visual type / fields / grid / measure / page) → rebuild → one-line Korean confirmation. Remind them once that rebuilding overwrites the generated project, so Desktop-side edits should wait until they are happy.

## Quick reference
- Field ref: `Table.Name` (renamed names), aggregation prefix `sum:|avg:|min:|max:|count:|countDistinct:` on columns. Relationship keys reference *renamed* names; composite = arrays on both sides, same length/order.
- Roles: card→Values(1); charts→Category,Y[,Series]; scatter→Category,X,Y[,Size]; tableEx→Values; pivotTable→Rows,Values[,Columns]; slicer→Values(1); gauge→Y[,MinValue,MaxValue,TargetValue]; textbox→`text`.
- Grid `[col,row,w,h]` on 12×8. Title `[0,0,12,1]`, KPI `[c,1,3,2]`, chart `[c,3,6,5]`, slicer `[c,r,3,2]`.
- Types: int64 | double | decimal | text | date | datetime | boolean. csv `encoding` 65001 or 949 from the profile.
