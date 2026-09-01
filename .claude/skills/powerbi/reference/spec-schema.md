# report-spec.json — contract

One JSON file per report at `output/<Name>/report-spec.json`. The builder validates it and emits the PBIP. Keys marked * are required.

```jsonc
{
  "name*": "SalesDashboard",        // letters/digits/_/- ; becomes the .pbip name
  "locale": "ko-KR",                // default en-US
  "autoDateTable": true,            // add a Calendar table for date columns not linked to a date dimension
  "tables*": [{
    "name*": "Sales",               // model table name; no "."; "Calendar" is reserved
    "file": "sales.csv",            // file inside rawdata/ — OR:
    "filePattern": "sales_*.csv",   // several files with identical columns → one table; exactly one *; new files load on Refresh
    "sheet": "Data",                // xlsx only (required for xlsx)
    "headerRow": 1,                 // from the profile: header_row
    "encoding": 65001,              // csv only: 65001 or 949, from the profile
    "delimiter": ",",               // csv only
    "clean": {                      // table-level cleaning (all default off)
      "removeBlankRows": true,      //   drop rows where every loaded column is empty (trailer/footer lines)
      "dropDuplicates": true,       //   drop EXACT duplicate rows only (safe: identical double-entries)
      "errorsToNull": true,         //   any remaining conversion error becomes null instead of blocking Refresh - safe default for big fact tables
      "dedupeBy": {                 //   same key, DIFFERENT rows -> resolve by rule instead of blind deletion:
        "keys": ["Order ID"],       //     renamed column(s) that should be unique
        "keep": "last",             //     first | last | mostComplete (row with fewest nulls wins)
        "orderBy": "Updated Date"   //     required for first/last: the column that decides which row wins
      }
    },
    "derived": [                    // new columns computed AFTER cleaning/typing (M expression over renamed columns)
      { "name": "Total Guests", "expr": "[Adults] + [Children] + [Babies]", "type": "int64", "format": "#,0", "summarizeBy": "sum", "hidden": false }
    ],
    "columns*": [{                  // ONLY listed columns are loaded; list the ones the report needs
      "name*": "track_name",        // source header exactly as in the profile
      "rename": "App Name",         // English model name (translate Korean headers)
      "type*": "text",              // int64 | double | decimal | text | date | datetime | boolean
      "key": true,                  // unique id column (also sets summarizeBy none)
      "hidden": true,               // hide keys / technical columns
      "format": "#,0",              // numeric format string
      "summarizeBy": "average",     // none|sum|average|count|min|max|distinctCount (default: sum for numbers, else none)
      "description": "Track Name",  // from the dictionary sheet if any
      "sortBy": "MonthNo",          // sort this column by another column of the same table
      "trim": true,                 // clean: strip leading/trailing whitespace (before typing)
      "case": "upper",              // clean: upper|lower - normalize codes/keys ("prt" -> "PRT")
      "nullValues": ["TBD", "N/A"], // clean: these exact values (after trim/case) become null
      "dateFormats": ["dd/MM/yyyy", "MM-dd-yyyy"],  // date/datetime only: extra parse formats tried after the default; unparseable values become null
      "valueMap": { "BB": "Bed & Breakfast", "HB": "Half Board" },  // clean: code -> label (unmapped values pass through)
      "numberClean": true,          // clean: "₩1,234", "(500)", "12%" -> 1234, -500, 0.12 (unparseable -> null)
      "fillDown": true,             // clean: empty cells inherit the value above (merged-cell style exports)
      "validRange": [0, 20]         // numeric only: values outside [min,max] become null (impossible values, e.g. negative qty)
    }]
  },
  {
    "name": "MonthlySummary",       // summary table: granularity change via Group By (no file/columns/clean here -
    "from": "Sales",                //   it inherits the SOURCE table's fully cleaned pipeline, then groups it)
    "groupBy": ["Region"],          // renamed source columns to group on
    "aggregations": [               // agg: count | countDistinct | sum | average | min | max ("column" required except count)
      { "name": "Order Count", "agg": "count" },
      { "name": "Total Amount", "agg": "sum", "column": "Amount", "format": "#,0" }
    ]
  }],
  "relationships": [
    { "from*": "Sales.CityKey", "to*": "City.CityKey", "active": true, "crossFilter": "single" },   // names may differ; use the profile's KEY MATCHES
    { "from": ["Sales.Month", "Sales.Region Code"], "to": ["Targets.Month", "Targets.Region Code"] }   // composite key: arrays, same length/order; builder adds hidden "_key_Month_Region Code" columns
  ],
  "measures": [{ "table*": "Sales", "name*": "Revenue", "dax*": "SUM(Sales[Amount])", "format": "#,0", "description": "" }],
  "pages*": [{
    "name*": "Overview",
    "filters": [{ "field": "Sales.Region", "in": ["Seoul", "Busan"] }],
    "visuals*": [{
      "type*": "card",              // see visual-catalog.md
      "title": "Total Revenue",
      "fields": { "Values": ["Sales.Revenue"] },   // role → list of field refs
      "grid*": [0, 0, 3, 2],        // [col,row,w,h] on a 12x8 grid (1920x1080 page)
      "sort": { "field": "Sales.Revenue", "direction": "desc" },
      "topN": { "n": 10, "by": "Sales.Revenue" },   // needs a Category (or Rows) field
      "filters": [{ "field": "Sales.Region", "in": ["Seoul"] }],
      "text": "Title", "fontSize": 24               // textbox only
    }]
  }]
}
```

**Field refs:** `Table.Name` — the builder checks measures first, then columns (use the *renamed* name). Aggregate a raw column with a prefix: `sum:` `avg:` `min:` `max:` `count:` `countDistinct:` (columns only, never measures). Calendar fields when `autoDateTable` produced one: `Calendar.Date`, `Calendar.Year`, `Calendar.Quarter`, `Calendar.Month`, `Calendar.YearMonth`.

**Cleaning:** the profiler flags dirt per column (`DIRTY: n non-date ...`, `placeholders:`, `whitespace:`, `case-variants:`) and per table (`DIRT: ... blank/duplicate row(s)`). Map them 1:1 to the options above. Relationship columns must be cleaned on **both** sides (`trim` + `case` on the fact column AND the lookup key) or the join misses rows. Order of operations in the built query: removeBlankRows → rename → trim/case/nullValues/valueMap/numberClean → fillDown → dateFormats → type conversion → errorsToNull → validRange → dropDuplicates → dedupeBy → derived → (groupBy).

**Duplicates - resolve logically, never blindly:** `dropDuplicates` only removes rows that are identical in every loaded column (double-entries; always safe). When the profiler shows `CONFLICTING dup ids` (same id, different data), pick a rule with `dedupeBy`: `keep: "last"` + `orderBy: <update/status date>` when the newest row wins, `keep: "mostComplete"` when the fullest row wins - and load the `orderBy` column so the rule can run. State the chosen rule in the handoff so the user knows which rows were dropped.

**Grid math:** 12 columns × 8 rows, 40 px margin, 20 px gutter. Common shapes: KPI card `[c,r,3,2]` (445 × 235 px), half-width chart `[c,r,6,4]`, full-width chart `[0,r,12,4]`, slicer `[c,r,3,2]`, title textbox `[0,0,12,1]`.

**DAX inside `dax`:** reference model names: `Sales[매출액]`, `'Table with space'[Col]`, `[OtherMeasure]`. Use `\n` for multi-line.
