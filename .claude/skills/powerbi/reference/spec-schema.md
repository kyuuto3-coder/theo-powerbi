# report-spec.json — contract

One JSON file per report at `output/<Name>/report-spec.json`. The builder validates it and emits the PBIP. Keys marked * are required.

```jsonc
{
  "name*": "SalesDashboard",        // letters/digits/_/- ; becomes the .pbip name
  "locale": "ko-KR",                // default en-US
  "autoDateTable": true,            // add a Calendar table for date columns not linked to a date dimension
  "tables*": [{
    "name*": "Sales",               // model table name; no "."; "Calendar" is reserved
    "file*": "sales.csv",           // file inside rawdata/
    "sheet": "Data",                // xlsx only (required for xlsx)
    "headerRow": 1,                 // from the profile: header_row
    "encoding": 65001,              // csv only: 65001 or 949, from the profile
    "delimiter": ",",               // csv only
    "columns*": [{                  // ONLY listed columns are loaded; list the ones the report needs
      "name*": "track_name",        // source header exactly as in the profile
      "rename": "앱 이름",           // model name (optional)
      "type*": "text",              // int64 | double | decimal | text | date | datetime | boolean
      "key": true,                  // unique id column (also sets summarizeBy none)
      "hidden": true,               // hide keys / technical columns
      "format": "#,0",              // numeric format string
      "summarizeBy": "average",     // none|sum|average|count|min|max|distinctCount (default: sum for numbers, else none)
      "description": "Track Name",  // from the dictionary sheet if any
      "sortBy": "MonthNo"           // sort this column by another column of the same table
    }]
  }],
  "relationships": [{ "from*": "Sales.CityKey", "to*": "City.CityKey", "active": true, "crossFilter": "single" }],
  "measures": [{ "table*": "Sales", "name*": "매출", "dax*": "SUM(Sales[Amount])", "format": "#,0", "description": "" }],
  "pages*": [{
    "name*": "개요",
    "filters": [{ "field": "Sales.Region", "in": ["Seoul", "Busan"] }],
    "visuals*": [{
      "type*": "card",              // see visual-catalog.md
      "title": "총 매출",
      "fields": { "Values": ["Sales.매출"] },   // role → list of field refs
      "grid*": [0, 0, 3, 2],        // [col,row,w,h] on a 12x8 grid (1920x1080 page)
      "sort": { "field": "Sales.매출", "direction": "desc" },
      "topN": { "n": 10, "by": "Sales.매출" },   // needs a Category (or Rows) field
      "filters": [{ "field": "Sales.Region", "in": ["Seoul"] }],
      "text": "제목", "fontSize": 24                // textbox only
    }]
  }]
}
```

**Field refs:** `Table.Name` — the builder checks measures first, then columns (use the *renamed* name). Aggregate a raw column with a prefix: `sum:` `avg:` `min:` `max:` `count:` `countDistinct:` (columns only, never measures). Calendar fields when `autoDateTable` produced one: `Calendar.Date`, `Calendar.Year`, `Calendar.Quarter`, `Calendar.Month`, `Calendar.YearMonth`.

**Grid math:** 12 columns × 8 rows, 40 px margin, 20 px gutter. Common shapes: KPI card `[c,r,3,2]` (445 × 235 px), half-width chart `[c,r,6,4]`, full-width chart `[0,r,12,4]`, slicer `[c,r,3,2]`, title textbox `[0,0,12,1]`.

**DAX inside `dax`:** reference model names: `Sales[매출액]`, `'Table with space'[Col]`, `[OtherMeasure]`. Use `\n` for multi-line.
