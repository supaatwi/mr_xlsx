# mr_xlsx

A Rust library for reading and writing `.xlsx` files with a focus on **streaming** and **low memory usage**.

Designed for workloads where memory matters — writing millions of rows or reading large files without loading everything into RAM.

---

## Features

- **Streaming writer** — rows are written to a temp file immediately, never buffered in memory
- **Streaming reader** — parse sheet rows one at a time via an iterator
- **Sheet reordering** — write sheets in any order, set the final tab order at `finish()`
- **Style support** — font, fill, border, number format via a builder API
- **xlsx → CSV** — convert any sheet to CSV row by row
- **Style deduplication** — identical styles are interned automatically, no bloat in `styles.xml`

---

## Writing

### Basic

```rust
use mr_xlsx::{Workbook, CellValue};

let mut wb = Workbook::new("output.xlsx");
let sheet = wb.add_sheet("Sales")?;

sheet.write_row(&[
    CellValue::text("Name"),
    CellValue::text("Score"),
    CellValue::bool(true),
], None)?;

sheet.write_row(&[
    CellValue::text("Alice"),
    CellValue::num(95.0),
    CellValue::bool(true),
], None)?;

wb.finish()?;
```

### Cell types

| Method | Excel type | Example |
|---|---|---|
| `CellValue::text("hello")` | Inline string | `"hello"` |
| `CellValue::num(42.0)` | Number | `42` |
| `CellValue::bool(true)` | Boolean | `TRUE` |
| `CellValue::formula("SUM(A1:A10)")` | Formula | `=SUM(A1:A10)` |
| `CellValue::blank()` | Empty cell | |

### Multiple sheets with custom tab order

```rust
let mut wb = Workbook::new("report.xlsx");

let data  = wb.add_sheet("Data")?;
// ... write rows to data ...

let summary = wb.add_sheet("Summary")?;
// ... write rows to summary ...

// Summary tab appears first, regardless of write order
wb.finish_ordered(&["Summary", "Data"])?;
```

---

## Styling

Styles use a builder pattern. Pass a `Style` as the second argument to `write_row()` — all cells in that row share the same style.

```rust
use mr_xlsx::{Style, BorderStyle, NumberFormat};

let header = Style::new()
    .bold()
    .font_color("FFFFFF")
    .font_size(12)
    .bg("4472C4")
    .border_all(BorderStyle::Thin);

let money = Style::new()
    .format(NumberFormat::Currency);

let date = Style::new()
    .format(NumberFormat::Date)
    .italic();

sheet.write_row(&[CellValue::text("Name"), CellValue::text("Salary")], Some(&header))?;
sheet.write_row(&[CellValue::text("Alice"), CellValue::num(85000.0)], Some(&money))?;
```

### Available style options

**Font**
```rust
Style::new()
    .bold()
    .italic()
    .underline()
    .font_size(14)          // pt
    .font_color("FF0000")   // RGB hex
    .font_name("Arial")
```

**Fill**
```rust
Style::new()
    .bg("4472C4")           // RGB hex background
```

**Border**
```rust
Style::new()
    .border_all(BorderStyle::Thin)
    .border_bottom(BorderStyle::Thick)
    .border_color("CCCCCC")
```

Border styles: `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`

**Number format**
```rust
Style::new().format(NumberFormat::Currency)
Style::new().format(NumberFormat::Date)
Style::new().format(NumberFormat::Percent)
Style::new().custom_format("#,##0.000")
```

Built-in formats: `General`, `Integer`, `Decimal2`, `Percent`, `Currency`, `Date`, `DateTime`

---

## Reading

### Stream rows from a sheet

```rust
use mr_xlsx::XlsxReader;

let reader = XlsxReader::open("input.xlsx")?;

println!("{:?}", reader.sheet_names()); // ["Sheet1", "Sheet2"]

for row in reader.stream_rows("Sheet1")? {
    let row = row?;
    println!("{:?}", row.cells);
    // each row is dropped before the next is read
}
```

Memory at any point: `sharedStrings` + current row. Not proportional to the number of rows.

### Convert to CSV

```rust
// Single sheet to a writer
let mut file = BufWriter::new(File::create("output.csv")?);
reader.sheet_to_csv("Sheet1", &mut file)?;

// All sheets to separate files: output_Sheet1.csv, output_Sheet2.csv, ...
reader.all_sheets_to_csv("output")?;
```

---

## Memory model

```
Writing
  write_row()  →  temp file on disk     (O(1) RAM per row)
  finish()     →  zip temp files        (O(1) RAM)

Reading
  open()       →  load sharedStrings    (O(unique strings))
  stream_rows()→  load sheet XML        (O(XML bytes))
  next()       →  parse one row         (O(columns))
```

`sharedStrings` must be fully loaded because cells reference strings by index (random access). Everything else is streamed.

---