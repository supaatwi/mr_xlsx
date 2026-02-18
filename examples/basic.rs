use mr_xlsx::error::MrXlsxError;
use mr_xlsx::workbook::builder::WorkbookBuilder;
use mr_xlsx::workbook::cell::CellValue;

fn main() -> Result<(), MrXlsxError> {
    let mut wb = WorkbookBuilder::new("./example.xlsx")
        .set_sheets(vec!["Summary", "AA-001"])
        .build()?;

    let sheet1 = wb.add_sheet("Summary")?;
    sheet1.write_row(&[
        CellValue::text("Name"),
        CellValue::text("Score"),
        CellValue::text("Pass"),
    ])?;

    for i in 1..1_000_000 {
        sheet1.write_row(&[
            CellValue::text(&format!("A{i}")),
            CellValue::text(&format!("{}", i)),
            CellValue::text("Pass"),
        ])?;
        
        if i % 10000 == 0 {
            println!("write row 10000 row");
            std::thread::sleep(std::time::Duration::from_secs(2));
        }
    }

    wb.finish()?;
    Ok(())
}
