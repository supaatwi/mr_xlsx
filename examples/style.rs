use mr_xlsx::{error::MrXlsxError, workbook::{builder::WorkbookBuilder, cell::CellValue, style::Style}};



fn main() -> Result<(), MrXlsxError> {

    let header = Style::new()
    .bold()
    .bg("2784F5")
    .font_color("000000");

    let mut wb = WorkbookBuilder::new("./example.xlsx").build()?;

    for x in 1..10 {
        let sheet_name = format!("Sheet ({})", x);
        let sheet = wb.add_sheet(&sheet_name)?;
        sheet.write_row_with_style(&[
            CellValue::text("Name"),
            CellValue::text("Score"),
            CellValue::text("Pass"),
        ], &header)?;
        for i in 1..1_000 {
            sheet.write_row(&[
                CellValue::text(&format!("A{i}")),
                CellValue::text(&format!("{}", i)),
                CellValue::text("Pass"),
            ])?;
            
            if i % 100_000 == 0 {
                std::thread::sleep(std::time::Duration::from_secs(1));
            }
        }
        println!("Done sheet : {}", sheet_name);
    }

    wb.finish()?;
    Ok(())
}