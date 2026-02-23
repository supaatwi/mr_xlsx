use mr_xlsx::error::MrXlsxError;
use mr_xlsx::workbook::builder::WorkbookBuilder;
use mr_xlsx::workbook::cell::CellValue;

fn main() -> Result<(), MrXlsxError> {
    let mut wb = WorkbookBuilder::new("./example.xlsx")
        .build()?;

    for x in 1..10 {
        let sheet_name = format!("Sheet ({})", x);
        let sheet = wb.add_sheet(&sheet_name)?;
        sheet.write_row(&[
            CellValue::text("Name"),
            CellValue::text("Score"),
            CellValue::text("Pass"),
        ])?;
        for i in 1..1_000 {
            sheet.write_row(&[
                CellValue::text(&format!("A{i}")),
                CellValue::text(&format!("{}", i)),
                CellValue::text("Pass"),
            ])?;
           
        }
        println!("Done sheet : {}", sheet_name);
    }

    let sheet = wb.add_sheet("Summary")?;
    sheet.write_row(&[
            CellValue::text("Name"),
            CellValue::text("Score"),
            CellValue::text("Pass"),
    ])?;
    

    wb.finish_by_order(&["Summary"])?;
    Ok(())
}
