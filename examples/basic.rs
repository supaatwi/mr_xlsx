use mr_xlsx::error::MrXlsxError;
use mr_xlsx::workbook::builder::WorkbookBuilder;
use mr_xlsx::workbook::cell::CellValue;

fn main() -> Result<(), MrXlsxError> {
    let mut wb = WorkbookBuilder::new("./example.xlsx")
        .set_sheets(vec!["Summary"])
        .build()?;

    for x in 1..10 {
        let sheet_name = format!("Sheet ({})", x);
        let sheet = wb.add_sheet(&sheet_name)?;
        sheet.write_row(&[
            CellValue::text("Name"),
            CellValue::text("Score"),
            CellValue::text("Pass"),
        ])?;
        for i in 1..1_000_000 {
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

    match wb.get_sheet("Summary") {
        Some(sheet) => {
             sheet.write_row(&[
                CellValue::text("Summary Name"),
                CellValue::text("Summary Score"),
                CellValue::text("Summary Pass"),
            ])?;
        }
        None => {

        }
    }
    

    wb.finish()?;
    Ok(())
}
