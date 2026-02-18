use std::io::{BufWriter, Write};

use tempfile::NamedTempFile;

use crate::{
    Result,
    workbook::{cell::CellValue, make_cell_ref, write_cell},
};

pub struct SheetWriter {
    name: String,
    pub(crate) temp: BufWriter<NamedTempFile>,
    current_row: u32,
    max_col: u32,
}

impl SheetWriter {
    pub(crate) fn new(name: &str) -> Result<Self> {
        let temp_file = NamedTempFile::new()?;
        let mut writer = BufWriter::new(temp_file);

        write!(
            writer,
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                    <sheetViews>
                        <sheetView workbookViewId="0"/></sheetViews>
                    <sheetFormatPr defaultRowHeight="15"/>
                    <sheetData>
            "#
        )?;

        Ok(SheetWriter {
            name: name.to_string(),
            temp: writer,
            current_row: 0,
            max_col: 0,
        })
    }

    pub fn get_name(&self) -> String {
        self.name.clone()
    }

    pub fn write_row(&mut self, cells: &[CellValue]) -> Result<()> {
        self.current_row += 1;
        let row = self.current_row;

        if cells.is_empty() {
            return Ok(());
        }

        if cells.len() as u32 > self.max_col {
            self.max_col = cells.len() as u32;
        }

        write!(self.temp, "<row r=\"{row}\">")?;

        for (col_idx, cell) in cells.iter().enumerate() {
            let col = col_idx as u32; // 0-based
            let cell_ref = make_cell_ref(row, col); // e.g "A1", "B2"
            write_cell(&mut self.temp, &cell_ref, cell)?;
        }

        writeln!(self.temp, "</row>")?;

        Ok(())
    }

    pub(crate) fn finalize(&mut self) -> Result<()> {
        write!(
            self.temp,
            "</sheetData>\
                <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>\
             </worksheet>"
        )?;
        self.temp.flush()?;
        Ok(())
    }
}
