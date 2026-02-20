use std::{fs::File, io::BufWriter};

use mr_xlsx::{csv::reader::XlsxReader, error::MrXlsxError};


fn main() -> Result<(), MrXlsxError> {
    let reader = XlsxReader::open("example.xlsx")?;
    let mut file = BufWriter::new(File::create("example.csv")?);
    reader.sheet_to_csv("Sheet (1)", &mut file)?;

    Ok(())
}