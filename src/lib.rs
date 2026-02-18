
pub mod workbook;
pub mod csv;
pub mod error;

pub(crate) type Result<T> = std::result::Result<T, error::MrXlsxError>;