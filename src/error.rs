use zip::result::ZipError;



#[derive(Debug)]
pub enum MrXlsxError {
    AlreadyExists(String),
    NotFound(String),
    Io(String),
    ZipError(String)
}

impl From<std::io::Error> for MrXlsxError {
	fn from(e: std::io::Error) -> MrXlsxError {
		MrXlsxError::Io(e.to_string())
	}
}

impl From<ZipError> for MrXlsxError {
	fn from(e: ZipError) -> MrXlsxError {
		MrXlsxError::ZipError(e.to_string())
	}
}