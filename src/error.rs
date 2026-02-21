use std::fmt;

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


impl fmt::Display for MrXlsxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            MrXlsxError::Io(e) => write!(f, "IO error: {e}"),
			MrXlsxError::AlreadyExists(e) => write!(f, "Already Exists Sheet : {e}"),
			MrXlsxError::NotFound(e) => write!(f, "Mr Xlsx Not Found : {e}"),
			MrXlsxError::ZipError(e) => write!(f, "Zip Error : {e}")
        }
    }
}

impl std::error::Error for MrXlsxError {}