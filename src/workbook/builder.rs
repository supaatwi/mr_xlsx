use crate::{Result, workbook::Workbook};

pub struct WorkbookBuilder<T: Into<String>> {
    path: T,
    sheets: Vec<T>,
}

impl<T> WorkbookBuilder<T>
where
    T: Into<String>,
{
    pub fn new(path: T) -> Self {
        Self {
            path,
            sheets: vec![],
        }
    }

    pub fn set_sheets(mut self, sheets: Vec<T>) -> Self {
        self.sheets = sheets;
        self
    }

    pub fn build(self) -> Result<Workbook> {
        Ok(Workbook::new_with_builder(
            self.path.into(),
            self.sheets.into_iter().map(|s| s.into()).collect(),
        )?)
    }
}
