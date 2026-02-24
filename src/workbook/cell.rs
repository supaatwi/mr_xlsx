pub enum CellValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Formula(String),
    Blank,
}

impl CellValue {
    pub fn num(v: f64) -> Self {
        CellValue::Number(v)
    }
    pub fn text<S: Into<String>>(v: S) -> Self {
        CellValue::Text(v.into())
    }
    pub fn bool(v: bool) -> Self {
        CellValue::Bool(v)
    }
    pub fn formula<S: Into<String>>(v: S) -> Self {
        CellValue::Formula(v.into())
    }
}
