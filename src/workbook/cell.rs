pub enum CellValue<'a> {
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Formula(&'a str),
    Blank,
}

impl<'a> CellValue<'a> {
    pub fn num(v: f64)     -> Self { CellValue::Number(v) }
    pub fn text(v: &'a str) -> Self { CellValue::Text(v) }
    pub fn bool(v: bool)   -> Self { CellValue::Bool(v) }
    pub fn formula(v: &'a str) -> Self { CellValue::Formula(v) }
}