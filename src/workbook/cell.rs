/// ประเภทค่าที่ cell รองรับ
pub enum CellValue<'a> {
    /// ตัวเลข (integer, float รวมกันหมด — Excel เก็บเป็น f64 ทั้งนั้น)
    Number(f64),
    /// String — ใช้ inline string ไม่ผ่าน sharedStrings เพื่อ streaming ง่าย
    Text(&'a str),
    /// Boolean
    Bool(bool),
    /// Formula เช่น "SUM(A1:A10)" — ไม่ต้องใส่ = นำหน้า
    Formula(&'a str),
    /// Cell ว่างเปล่า
    Blank,
}

// shorthand constructors เพื่อให้เขียนสั้นลง
impl<'a> CellValue<'a> {
    pub fn num(v: f64)     -> Self { CellValue::Number(v) }
    pub fn text(v: &'a str) -> Self { CellValue::Text(v) }
    pub fn bool(v: bool)   -> Self { CellValue::Bool(v) }
    pub fn formula(v: &'a str) -> Self { CellValue::Formula(v) }
}