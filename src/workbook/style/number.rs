#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum NumberFormat {
    General,  // (default)
    Integer,  // e.g 1,234
    Decimal2, // e.g 1,234.56
    Percent,  // e.g 12.3%
    Currency, // e.g $1,234.56
    Date,     // e.g 2024-01-31
    DateTime, // e.g 2024-01-31 14:30
    Custom(String),
}

impl NumberFormat {
    pub fn builtin_id(&self) -> Option<u32> {
        match self {
            NumberFormat::General => Some(0),
            NumberFormat::Integer => Some(1),   // "0"
            NumberFormat::Decimal2 => Some(4),  // "#,##0.00"
            NumberFormat::Percent => Some(10),  // "0.00%"
            NumberFormat::Currency => Some(7),  // "$#,##0.00"
            NumberFormat::Date => Some(14),     // "m/d/yyyy"
            NumberFormat::DateTime => Some(22), // "m/d/yyyy h:mm"
            NumberFormat::Custom(_) => None,
        }
    }
}
