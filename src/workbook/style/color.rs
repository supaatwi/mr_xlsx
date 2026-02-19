#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Color(pub String);

impl Color {
    pub fn new(hex: &str) -> Self {
        let hex = hex.trim_start_matches('#');
        if hex.len() == 6 {
            Color(format!("FF{}", hex.to_uppercase()))
        } else {
            Color(hex.to_uppercase())
        }
    }
    pub(crate) fn as_argb(&self) -> &str {
        &self.0
    }
}
