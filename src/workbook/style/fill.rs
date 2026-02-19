use crate::workbook::style::color::Color;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Fill {
    None,
    Solid(Color),
}

impl Fill {
    pub fn to_xml(&self) -> String {
        match self {
            Fill::None => "<fill><patternFill/></fill>".into(),
            Fill::Solid(c) => format!(
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/></patternFill></fill>",
                c.as_argb()
            ),
        }
    }
}
