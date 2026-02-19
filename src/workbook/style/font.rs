use crate::workbook::style::color::Color;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub size: u32,
    pub color: Option<Color>,
    pub name: String,
}

impl Default for Font {
    fn default() -> Self {
        Font {
            bold: false,
            italic: false,
            underline: false,
            size: 220,
            color: None,
            name: "Calibri".into(),
        }
    }
}

impl Font {
    pub fn to_xml(&self) -> String {
        let mut xml = String::from("<font>");
        if self.bold {
            xml.push_str("<b/>");
        }
        if self.italic {
            xml.push_str("<i/>");
        }
        if self.underline {
            xml.push_str("<u/>");
        }

        let pt = self.size / 20;
        xml.push_str(&format!("<sz val=\"{pt}\"/>"));
        xml.push_str(&format!("<name val=\"{}\"/>", self.name));

        if let Some(c) = &self.color {
            xml.push_str(&format!("<color rgb=\"{}\"/>", c.as_argb()));
        }

        xml.push_str("</font>");
        xml
    }
}
