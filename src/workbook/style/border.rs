use crate::workbook::style::color::Color;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum BorderStyle {
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
}

impl BorderStyle {
    fn as_xml_attr(&self) -> Option<&str> {
        match self {
            BorderStyle::None => None,
            BorderStyle::Thin => Some("thin"),
            BorderStyle::Medium => Some("medium"),
            BorderStyle::Thick => Some("thick"),
            BorderStyle::Dashed => Some("dashed"),
            BorderStyle::Dotted => Some("dotted"),
            BorderStyle::Double => Some("double"),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Border {
    pub left: BorderStyle,
    pub right: BorderStyle,
    pub top: BorderStyle,
    pub bottom: BorderStyle,
    pub color: Option<Color>,
}

impl Default for Border {
    fn default() -> Self {
        Border {
            left: BorderStyle::None,
            right: BorderStyle::None,
            top: BorderStyle::None,
            bottom: BorderStyle::None,
            color: None,
        }
    }
}

impl Border {
    fn side_xml(&self, tag: &str, style: &BorderStyle) -> String {
        match style.as_xml_attr() {
            None => format!("<{tag}/>"),
            Some(s) => {
                let color = self
                    .color
                    .as_ref()
                    .map(|c| format!("<color rgb=\"{}\"/>", c.as_argb()))
                    .unwrap_or_default();
                format!("<{tag} style=\"{s}\">{color}</{tag}>")
            }
        }
    }

    pub fn to_xml(&self) -> String {
        format!(
            "<border>{}{}{}{}<diagonal/></border>",
            self.side_xml("left", &self.left),
            self.side_xml("right", &self.right),
            self.side_xml("top", &self.top),
            self.side_xml("bottom", &self.bottom),
        )
    }
}
