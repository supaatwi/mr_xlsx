use std::collections::HashMap;

use crate::workbook::style::{
    border::{Border, BorderStyle},
    color::Color,
    fill::Fill,
    font::Font,
    number::NumberFormat,
};
pub mod border;
pub mod color;
pub mod fill;
pub mod font;
pub mod number;

#[derive(Debug, Clone)]
pub struct Style {
    pub font: Font,
    pub fill: Fill,
    pub border: Border,
    pub number_format: NumberFormat,
}

impl Default for Style {
    fn default() -> Self {
        Style {
            font: Font::default(),
            fill: Fill::None,
            border: Border::default(),
            number_format: NumberFormat::General,
        }
    }
}

impl Style {
    pub fn new() -> Self {
        Style::default()
    }

    pub fn bold(mut self) -> Self {
        self.font.bold = true;
        self
    }
    pub fn italic(mut self) -> Self {
        self.font.italic = true;
        self
    }
    pub fn underline(mut self) -> Self {
        self.font.underline = true;
        self
    }
    pub fn font_size(mut self, pt: u32) -> Self {
        self.font.size = pt * 20;
        self
    }
    pub fn font_color(mut self, hex: &str) -> Self {
        self.font.color = Some(Color::new(hex));
        self
    }
    pub fn font_name(mut self, name: &str) -> Self {
        self.font.name = name.into();
        self
    }

    pub fn bg(mut self, hex: &str) -> Self {
        self.fill = Fill::Solid(Color::new(hex));
        self
    }

    pub fn border_all(mut self, style: BorderStyle) -> Self {
        self.border.left = style.clone();
        self.border.right = style.clone();
        self.border.top = style.clone();
        self.border.bottom = style.clone();
        self
    }
    pub fn border_left(mut self, style: BorderStyle) -> Self {
        self.border.left = style;
        self
    }
    pub fn border_right(mut self, style: BorderStyle) -> Self {
        self.border.right = style;
        self
    }
    pub fn border_top(mut self, style: BorderStyle) -> Self {
        self.border.top = style;
        self
    }
    pub fn border_bottom(mut self, style: BorderStyle) -> Self {
        self.border.bottom = style;
        self
    }

    pub fn border_color(mut self, hex: &str) -> Self {
        self.border.color = Some(Color::new(hex));
        self
    }

    pub fn format(mut self, fmt: NumberFormat) -> Self {
        self.number_format = fmt;
        self
    }
    pub fn custom_format(mut self, fmt: &str) -> Self {
        self.number_format = NumberFormat::Custom(fmt.into());
        self
    }
}

pub struct StyleRegistry {
    fonts: Vec<Font>,
    fills: Vec<Fill>,
    borders: Vec<Border>,
    num_fmts: Vec<(u32, String)>,
    font_index: HashMap<Font, usize>,
    fill_index: HashMap<Fill, usize>,
    border_index: HashMap<Border, usize>,
    num_fmt_index: HashMap<String, u32>,
    xfs: Vec<(usize, usize, usize, u32)>,
    xf_index: HashMap<(usize, usize, usize, u32), usize>,

    next_num_fmt_id: u32,
}

impl StyleRegistry {
    pub fn new() -> Self {
        let mut reg = StyleRegistry {
            fonts: Vec::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            num_fmts: Vec::new(),
            font_index: HashMap::new(),
            fill_index: HashMap::new(),
            border_index: HashMap::new(),
            num_fmt_index: HashMap::new(),
            xfs: Vec::new(),
            xf_index: HashMap::new(),
            next_num_fmt_id: 164,
        };

        reg.intern_font(Font::default());
        reg.intern_fill(Fill::None);
        reg.intern_fill(Fill::None);
        reg.intern_border(Border::default());
        reg.intern_xf(0, 0, 0, 0);

        reg
    }

    pub fn register(&mut self, style: &Style) -> usize {
        let font_id = self.intern_font(style.font.clone());
        let fill_id = self.intern_fill(style.fill.clone());
        let border_id = self.intern_border(style.border.clone());
        let fmt_id = self.intern_num_fmt(&style.number_format);
        self.intern_xf(font_id, fill_id, border_id, fmt_id)
    }

    fn intern_font(&mut self, font: Font) -> usize {
        if let Some(&i) = self.font_index.get(&font) {
            return i;
        }
        let i = self.fonts.len();
        self.font_index.insert(font.clone(), i);
        self.fonts.push(font);
        i
    }

    fn intern_fill(&mut self, fill: Fill) -> usize {
        if let Some(&i) = self.fill_index.get(&fill) {
            return i;
        }
        let i = self.fills.len();
        self.fill_index.insert(fill.clone(), i);
        self.fills.push(fill);
        i
    }

    fn intern_border(&mut self, border: Border) -> usize {
        if let Some(&i) = self.border_index.get(&border) {
            return i;
        }
        let i = self.borders.len();
        self.border_index.insert(border.clone(), i);
        self.borders.push(border);
        i
    }

    fn intern_num_fmt(&mut self, fmt: &NumberFormat) -> u32 {
        if let Some(id) = fmt.builtin_id() {
            return id;
        }
        if let NumberFormat::Custom(code) = fmt {
            if let Some(&id) = self.num_fmt_index.get(code) {
                return id;
            }
            let id = self.next_num_fmt_id;
            self.next_num_fmt_id += 1;
            self.num_fmt_index.insert(code.clone(), id);
            self.num_fmts.push((id, code.clone()));
            return id;
        }
        0
    }

    fn intern_xf(
        &mut self,
        font_id: usize,
        fill_id: usize,
        border_id: usize,
        num_fmt_id: u32,
    ) -> usize {
        let key = (font_id, fill_id, border_id, num_fmt_id);
        if let Some(&i) = self.xf_index.get(&key) {
            return i;
        }
        let i = self.xfs.len();
        self.xf_index.insert(key, i);
        self.xfs.push(key);
        i
    }

    pub fn to_xml(&self) -> String {
        let mut out = String::new();
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        out.push_str(
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        );

        if self.num_fmts.is_empty() {
            out.push_str("<numFmts count=\"0\"/>\n");
        } else {
            out.push_str(&format!("<numFmts count=\"{}\">\n", self.num_fmts.len()));
            for (id, code) in &self.num_fmts {
                out.push_str(&format!(
                    "<numFmt numFmtId=\"{id}\" formatCode=\"{code}\"/>\n"
                ));
            }
            out.push_str("</numFmts>\n");
        }

        out.push_str(&format!("<fonts count=\"{}\">\n", self.fonts.len()));
        for font in &self.fonts {
            out.push_str(&format!("{}\n", font.to_xml()));
        }
        out.push_str("</fonts>\n");

        out.push_str(&format!("<fills count=\"{}\">\n", self.fills.len()));
        for (i, fill) in self.fills.iter().enumerate() {
            if i == 1 {
                out.push_str("<fill><patternFill patternType=\"gray125\"/></fill>\n");
            } else {
                out.push_str(&format!("{}\n", fill.to_xml()));
            }
        }
        out.push_str("</fills>\n");

        out.push_str(&format!("<borders count=\"{}\">\n", self.borders.len()));
        for border in &self.borders {
            out.push_str(&format!("{}\n", border.to_xml()));
        }
        out.push_str("</borders>\n");

        out.push_str("<cellStyleXfs count=\"1\">\n");
        out.push_str("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
        out.push_str("</cellStyleXfs>\n");

        out.push_str(&format!("<cellXfs count=\"{}\">\n", self.xfs.len()));
        for (font_id, fill_id, border_id, num_fmt_id) in &self.xfs {
            out.push_str(&format!("<xf numFmtId=\"{num_fmt_id}\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"{border_id}\" xfId=\"0\"/>\n"));
        }
        out.push_str("</cellXfs>\n");

        out.push_str("<cellStyles count=\"1\">\n");
        out.push_str("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n");
        out.push_str("</cellStyles>\n");

        out.push_str("</styleSheet>");
        out
    }
}
