use std::{
    collections::HashMap,
    fs::File,
    io::{Seek, SeekFrom, Write},
};

use zip::{ZipWriter, write::SimpleFileOptions};

use crate::{
    Result,
    error::MrXlsxError,
    workbook::{cell::CellValue, sheet::SheetWriter, style::StyleRegistry},
};
pub mod builder;
pub mod cell;
pub mod sheet;
pub mod style;

const RELS_DOT_RELS: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
    r#"</Relationships>"#,
);

fn workbook_xml(order: &[String]) -> String {
    let mut sheets = String::new();
    for (i, name) in order.iter().enumerate() {
        let sheet_id = i + 1;
        let r_id = format!("rId{}", i + 1);
        let escaped_name = xml_escape(name);
        sheets.push_str(&format!(
            r#"<sheet name="{escaped_name}" sheetId="{sheet_id}" r:id="{r_id}"/>"#
        ));
    }

    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<bookViews><workbookView activeTab="0"/></bookViews>"#,
            r#"<sheets>{}</sheets>"#,
            r#"<calcPr fullCalcOnLoad="1"/>"#,
            r#"</workbook>"#,
        ),
        sheets
    )
}

fn workbook_rels_xml(sheet_count: usize) -> String {
    let mut rels = String::new();

    for i in 1..=sheet_count {
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        ));
    }

    let styles_id = sheet_count + 1;
    rels.push_str(&format!(
        r#"<Relationship Id="rId{styles_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#
    ));

    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            r#"{}"#,
            r#"</Relationships>"#,
        ),
        rels
    )
}

fn content_types_xml(sheet_count: usize) -> String {
    let mut overrides = String::new();

    for i in 1..=sheet_count {
        overrides.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
    }

    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
            r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
            r#"<Default Extension="xml" ContentType="application/xml"/>"#,
            r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
            r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
            r#"{}"#,
            r#"</Types>"#,
        ),
        overrides
    )
}

pub struct Workbook {
    output_path: String,
    sheets: HashMap<String, SheetWriter>,
    insertion_order: Vec<String>,
    style_reg: Box<StyleRegistry>,
}

impl Workbook {
    pub(crate) fn new_with_builder(path: String, sheets: Vec<String>) -> Result<Self> {
        let mut insertion_order = vec![];
        let mut _sheets = HashMap::new();
        let mut style_reg = Box::new(StyleRegistry::new());
        let reg_ptr: *mut StyleRegistry = &mut *style_reg;

        sheets.into_iter().try_for_each(|name| -> Result<()> {
            let sheet_writer = SheetWriter::new(&name, reg_ptr)?;
            insertion_order.push(name.clone());
            _sheets.insert(name, sheet_writer);
            Ok(())
        })?;

        Ok(Self {
            output_path: path,
            sheets: _sheets,
            insertion_order,
            style_reg,
        })
    }

    pub fn get_sheet(&mut self, name: &str) -> Option<&mut SheetWriter> {
        self.sheets.get_mut(name)
    }

    pub fn add_sheet(&mut self, name: &str) -> Result<&mut SheetWriter> {
        if self.sheets.contains_key(name) {
            return Err(MrXlsxError::AlreadyExists(format!(
                "Sheet '{name}' already exists"
            )));
        }
        let reg_ptr: *mut StyleRegistry = &mut *self.style_reg;
        let writer = SheetWriter::new(name, reg_ptr)?;
        self.sheets.insert(name.to_string(), writer);
        self.insertion_order.push(name.to_string());
        let sheet = match self.sheets.get_mut(name) {
            Some(s) => s,
            None => return Err(MrXlsxError::NotFound(format!("Sheet {name} not found!!"))),
        };
        Ok(sheet)
    }

    pub fn finish_by_order(mut self, sheet_order: &[&str]) -> Result<()> {
        let _sheet_order = sheet_order
            .into_iter()
            .map(|s| s.to_string())
            .collect::<Vec<_>>();

        let mut insertion_order = vec![];
        insertion_order.extend(
            sheet_order
                .into_iter()
                .map(|s| s.to_string())
                .collect::<Vec<String>>(),
        );
        self.insertion_order.into_iter().for_each(|sheet| {
            if !insertion_order.contains(&sheet) {
                insertion_order.push(sheet);
            }
        });

         for name in &insertion_order {
            match self.sheets.get_mut(name) {
                Some(s) => s.finalize()?,
                None => {
                    return Err(MrXlsxError::NotFound(format!("Sheet name : {name}!!")));
                }
            }
        }
       
       let output_file = File::create(&self.output_path)?;
        let mut zip = ZipWriter::new(output_file);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        zip_write_str(
            &mut zip,
            "[Content_Types].xml",
            &content_types_xml(insertion_order.len()),
            options,
        )?;
        zip_write_str(&mut zip, "_rels/.rels", RELS_DOT_RELS, options)?;
        zip_write_str(
            &mut zip,
            "xl/workbook.xml",
            &workbook_xml(&insertion_order),
            options,
        )?;
        zip_write_str(
            &mut zip,
            "xl/_rels/workbook.xml.rels",
            &workbook_rels_xml(insertion_order.len()),
            options,
        )?;

        let styles_xml = self.style_reg.to_xml();
        zip_write_str(&mut zip, "xl/styles.xml", &styles_xml, options)?;

        for (i, name) in insertion_order.iter().enumerate() {
            let sheet = self.sheets.get_mut(name).unwrap();
            let zip_path = format!("xl/worksheets/sheet{}.xml", i + 1);

            zip.start_file(&zip_path, options)?;

            let temp_file = sheet.temp.get_mut();
            temp_file.seek(SeekFrom::Start(0))?;

            let mut buf = [0u8; 64 * 1024];
            loop {
                use std::io::Read;
                let n = temp_file.read(&mut buf)?;
                if n == 0 {
                    break;
                }
                zip.write_all(&buf[..n])?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    pub fn finish(mut self) -> Result<()> {
        for name in &self.insertion_order {
            match self.sheets.get_mut(name) {
                Some(s) => s.finalize()?,
                None => {
                    return Err(MrXlsxError::NotFound(format!("Sheet name : {name}!!")));
                }
            }
        }

        let output_file = File::create(&self.output_path)?;
        let mut zip = ZipWriter::new(output_file);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        zip_write_str(
            &mut zip,
            "[Content_Types].xml",
            &content_types_xml(self.insertion_order.len()),
            options,
        )?;
        zip_write_str(&mut zip, "_rels/.rels", RELS_DOT_RELS, options)?;
        zip_write_str(
            &mut zip,
            "xl/workbook.xml",
            &workbook_xml(&self.insertion_order),
            options,
        )?;
        zip_write_str(
            &mut zip,
            "xl/_rels/workbook.xml.rels",
            &workbook_rels_xml(self.insertion_order.len()),
            options,
        )?;

        let styles_xml = self.style_reg.to_xml();
        zip_write_str(&mut zip, "xl/styles.xml", &styles_xml, options)?;

        for (i, name) in self.insertion_order.iter().enumerate() {
            let sheet = self.sheets.get_mut(name).unwrap();
            let zip_path = format!("xl/worksheets/sheet{}.xml", i + 1);

            zip.start_file(&zip_path, options)?;

            let temp_file = sheet.temp.get_mut();
            temp_file.seek(SeekFrom::Start(0))?;

            let mut buf = [0u8; 64 * 1024];
            loop {
                use std::io::Read;
                let n = temp_file.read(&mut buf)?;
                if n == 0 {
                    break;
                }
                zip.write_all(&buf[..n])?;
            }
        }

        zip.finish()?;
        Ok(())
    }
}

pub(crate) fn make_cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", col_to_letters(col), row)
}

pub(crate) fn col_to_letters(mut col: u32) -> String {
    let mut result = Vec::new();
    loop {
        result.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result.reverse();
    String::from_utf8(result).unwrap()
}

pub(crate) fn xml_escape(s: &str) -> String {
    if !s.contains(['&', '<', '>', '"', '\'']) {
        return s.to_string();
    }
    let mut out = String::with_capacity(s.len() + 8);
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}
pub(crate) fn write_cell<W: Write>(
    w: &mut W,
    cell_ref: &str,
    value: &CellValue,
    style_idx: Option<usize>,
) -> Result<()> {
    let s = match style_idx {
        Some(0) | None => String::new(),
        Some(n) => format!(" s=\"{n}\""),
    };

    match value {
        CellValue::Blank => {
            write!(w, "<c r=\"{cell_ref}\"{s}/>")?;
        }
        CellValue::Number(n) => {
            write!(w, "<c r=\"{cell_ref}\"{s}><v>{n}</v></c>")?;
        }
        CellValue::Text(text) => {
            let escaped = xml_escape(text);
            write!(
                w,
                "<c r=\"{cell_ref}\"{s} t=\"inlineStr\"><is><t>{escaped}</t></is></c>"
            )?;
        }
        CellValue::Bool(b) => {
            let val = if *b { 1 } else { 0 };
            write!(w, "<c r=\"{cell_ref}\"{s} t=\"b\"><v>{val}</v></c>")?;
        }
        CellValue::Formula(f) => {
            let escaped = xml_escape(f);
            write!(w, "<c r=\"{cell_ref}\"{s}><f>{escaped}</f><v/></c>")?;
        }
    }
    Ok(())
}

pub(crate) fn zip_write_str<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    path: &str,
    content: &str,
    options: SimpleFileOptions,
) -> Result<()> {
    zip.start_file(path, options)?;
    zip.write_all(content.as_bytes())?;
    Ok(())
}
