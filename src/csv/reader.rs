use std::collections::HashMap;
use std::fs::File;
use std::io::{self, BufWriter, Read, Write};
use std::sync::Arc;

use quick_xml::Reader as XmlReader;
use quick_xml::events::Event;
use zip::ZipArchive;

#[derive(Debug)]
pub struct Row {
    pub cells: Vec<String>,
}

#[inline]
fn attr_val(attr: &quick_xml::events::attributes::Attribute) -> String {
    let raw = std::str::from_utf8(attr.value.as_ref()).unwrap_or("");
    quick_xml::escape::unescape(raw)
        .unwrap_or_default()
        .into_owned()
}

#[inline]
fn text_val(e: &quick_xml::events::BytesText) -> String {
    let raw = std::str::from_utf8(e.as_ref()).unwrap_or("");
    quick_xml::escape::unescape(raw)
        .unwrap_or_default()
        .into_owned()
}

pub struct XlsxReader {
    path: String,
    sheet_paths: HashMap<String, String>,
    sheet_order: Vec<String>,
    shared_strings: Arc<Vec<String>>,
}

impl XlsxReader {
    pub fn open(path: &str) -> io::Result<Self> {
        let file = File::open(path)?;
        let mut archive =
            ZipArchive::new(file).map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;

        let (sheet_order, rid_to_name) = parse_workbook(&mut archive)?;
        let rid_to_path = parse_workbook_rels(&mut archive)?;

        let sheet_paths: HashMap<String, String> = rid_to_name
            .into_iter()
            .filter_map(|(rid, name)| rid_to_path.get(&rid).map(|p| (name, p.clone())))
            .collect();

        let shared_strings = Arc::new(parse_shared_strings(&mut archive)?);

        Ok(XlsxReader {
            path: path.to_string(),
            sheet_paths,
            sheet_order,
            shared_strings,
        })
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_order
    }

    pub fn stream_rows(&self, sheet_name: &str) -> io::Result<RowIter> {
        let zip_path = self
            .sheet_paths
            .get(sheet_name)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("sheet '{sheet_name}' not found"),
                )
            })?
            .clone();

        let file = File::open(&self.path)?;
        let mut archive =
            ZipArchive::new(file).map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;

        let xml: Box<[u8]> = {
            let mut entry = archive
                .by_name(&zip_path)
                .map_err(|e| io::Error::new(io::ErrorKind::NotFound, e))?;
            let mut buf = Vec::with_capacity(entry.size() as usize);
            let mut chunk = [0u8; 65536];
            loop {
                let n = entry.read(&mut chunk)?;
                if n == 0 {
                    break;
                }
                buf.extend_from_slice(&chunk[..n]);
            }
            buf.into_boxed_slice()
        };

        Ok(RowIter {
            xml,
            offset: 0,
            shared_strings: Arc::clone(&self.shared_strings),
            state: ParseState::new(),
            buf: Vec::with_capacity(256),
            done: false,
        })
    }

    pub fn sheet_to_csv<W: Write>(&self, sheet_name: &str, out: &mut W) -> io::Result<usize> {
        let mut count = 0;
        for row in self.stream_rows(sheet_name)? {
            let row = row?;
            let line = row
                .cells
                .iter()
                .map(|c| csv_escape(c))
                .collect::<Vec<_>>()
                .join(",");
            writeln!(out, "{line}")?;
            count += 1;
        }
        Ok(count)
    }

    pub fn all_sheets_to_csv(&self, prefix: &str) -> io::Result<()> {
        for name in &self.sheet_order {
            let filename = format!("{}_{}.csv", prefix, name.replace(' ', "_"));
            let file = File::create(&filename)?;
            let mut out = BufWriter::new(file);
            let n = self.sheet_to_csv(name, &mut out)?;
            println!("  {name} → {filename} ({n} rows)");
        }
        Ok(())
    }
}

pub struct RowIter {
    xml: Box<[u8]>,
    offset: usize,
    shared_strings: Arc<Vec<String>>,
    state: ParseState,
    buf: Vec<u8>,
    done: bool,
}

struct ParseState {
    row: Vec<String>,
    col: u32,
    next_col: u32,
    cell_type: CellType,
    in_v: bool,
    in_t: bool,
    value_buf: String,
    in_row: bool,
}

#[derive(Clone, Copy)]
enum CellType {
    Number,
    SharedStr,
    Inline,
    Bool,
    Str,
    Error,
}

impl ParseState {
    fn new() -> Self {
        ParseState {
            row: Vec::new(),
            col: 0,
            next_col: 0,
            cell_type: CellType::Number,
            in_v: false,
            in_t: false,
            value_buf: String::new(),
            in_row: false,
        }
    }
}

impl Iterator for RowIter {
    type Item = io::Result<Row>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        let slice = &self.xml[self.offset..];
        let mut xml = XmlReader::from_reader(slice);
        xml.config_mut().trim_text(true);

        loop {
            self.buf.clear();

            match xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if e.name().as_ref() == b"row" =>
                {
                    self.state.row.clear();
                    self.state.next_col = 0;
                    self.state.in_row = true;
                }

                Ok(Event::Start(ref e)) if e.name().as_ref() == b"c" => {
                    let mut col_ref = String::new();
                    let mut cell_type = CellType::Number;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => col_ref = attr_val(&attr),
                            b"t" => {
                                cell_type = match attr_val(&attr).as_str() {
                                    "s" => CellType::SharedStr,
                                    "inlineStr" => CellType::Inline,
                                    "b" => CellType::Bool,
                                    "str" => CellType::Str,
                                    "e" => CellType::Error,
                                    _ => CellType::Number,
                                }
                            }
                            _ => {}
                        }
                    }

                    let col = col_ref_to_index(&col_ref);

                    while self.state.next_col < col {
                        self.state.row.push(String::new());
                        self.state.next_col += 1;
                    }

                    self.state.col = col;
                    self.state.cell_type = cell_type;
                    self.state.value_buf.clear();
                    self.state.in_v = false;
                    self.state.in_t = false;
                }

                Ok(Event::Start(ref e)) => match e.name().as_ref() {
                    b"v" => self.state.in_v = true,
                    b"t" => self.state.in_t = true,
                    _ => {}
                },

                Ok(Event::Text(ref e)) => {
                    if self.state.in_v || self.state.in_t {
                        self.state.value_buf.push_str(&text_val(e));
                    }
                }

                Ok(Event::End(ref e)) => match e.name().as_ref() {
                    b"v" => self.state.in_v = false,
                    b"t" => self.state.in_t = false,
                    b"c" => {
                        let raw = self.state.value_buf.trim().to_string();
                        let value = match self.state.cell_type {
                            CellType::SharedStr => {
                                let idx: usize = raw.parse().unwrap_or(0);
                                self.shared_strings.get(idx).cloned().unwrap_or_default()
                            }
                            CellType::Bool => {
                                if raw == "1" {
                                    "TRUE".into()
                                } else {
                                    "FALSE".into()
                                }
                            }
                            _ => raw,
                        };
                        self.state.row.push(value);
                        self.state.next_col = self.state.col + 1;
                    }

                    b"row" => {
                        self.offset += xml.buffer_position() as usize;

                        if self.state.in_row && !self.state.row.is_empty() {
                            self.state.in_row = false;
                            return Some(Ok(Row {
                                cells: std::mem::take(&mut self.state.row),
                            }));
                        }
                    }

                    b"sheetData" => {
                        self.done = true;
                        return None;
                    }

                    _ => {}
                },

                Ok(Event::Eof) => {
                    self.done = true;
                    return None;
                }

                Err(e) => {
                    self.done = true;
                    return Some(Err(io::Error::new(io::ErrorKind::InvalidData, e)));
                }

                _ => {}
            }
        }
    }
}

fn parse_workbook(
    archive: &mut ZipArchive<File>,
) -> io::Result<(Vec<String>, HashMap<String, String>)> {

    let bytes = slurp_entry(archive, "xl/workbook.xml")?;
    let mut xml = XmlReader::from_reader(bytes.as_slice());
    xml.config_mut().trim_text(true);

    let mut order = Vec::new();
    let mut rid_map = HashMap::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheet" => {
                let (mut name, mut rid) = (String::new(), String::new());
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => name = attr_val(&attr),
                        b"r:id" | b"id" => rid = attr_val(&attr),
                        _ => {}
                    }
                }
                if !name.is_empty() && !rid.is_empty() {
                    order.push(name.clone());
                    rid_map.insert(rid, name);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
            _ => {}
        }
        buf.clear();
    }
    Ok((order, rid_map))
}

fn parse_workbook_rels(archive: &mut ZipArchive<File>) -> io::Result<HashMap<String, String>> {
    let bytes = slurp_entry(archive, "xl/_rels/workbook.xml.rels")?;
    let mut xml = XmlReader::from_reader(bytes.as_slice());
    xml.config_mut().trim_text(true);

    let mut map = HashMap::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let (mut id, mut target, mut is_sheet) = (String::new(), String::new(), false);
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = attr_val(&attr),
                        b"Target" => target = attr_val(&attr),
                        b"Type" => is_sheet = attr_val(&attr).contains("worksheet"),
                        _ => {}
                    }
                }
                if is_sheet && !id.is_empty() {
                    map.insert(id, normalize_path(&target));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(map)
}

fn parse_shared_strings(archive: &mut ZipArchive<File>) -> io::Result<Vec<String>> {
    if archive.by_name("xl/sharedStrings.xml").is_err() {
        return Ok(Vec::new());
    }

    let mut entry = archive
        .by_name("xl/sharedStrings.xml")
        .map_err(|e| io::Error::new(io::ErrorKind::NotFound, e))?;
    let buf_reader = std::io::BufReader::with_capacity(64 * 1024, &mut entry);
    let mut xml = XmlReader::from_reader(buf_reader);
    xml.config_mut().trim_text(false);

    let mut strings = Vec::new();
    let mut current = String::new();
    let mut in_t = false;
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"si" => current.clear(),
                b"t" => in_t = true,
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.name().as_ref() {
                b"si" => strings.push(current.clone()),
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(ref e)) if in_t => current.push_str(&text_val(e)),
            Ok(Event::Eof) => break,
            Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(strings)
}

fn slurp_entry(archive: &mut ZipArchive<File>, path: &str) -> io::Result<Vec<u8>> {
    let mut entry = archive
        .by_name(path)
        .map_err(|e| io::Error::new(io::ErrorKind::NotFound, format!("'{path}': {e}")))?;
    let mut buf = Vec::with_capacity(entry.size() as usize);
    io::copy(&mut entry, &mut buf)?;
    Ok(buf)
}

fn normalize_path(target: &str) -> String {
    let t = target.trim_start_matches('/');
    if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("xl/{t}")
    }
}

fn col_ref_to_index(cell_ref: &str) -> u32 {
    let letters = cell_ref.trim_end_matches(|c: char| c.is_ascii_digit());
    if letters.is_empty() {
        return 0;
    }
    letters
        .bytes()
        .fold(0u32, |acc, b| acc * 26 + (b - b'A') as u32 + 1)
        - 1
}

pub fn csv_escape(s: &str) -> String {
    if s.is_empty() {
        return String::new();
    }
    if !s.contains([',', '"', '\n', '\r']) {
        return s.to_string();
    }
    format!("\"{}\"", s.replace('"', "\"\""))
}
