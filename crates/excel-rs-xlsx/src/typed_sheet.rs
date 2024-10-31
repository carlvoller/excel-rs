use std::{
    collections::VecDeque,
    io::{Seek, Write},
};

use anyhow::Result;
use zip::{write::SimpleFileOptions, ZipWriter};
use chrono::NaiveDateTime;

pub const TYPE_NUMBER: &'static str = "n";
pub const TYPE_DATE: &'static str = "d";
pub const TYPE_STRING: &'static str = "str";

pub struct TypedSheet<'a, W: Write + Seek> {
    pub sheet_buf: &'a mut ZipWriter<W>,
    pub _name: String,
    col_num_to_letter: Vec<Vec<u8>>,
    current_row_num: u32,
    has_auto_filter: bool,
    sheet_data_started: bool,
    freeze_top_row: bool,
}

impl<'a, W: Write + Seek> TypedSheet<'a, W> {
    pub fn new(name: String, id: u16, writer: &'a mut ZipWriter<W>) -> Self {
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(1))
            .large_file(true);

        writer
            .start_file(format!("xl/worksheets/sheet{}.xml", id), options)
            .ok();

        writer.write(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
            xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n").ok();

        TypedSheet {
            sheet_buf: writer,
            _name: name,
            col_num_to_letter: Vec::with_capacity(64),
            current_row_num: 0,
            has_auto_filter: false,
            sheet_data_started: false,
            freeze_top_row: false,
        }
    }

    pub fn freeze_top_row(&mut self) {
        self.freeze_top_row = true;
    }

    pub fn add_auto_filter(&mut self) {
        self.has_auto_filter = true;
    }

    fn write_sheet_views(&mut self) -> Result<()> {
        if self.sheet_data_started {
            return Ok(());
        }
        
        self.sheet_buf.write(b"<sheetViews>\n\
            <sheetView tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\">\n\
            <pane ySplit=\"1\" xSplit=\"0\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" />\n\
            <selection pane=\"topLeft\" />\n\
            <selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\" />\n\
            </sheetView>\n\
            </sheetViews>\n")?;

        Ok(())
    }

    pub fn init_sheet(&mut self) -> Result<()> {
        if self.freeze_top_row {
            self.write_sheet_views()?;
        }
        self.sheet_buf.write(b"<sheetData>\n")?;
        self.sheet_data_started = true;
        Ok(())
    }

    pub fn write_row(&mut self, data: Vec<&[u8]>, types: &Vec<&'static str>) -> Result<()> {
        self.current_row_num += 1;

        let mut final_vec = Vec::with_capacity(512 * data.len());

        let (row_in_chars_arr, digits) = self.num_to_bytes(self.current_row_num);

        final_vec.write(b"<row r=\"")?;
        final_vec.write(&row_in_chars_arr[9 - digits..])?;
        final_vec.write(b"\">")?;

        let mut col = 0;
        if self.current_row_num == 1 {
            for datum in data {
                let (ref_id, pos) = self.ref_id(col, (row_in_chars_arr, digits))?;

                final_vec.write(b"<c r=\"")?;
                final_vec.write(&ref_id.as_slice()[0..pos])?;
                final_vec.write(b"\" t=\"str\"><v>")?;

                let (mut chars, chars_pos) = self.escape_in_place(datum);
                let mut current_pos = 0;
                for char_pos in chars_pos {
                    final_vec.write(&datum[current_pos..char_pos])?;
                    final_vec.write(chars.pop_front().unwrap())?;
                    current_pos = char_pos + 1;
                }

                final_vec.write(&datum[current_pos..])?;
                final_vec.write(b"</v></c>")?;

                col += 1;
            }
        } else {
            for datum in data {
                let (ref_id, pos) = self.ref_id(col, (row_in_chars_arr, digits))?;

                let col_type = *types.get(col).unwrap_or(&"s");

                final_vec.write(b"<c r=\"")?;
                final_vec.write(&ref_id.as_slice()[0..pos])?;
                final_vec.write(b"\" t=\"")?;
                final_vec.write(col_type.as_bytes())?;
                final_vec.write(b"\"><v>")?;

                let (mut chars, chars_pos) = self.escape_in_place(datum);
                let mut current_pos = 0;
                for char_pos in chars_pos {
                    final_vec.write(&datum[current_pos..char_pos])?;
                    final_vec.write(chars.pop_front().unwrap())?;
                    current_pos = char_pos + 1;
                }

                final_vec.write(&datum[current_pos..])?;
                final_vec.write(b"</v></c>")?;

                col += 1;
            }
        }

        final_vec.write(b"</row>")?;

        self.sheet_buf.write(&final_vec)?;

        Ok(())
    }

    pub fn infer_row_types(&self, data: &[&[u8]]) -> Vec<&'static str> {
        data.iter()
            .map(|field| {
                let s = String::from_utf8_lossy(field);
                if s.parse::<i64>().is_ok() {
                    TYPE_NUMBER
                } else if s.parse::<f64>().is_ok() {
                    TYPE_NUMBER
                } else if let Ok(_) = NaiveDateTime::parse_from_str(&s, "%Y-%m-%d") {
                    TYPE_DATE
                } else if let Ok(_) = NaiveDateTime::parse_from_str(&s, "%m/%d/%Y") {
                    TYPE_DATE
                } else if let Ok(_) = NaiveDateTime::parse_from_str(&s, "%d/%m/%Y") {
                    TYPE_DATE
                } else {
                    TYPE_STRING
                }
            })
            .collect()
    }

    fn escape_in_place(&self, bytes: &[u8]) -> (VecDeque<&[u8]>, VecDeque<usize>) {
        let mut special_chars: VecDeque<&[u8]> = VecDeque::new();
        let mut special_char_pos: VecDeque<usize> = VecDeque::new();
        let len = bytes.len();
        for x in 0..len {
            let _ = match bytes[x] {
                b'<' => {
                    special_chars.push_back(b"&lt;".as_slice());
                    special_char_pos.push_back(x);
                }
                b'>' => {
                    special_chars.push_back(b"&gt;".as_slice());
                    special_char_pos.push_back(x);
                }
                b'\'' => {
                    special_chars.push_back(b"&apos;".as_slice());
                    special_char_pos.push_back(x);
                }
                b'&' => {
                    special_chars.push_back(b"&amp;".as_slice());
                    special_char_pos.push_back(x);
                }
                b'"' => {
                    special_chars.push_back(b"&quot;".as_slice());
                    special_char_pos.push_back(x);
                }
                _ => (),
            };
        }

        (special_chars, special_char_pos)
    }

    pub fn close(&mut self) -> Result<()> {
        self.sheet_buf.write(b"</sheetData>\n")?;

        if self.has_auto_filter {
            let num_columns = self.col_num_to_letter.len();
            if num_columns > 0 {
                let last_col_letter = self.col_to_letter(num_columns - 1);
                let auto_filter_range = format!("A1:{}1", String::from_utf8_lossy(last_col_letter));
                self.sheet_buf.write(format!("<autoFilter ref=\"{}\"/>\n", auto_filter_range).as_bytes())?;
            }
        }

        self.sheet_buf.write(b"</worksheet>")?;
        Ok(())
    }

    fn num_to_bytes(&self, n: u32) -> ([u8; 9], usize) {
        let mut row_in_chars_arr: [u8; 9] = [0; 9];
        let mut row = n;
        let mut char_pos = 8;
        let mut digits = 0;
        while row > 0 {
            row_in_chars_arr[char_pos] = b'0' + (row % 10) as u8;
            row = row / 10;
            char_pos -= 1;
            digits += 1;
        }

        (row_in_chars_arr, digits)
    }

    fn ref_id(&mut self, col: usize, row: ([u8; 9], usize)) -> Result<([u8; 12], usize)> {
        let mut final_arr: [u8; 12] = [0; 12];
        let letter = self.col_to_letter(col);

        let mut pos: usize = 0;
        for c in letter {
            final_arr[pos] = *c;
            pos += 1;
        }

        let (row_in_chars_arr, digits) = row;

        for i in 0..digits {
            final_arr[pos] = row_in_chars_arr[(8 - digits) + i + 1];
            pos += 1;
        }

        Ok((final_arr, pos))
    }

    fn col_to_letter(&mut self, col: usize) -> &[u8] {
        if self.col_num_to_letter.len() < col + 1 as usize {
            let mut result = Vec::with_capacity(2);
            let mut col = col as i16;

            loop {
                result.push(b'A' + (col % 26) as u8);
                col = col / 26 - 1;
                if col < 0 {
                    break;
                }
            }

            result.reverse();
            self.col_num_to_letter.push(result);
        }

        &self.col_num_to_letter[col]
    }
}
