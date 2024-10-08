use std::{
    collections::VecDeque,
    io::{Seek, Write},
};

use anyhow::Result;
use zip::{write::SimpleFileOptions, ZipWriter};

pub struct TypedSheet<'a, W: Write + Seek> {
    pub sheet_buf: &'a mut ZipWriter<W>,
    pub _name: String,
    // pub id: u16,
    // pub is_closed: bool,
    col_num_to_letter: Vec<Vec<u8>>,
    current_row_num: u32,
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

        // Writes Sheet Header
        writer.write(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n<sheetData>\n").ok();

        TypedSheet {
            sheet_buf: writer,
            // id,
            _name: name,
            // is_closed: false,
            col_num_to_letter: Vec::with_capacity(64),
            current_row_num: 0,
        }
    }

    // TOOD: Use ShortVec over Vec for cell ID
    pub fn write_row(&mut self, data: Vec<&[u8]>, types: &Vec<&str>) -> Result<()> {
        self.current_row_num += 1;

        let mut final_vec = Vec::with_capacity(512 * data.len());

        // TODO: Proper Error Handling
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
        self.sheet_buf.write(b"\n</sheetData>\n</worksheet>\n")?;
        Ok(())
    }

    fn num_to_bytes(&self, n: u32) -> ([u8; 9], usize) {
        // Convert from number to string manually
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
