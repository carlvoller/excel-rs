use std::io::{Seek, Write};

use anyhow::Result;
use zip::{write::SimpleFileOptions, ZipWriter};

pub struct Sheet<'a, W: Write + Seek> {
    pub sheet_buf: &'a mut ZipWriter<W>,
    pub _name: String,
    // pub id: u16,
    pub is_closed: bool,
}

impl<'a, W: Write + Seek> Sheet<'a, W> {
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

        Sheet {
            sheet_buf: writer,
            // id,
            _name: name,
            is_closed: false,
        }
    }

    pub fn write_row(&mut self, row_num: u32, data: Vec<&[u8]>) -> Result<()> {
        let mut escaped_vec = [0; 512];
        let mut final_vec = Vec::with_capacity(512 * data.len());

        // TODO: Proper Error Handling
        let (row_in_chars_arr, digits) = self.num_to_bytes(row_num);

        final_vec.write(b"<row r=\"")?;
        final_vec.write(&row_in_chars_arr[9 - digits..])?;
        final_vec.write(b"\">")?;

        let mut col = 0;
        for datum in data {
            let (ref_id, pos) = self.ref_id(col, (row_in_chars_arr, digits))?;

            let length;
            (escaped_vec, length) = self.escape(datum, escaped_vec);

            final_vec.write(b"<c r=\"")?;
            final_vec.write(&ref_id.as_slice()[0..pos])?;
            final_vec.write(b"\" t=\"str\"><v>")?;
            final_vec.write(&escaped_vec[..length])?;
            final_vec.write(b"</v></c>")?;

            col += 1;
        }

        final_vec.write(b"</row>")?;

        self.sheet_buf.write(&final_vec)?;

        Ok(())
    }

    fn escape(&self, bytes: &[u8], mut escaped: [u8; 512]) -> ([u8; 512], usize) {
        let mut i = 0;
        let len = bytes.len();
        for x in 0..len {
            let c = bytes[x];
            if matches!(c, b'<' | b'>' | b'&' | b'\'' | b'\"') {
                let mut delta = 2;

                escaped[i] = b'&';
                i += 1;

                let _ = match c {
                    b'<' => &escaped[i..i + 2].copy_from_slice(b"lt"),
                    b'>' => &escaped[i..i + 2].copy_from_slice(b"gt"),
                    b'\'' => {
                        delta += 2;
                        &escaped[i..i + 4].copy_from_slice(b"apos")
                    }
                    b'&' => {
                        delta += 1;
                        &escaped[i..i + 3].copy_from_slice(b"amp")
                    }
                    b'"' => {
                        delta += 2;
                        &escaped[i..i + 4].copy_from_slice(b"quot")
                    }
                    b'\t' => &escaped[i..i + 2].copy_from_slice(b"#9"),
                    b'\n' => {
                        delta += 1;
                        &escaped[i..i + 3].copy_from_slice(b"#10")
                    }
                    b'\r' => {
                        delta += 1;
                        &escaped[i..i + 3].copy_from_slice(b"#13")
                    }
                    b' ' => {
                        delta += 1;
                        &escaped[i..i + 3].copy_from_slice(b"#32")
                    }
                    _ => {
                        unreachable!(
                            "Only '<', '>','\', '&', '\"', '\\t', '\\r', '\\n', and ' ' are escaped"
                        );
                    }
                };
                escaped[i + delta] = b';';
                i += delta + 1;
            } else {
                // TODO: Handle single cell >512 bytes long
                // if i == 512 {
                //     println!("{:?}", escaped);
                // }
                escaped[i] = c;
                i += 1;
            }
        }

        (escaped, i)
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

    fn ref_id(&self, col: u32, row: ([u8; 9], usize)) -> Result<([u8; 12], usize)> {
        let mut final_arr: [u8; 12] = [0; 12];
        let letter = self.col_to_letter(col);

        let mut pos: usize = 0;
        for c in letter {
            if c != 0 {
                final_arr[pos] = c;
                pos += 1;
            }
        }

        let (row_in_chars_arr, digits) = row;

        for i in 0..digits {
            final_arr[pos] = row_in_chars_arr[(8 - digits) + i + 1];
            pos += 1;
        }

        Ok((final_arr, pos))
    }

    fn col_to_letter(&self, col: u32) -> [u8; 3] {
        let mut result: [u8; 3] = [0; 3];
        let mut col = col as i16;

        result[0] = self.num_to_letter((col % 26) as u8);
        col = col / 26 - 1;

        if col >= 0 {
            result[1] = self.num_to_letter((col % 26) as u8);
            col = col / 26 - 1;
            if col >= 0 {
                result[2] = self.num_to_letter((col % 26) as u8);
            }
        }

        result.reverse();
        result
    }

    fn num_to_letter(&self, number: u8) -> u8 {
        b'A' + number
    }
}
