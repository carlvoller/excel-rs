use std::io::Write;

use anyhow::Result;

pub struct Sheet {
    pub sheet_buf: Vec<u8>,
    pub _name: String,
    pub id: u16,
    pub is_closed: bool,
}

impl Sheet {
    pub fn new(name: String, id: u16) -> Self {
        let mut sheet_buf = vec![];

        // Writes Sheet Header
        sheet_buf.write(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n<sheetData>\n").ok();

        Sheet {
            sheet_buf,
            id,
            _name: name,
            is_closed: false,
        }
    }

    pub fn write_row(&mut self, row_num: u32, data: Vec<&[u8]>) -> Result<()> {
        // let mut sheet_buf = self.sheet_buf;

        let mut escaped_vec = [0; 200];

        // TODO: Proper Error Handling
        if !self.is_closed {
            self.sheet_buf.write(b"<row r=\"")?;
            self.sheet_buf.write(row_num.to_string().as_bytes())?;
            self.sheet_buf.write(b"\">")?;

            let mut col = 0;
            for datum in data {
                let (ref_id, pos) = self.ref_id(col, row_num)?;

                let length;
                (escaped_vec, length) = self.escape(datum, escaped_vec);

                self.sheet_buf.write(b"<c r=\"")?;
                self.sheet_buf.write(&ref_id.as_slice()[0..pos])?;
                self.sheet_buf.write(b"\" t=\"str\"><v>")?;
                self.sheet_buf.write(&escaped_vec[..length])?;
                self.sheet_buf.write(b"</v></c>")?;

                escaped_vec.fill(0);

                col += 1;
            }

            self.sheet_buf.write(b"</row>")?;
        }

        Ok(())
    }

    fn escape(&self, bytes: &[u8], mut escaped: [u8; 200]) -> ([u8; 200], usize) {
        let mut i = 0;
        for c in bytes {
            if matches!(c, b'<' | b'>' | b'&' | b'\'' | b'\"') {
                let mut delta = 4;
                let _ = match c {
                    b'<' => &escaped[i..i + 4].copy_from_slice(b"&lt;"),
                    b'>' => &escaped[i..i + 4].copy_from_slice(b"&gt;"),
                    b'\'' => {
                        delta += 2;
                        &escaped[i..i + 6].copy_from_slice(b"&apos;")
                    }
                    b'&' => {
                        delta += 1;
                        &escaped[i..i + 5].copy_from_slice(b"&amp;")
                    },
                    b'"' => {
                        delta += 2;
                        &escaped[i..i + 6].copy_from_slice(b"&quot;")
                    },
                    b'\t' => &escaped[i..i + 4].copy_from_slice(b"&#9;"),
                    b'\n' => {
                        delta += 1;
                        &escaped[i..i + 5].copy_from_slice(b"&#10;")
                    }
                    b'\r' => {
                        delta += 1;
                        &escaped[i..i + 5].copy_from_slice(b"&#13;")
                    }
                    b' ' => {
                        delta += 1;
                        &escaped[i..i + 5].copy_from_slice(b"&#32;")
                    }
                    _ => {
                        unreachable!(
                            "Only '<', '>','\', '&', '\"', '\\t', '\\r', '\\n', and ' ' are escaped"
                        );
                    }
                };
                i += delta;
            } else {
                escaped[i] = *c;
                i += 1;
            }
        }

        (escaped, i)
    }

    pub fn close(mut self) -> Result<Vec<u8>> {
        self.sheet_buf.write(b"\n</sheetData>\n</worksheet>\n")?;
        Ok(self.sheet_buf)
    }

    fn ref_id(&self, col: u32, row: u32) -> Result<([u8; 12], usize)> {
        let mut final_arr: [u8; 12] = [0; 12];
        let letter = self.col_to_letter(col);

        let mut pos: usize = 0;
        for c in letter {
            if c != 0 {
                final_arr[pos] = c;
                pos += 1;
            }
        }

        // Convert from number to string manually
        let mut row_in_chars_arr: [u8; 9] = [0; 9];
        let mut row = row;
        let mut char_pos = 8;
        let mut digits = 0;
        while row > 0 {
            row_in_chars_arr[char_pos] = b'0' + (row % 10) as u8;
            row = row / 10;
            char_pos -= 1;
            digits += 1;
        }

        for i in 0..digits {
            final_arr[pos] = row_in_chars_arr[char_pos + i + 1];
            pos += 1;
        }

        // for c in lexical::to_string(row).as_bytes() {
        //     if *c != 0 {
        //         final_arr[pos] = c.to_owned();
        //         pos += 1;
        //     }
        // }

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
