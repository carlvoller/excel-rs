use super::format::XlsxFormatter;
use std::{io::Seek, io::Write};
use anyhow::Result;
use zip::ZipWriter;

use super::sheet::Sheet;
use super::typed_sheet::TypedSheet;

pub struct WorkBook<W: Write + Seek> {
    formatter: XlsxFormatter<W>,
    num_of_sheets: u16,
}

impl<W: Write + Seek> WorkBook<W> {
    pub fn new(writer: W) -> Self {
        let zip_writer = ZipWriter::new(writer);

        WorkBook {
            formatter: XlsxFormatter::new(zip_writer),
            num_of_sheets: 0,
        }
    }

    pub fn get_worksheet(&mut self, name: String) -> Sheet<W> {
        self.num_of_sheets += 1;
        Sheet::new(name, self.num_of_sheets, &mut self.formatter.zip_writer)
    }

    pub fn get_typed_worksheet(&mut self, name: String) -> TypedSheet<W> {
        self.num_of_sheets += 1;
        TypedSheet::new(name, self.num_of_sheets, &mut self.formatter.zip_writer)
    }
    
    pub fn finish(self) -> Result<W> {
        let result = self.formatter.finish(self.num_of_sheets)?;
        Ok(result)
    }
}
