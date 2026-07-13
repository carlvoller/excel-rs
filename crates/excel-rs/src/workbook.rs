use super::format::XlsxFormatter;
use crate::error::{ExcelError, Result};
use std::io::{Seek, Write};
use zip::ZipWriter;

use super::sheet::Sheet;

/// A WorkBook represents one excel file.
pub struct WorkBook<W: Write + Seek> {
    formatter: XlsxFormatter<W>,
    sheet_names: Vec<String>,
}

impl<W: Write + Seek> WorkBook<W> {
    /// Create a new WorkBook. `writer` is the destination to write the excel file to
    pub fn new(writer: W) -> Self {
        let zip_writer = ZipWriter::new(writer);

        WorkBook {
            formatter: XlsxFormatter::new(zip_writer),
            sheet_names: Vec::new(),
        }
    }

    /// Creates a new Sheet with the given name.
    pub fn new_worksheet<'a>(&'a mut self, name: String) -> Result<Sheet<'a, W>> {
        if self.sheet_names.len() == crate::MAX_SHEETS {
            return Err(ExcelError::TooManySheets);
        }

        if self.sheet_names.contains(&name) {
            return Err(ExcelError::SheetAlreadyExists);
        }

        let id = self.sheet_names.len() as u16 + 1;
        self.sheet_names.push(name.clone());
        Sheet::new(name, id, &mut self.formatter.zip_writer)
    }

    /// Finish writing to the excel file. This closes the file and wraps up any remaining operations.
    pub fn finish(self) -> Result<W> {
        self.formatter.finish(self.sheet_names)
    }
}
