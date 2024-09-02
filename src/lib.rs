use std::str;

use anyhow::Result;
use pyo3::{prelude::*, types::PyBytes};
use rust_xlsxwriter::Workbook;

#[pymodule]
fn _csv2xlsx_rs<'py>(m: &Bound<'py, PyModule>) -> PyResult<()> {
    fn export_to_xlsx_rs(x: &[u8]) -> Result<Vec<u8>> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let mut reader = csv::ReaderBuilder::new().from_reader(x);
        let mut record = csv::ByteRecord::new();

        let headers = reader.byte_headers()?;
        for col in 0..headers.len() {
            let item = str::from_utf8(headers.get(col).unwrap_or(&[]))?;
            worksheet.write(0, col as u16, item)?;
        }

        let mut row = 1;
        while reader.read_byte_record(&mut record)? {
            for col in 0..record.len() {
                let item = str::from_utf8(record.get(col).unwrap_or(&[]))?;
                worksheet.write(row, col as u16, item)?;
            }
            row += 1;
        }

        let buffer = match workbook.save_to_buffer() {
            Ok(buf) => buf,
            Err(e) => panic!("{e}"),
        };

        Ok(buffer)
    }

    #[pyfn(m)]
    #[pyo3(name = "export_to_xlsx")]
    fn export_to_xlsx<'py>(py: Python<'py>, buf: Bound<'py, PyBytes>) -> Bound<'py, PyBytes> {
        let x = buf.as_bytes();
        let xlsx_bytes = match export_to_xlsx_rs(x) {
            Ok(b) => b,
            Err(e) => panic!("{e}"),
        };
        PyBytes::new_bound(py, &xlsx_bytes)
    }

    Ok(())
}
