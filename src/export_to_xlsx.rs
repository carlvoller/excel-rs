use std::io::Cursor;

use super::xlsx::WorkBook as NewWorkBook;

use anyhow::Result;

pub fn export_to_custom_xlsx(x: &[u8]) -> Result<Vec<u8>> {
    let output_buffer = vec![];
    let mut workbook = NewWorkBook::new(Cursor::new(output_buffer));
    let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

    let mut reader = csv::ReaderBuilder::new().from_reader(x);

    let headers = reader.byte_headers()?;
    worksheet.write_row(1, headers.iter().collect())?;

    let mut record = csv::ByteRecord::new();
    let mut row = 2;
    while reader.read_byte_record(&mut record)? {
        worksheet.write_row(row, record.iter().to_owned().collect())?;
        row += 1;
    }

    workbook.write_worksheet(worksheet)?;

    let final_buffer = workbook.finish()?;

    Ok(final_buffer.into_inner())
}
