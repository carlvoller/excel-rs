mod client;
mod sql_impl;
mod ssl;

use std::io::Cursor;

use anyhow::Result;
pub use client::PostgresClient;
use excel_rs_xlsx::WorkBook;
pub use postgres::fallible_iterator::FallibleIterator;
use postgres::RowIter;
pub use sql_impl::{ExcelBytes, ExcelBytesBorrowed};

pub fn postgres_to_xlsx<'a>(mut iter: RowIter<'a>) -> Result<Vec<u8>> {
    let output_buffer = vec![];
    let mut workbook = WorkBook::new(Cursor::new(output_buffer));
    let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

    let headers = iter.next().ok().unwrap().unwrap();
    let len = headers.len();

    // TODO: Add if len == 0 check

    // Write headers
    let mut row_vec: Vec<&[u8]> = vec![&[]; len];

    for col in 0..len {
        let column = headers.columns().get(col).unwrap();
        row_vec[col] = column.name().as_bytes();
    }

    worksheet.write_row(row_vec)?;

    while let Some(row) = iter.next()? {
        let mut row_vec: Vec<Box<[u8]>> = vec![Box::from([]); len];

        for col in 0..len {
            if let Ok(bytes) = row.try_get::<usize, ExcelBytesBorrowed>(col) {
                row_vec[col] = Box::from(bytes.0);
            } else if let Ok(bytes) = row.try_get::<usize, ExcelBytes>(col) {
                let asdasd = bytes.0;
                row_vec[col] = asdasd
            }
        }

        let new_vec: Vec<&[u8]> = row_vec.iter().map(|x| x.as_ref()).collect();

        worksheet.write_row(new_vec)?;
    }

    worksheet.close()?;

    let final_buffer = workbook.finish()?;

    Ok(final_buffer.into_inner())
}
