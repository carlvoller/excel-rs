use std::error::Error;
use std::io::Cursor;

use super::xlsx::WorkBook as NewWorkBook;

use anyhow::Result;
use chrono::DateTime;
use numpy::ndarray::Array2;
use postgres::types::{FromSql, Type};
use postgres::RowIter;
use postgres::{fallible_iterator::FallibleIterator, Client};
use postgres_money::Money;
use postgres_protocol::types;
use rust_decimal::Decimal;

struct ExcelBytesBorrowed<'a>(&'a [u8]);
struct ExcelBytes(Box<[u8]>);

// Int8, Money, Timestamp, VarChar, Text, Numeric
impl<'a> FromSql<'a> for ExcelBytesBorrowed<'a> {
    fn from_sql(
        pg_type: &Type,
        raw: &'a [u8],
    ) -> Result<ExcelBytesBorrowed<'a>, Box<dyn Error + Sync + Send>> {
        let out: ExcelBytesBorrowed<'a> = match *pg_type {
            Type::VARCHAR | Type::TEXT | Type::BPCHAR | Type::NAME | Type::UNKNOWN => {
                ExcelBytesBorrowed(raw)
            }
            _ => ExcelBytesBorrowed(&[]),
        };

        // let out: ExcelBytes<'a> = ExcelBytes(raw);
        Ok(out)
    }

    fn accepts(ty: &Type) -> bool {
        match *ty {
            Type::VARCHAR | Type::TEXT | Type::BPCHAR | Type::NAME | Type::UNKNOWN => true,
            ref ty if ty.name() == "citext" => true,
            _ => false,
        }
    }
}

impl<'a> FromSql<'a> for ExcelBytes {
    fn from_sql(pg_type: &Type, raw: &'a [u8]) -> Result<ExcelBytes, Box<dyn Error + Sync + Send>> {
        let out: ExcelBytes = match *pg_type {
            Type::TIMESTAMP => match types::timestamp_from_sql(raw) {
                Ok(parsed) => {
                    // println!("{parsed}");
                    let bytes = Box::from(
                        format!(
                            "{}",
                            DateTime::from_timestamp((parsed / 1000000) + 946684800, 0)
                                .unwrap()
                                .format("%Y-%m-%d %r")
                        )
                        .as_bytes(),
                    );

                    ExcelBytes(bytes)
                }
                Err(_) => ExcelBytes(Box::from([])),
            },
            Type::INT2 => ExcelBytes(match types::int2_from_sql(raw) {
                Ok(parsed) => Box::from(parsed.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::INT4 => ExcelBytes(match types::int4_from_sql(raw) {
                Ok(parsed) => Box::from(parsed.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::INT8 => ExcelBytes(match types::int8_from_sql(raw) {
                Ok(parsed) => Box::from(parsed.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::FLOAT4 => ExcelBytes(match types::float4_from_sql(raw) {
                Ok(parsed) => Box::from(parsed.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::FLOAT8 => ExcelBytes(match types::float8_from_sql(raw) {
                Ok(parsed) => Box::from(parsed.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::NUMERIC => ExcelBytes(match <Decimal as FromSql>::from_sql(pg_type, raw) {
                Ok(num) => Box::from(num.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            Type::MONEY => ExcelBytes(match <Money as FromSql>::from_sql(pg_type, raw) {
                Ok(money) => Box::from(money.to_string().as_bytes()),
                Err(_) => Box::from([]),
            }),
            _ => ExcelBytes(Box::from(raw)),
        };
        Ok(out)
    }

    fn accepts(ty: &Type) -> bool {
        match *ty {
            Type::TIMESTAMP
            | Type::MONEY
            | Type::NUMERIC
            | Type::INT8
            | Type::FLOAT4
            | Type::FLOAT8 => true,
            ref ty if ty.name() == "citext" => true,
            _ => false,
        }
    }
}

pub fn export_to_custom_xlsx(x: &[u8]) -> Result<Vec<u8>> {
    let output_buffer = vec![];
    let mut workbook = NewWorkBook::new(Cursor::new(output_buffer));
    let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

    let mut reader = csv::ReaderBuilder::new().from_reader(x);

    let headers = reader.byte_headers()?;
    worksheet.write_row(1, headers.iter().to_owned().collect())?;

    let mut record = csv::ByteRecord::new();
    let mut row = 2;
    while reader.read_byte_record(&mut record)? {
        let row_data = record.iter().to_owned().collect();
        worksheet.write_row(row, row_data)?;
        row += 1;
    }

    worksheet.close()?;

    let final_buffer = workbook.finish()?;

    Ok(final_buffer.into_inner())
}

pub fn export_ndarray_to_custom_xlsx(x: Array2<String>) -> Result<Vec<u8>> {
    let output_buffer = vec![];
    let mut workbook = NewWorkBook::new(Cursor::new(output_buffer));
    let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

    let mut row_num = 1;
    for row in x.rows() {
        let bytes = row.map(|x| x.as_bytes()).to_vec();
        worksheet.write_row(row_num, bytes)?;

        row_num += 1;
    }

    worksheet.close()?;

    let final_buffer = workbook.finish()?;

    Ok(final_buffer.into_inner())
}

pub fn export_pg_client_to_custom_xlsx<'a>(query: &str, client: &'a mut Client) -> Result<Vec<u8>> {
    let params: Vec<String> = vec![];
    let mut iter: RowIter<'a> = client.query_raw(query, params)?;

    let output_buffer = vec![];
    let mut workbook = NewWorkBook::new(Cursor::new(output_buffer));
    let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

    let mut row_num = 0;

    while let Some(row) = iter.next()? {
        let len = row.len();
        row_num += 1;
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

        worksheet.write_row(row_num, new_vec)?;
    }
    worksheet.close()?;

    let final_buffer = workbook.finish()?;

    Ok(final_buffer.into_inner())
}
