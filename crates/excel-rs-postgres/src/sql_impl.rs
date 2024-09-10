use std::error::Error;

use chrono::DateTime;
use postgres::types::{FromSql, Type};
use postgres_money::Money;
use postgres_protocol::types;
use rust_decimal::Decimal;

pub struct ExcelBytesBorrowed<'a>(pub &'a [u8]);
pub struct ExcelBytes(pub Box<[u8]>);

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
            Type::INT2 => ExcelBytes(Box::from(
                i16::from_be_bytes(raw.try_into().ok().unwrap())
                    .to_string()
                    .as_bytes(),
            )),
            Type::INT4 => ExcelBytes(Box::from(
                i32::from_be_bytes(raw.try_into().ok().unwrap())
                    .to_string()
                    .as_bytes(),
            )),
            Type::INT8 => ExcelBytes(Box::from(
                i64::from_be_bytes(raw.try_into().ok().unwrap())
                    .to_string()
                    .as_bytes(),
            )),
            Type::FLOAT4 => ExcelBytes(Box::from(
                f32::from_be_bytes(raw.try_into().ok().unwrap())
                    .to_string()
                    .as_bytes(),
            )),
            Type::FLOAT8 => ExcelBytes(Box::from(
                f64::from_be_bytes(raw.try_into().ok().unwrap())
                    .to_string()
                    .as_bytes(),
            )),
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
            | Type::INT4
            | Type::INT2
            | Type::FLOAT4
            | Type::FLOAT8 => true,
            ref ty if ty.name() == "citext" => true,
            _ => false,
        }
    }
}
