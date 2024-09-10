use std::{borrow::Cow, io::Cursor};

use excel_rs_postgres::{ExcelBytes, ExcelBytesBorrowed, FallibleIterator, PostgresClient};
use excel_rs_xlsx::WorkBook;
use pyo3::{pyclass, pymethods, PyResult};

#[pyclass]
pub struct PyPostgresClient {
    client: Option<PostgresClient>,
}

#[pymethods]
impl PyPostgresClient {
    #[staticmethod]
    pub fn new(conn_string: &str) -> PyPostgresClient {
        PyPostgresClient {
            client: Some(PostgresClient::new(conn_string)),
        }
    }

    pub fn get_columns(
        &mut self,
        table_name: &str,
        schema_name: &str,
        excluded: Vec<String>,
    ) -> PyResult<Vec<String>> {
        let mut query =
            String::from("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '");

        query.push_str(table_name);

        if schema_name != "" {
            query.push_str("' AND TABLE_SCHEMA = '");
            query.push_str(schema_name);
        }

        if excluded.len() > 0 {
            query.push_str("' AND COLUMN_NAME NOT IN (");
            query.push_str(
                &excluded
                    .iter()
                    .map(|x| format!("'{x}'")) // Add quotes
                    .collect::<Vec<String>>()
                    .join(", "), // Add commas
            );
            query.push_str(")");
        } else {
            query.push_str("'");
        }

        let res = match &mut self.client {
            Some(client) => client.make_query(&query, vec![]),
            None => panic!("Client not set up"),
        };

        let iter = match res {
            Ok(iter) => iter,
            Err(e) => panic!("{e}"),
        };

        let cols: Vec<String> = match iter.map(|row| Ok(row.get::<usize, String>(0))).collect() {
            Ok(s) => s,
            Err(e) => panic!("{e}"),
        };

        Ok(cols)
    }

    pub fn get_xlsx_from_query(&mut self, query: &str) -> PyResult<Cow<[u8]>> {
        let res = match &mut self.client {
            Some(client) => client.make_query(&query, vec![]),
            None => panic!("Client not set up"),
        };

        let mut iter = match res {
            Ok(iter) => iter,
            Err(e) => panic!("{e}"),
        };

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

        let headers = iter.next().ok().unwrap().unwrap();
        let len = headers.len();

        // Write headers
        let mut row_vec: Vec<&[u8]> = vec![&[]; len];

        for col in 0..len {
            let column = headers.columns().get(col).unwrap();
            let name = column.name();
            row_vec[col] = name.as_bytes();
        }

        if let Err(e) = worksheet.write_row(row_vec) {
            panic!("{e}");
        }

        while let Ok(Some(row)) = iter.next() {
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

            if let Err(e) = worksheet.write_row(new_vec) {
                panic!("{e}");
            }
        }

        if let Err(e) = worksheet.close() {
            panic!("{e}");
        }

        let final_buffer = workbook.finish().ok().unwrap().into_inner();

        Ok(Cow::from(final_buffer))
    }

    pub fn close(&mut self) -> PyResult<()> {
        let client = Option::take(&mut self.client);

        if let Some(client) = client {
            client.close().ok();
        }

        Ok(())
    }
}
