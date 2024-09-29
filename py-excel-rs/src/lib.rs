mod postgres;
mod utils;

use std::io::Cursor;

use chrono::NaiveDateTime;
use excel_rs_csv::{bytes_to_csv, get_headers, get_next_record};
use excel_rs_xlsx::WorkBook;
use numpy::PyReadonlyArray2;
use postgres::PyPostgresClient;
use utils::chrono_to_xlsx_date;
use pyo3::{prelude::*, types::{PyBytes, PyList}};

#[pymodule]
fn _excel_rs<'py>(m: &Bound<'py, PyModule>) -> PyResult<()> {
    #[pyfn(m)]
    #[pyo3(name = "csv_to_xlsx")]
    fn csv_to_xlsx<'py>(py: Python<'py>, buf: Bound<'py, PyBytes>) -> Bound<'py, PyBytes> {
        let x = buf.as_bytes();

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

        let mut reader = bytes_to_csv(x);
        let headers = get_headers(&mut reader);

        if headers.is_some() {
            let headers_to_bytes = headers.unwrap().iter().to_owned().collect();
            if let Err(e) = worksheet.write_row(headers_to_bytes) {
                panic!("{e}");
            }
        }

        while let Some(record) = get_next_record(&mut reader) {
            let row_data = record.iter().to_owned().collect();
            if let Err(e) = worksheet.write_row(row_data) {
                panic!("{e}");
            }
        }

        if let Err(e) = worksheet.close() {
            panic!("{e}");
        }

        let final_buffer = workbook.finish().ok().unwrap();

        PyBytes::new_bound(py, &final_buffer.into_inner())
    }

    #[pyfn(m)]
    #[pyo3(name = "py_2d_to_xlsx")]
    fn py_2d_to_xlsx<'py>(
        py: Python<'py>,
        list: PyReadonlyArray2<'py, PyObject>,
    ) -> Bound<'py, PyBytes> {
        let ndarray = list.as_array();

        let ndarray_str = ndarray.mapv(|x| {
            if let Ok(inner_str) = x.extract::<String>(py) {
                inner_str
            } else {
                if let Ok(inner_num) = x.extract::<f64>(py) {
                    if inner_num.is_nan() {
                        String::from("")
                    } else {
                        inner_num.to_string()
                    }
                } else {
                    if let Ok(inner_date) = x.extract::<NaiveDateTime>(py) {
                        format!("{}", inner_date.format("%Y-%m-%d %r"))
                    } else {
                        String::from("")
                    }
                }
            }
        });

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

        for row in ndarray_str.rows() {
            let bytes = row.map(|x| x.as_bytes()).to_vec();
            if let Err(e) = worksheet.write_row(bytes) {
                panic!("{e}");
            }
        }

        if let Err(e) = worksheet.close() {
            panic!("{e}");
        }

        let final_buffer = workbook.finish().ok().unwrap();

        PyBytes::new_bound(py, &final_buffer.into_inner())
    }

    #[pyfn(m)]
    #[pyo3(name = "typed_py_2d_to_xlsx")]
    fn typed_py_2d_to_xlsx<'py>(
        py: Python<'py>,
        list: PyReadonlyArray2<'py, PyObject>,
        types: Bound<'py, PyList>,
    ) -> Bound<'py, PyBytes> {
        let ndarray = list.as_array();

        let ndarray_str = ndarray.mapv(|x| {
            if let Ok(inner_str) = x.extract::<String>(py) {
                inner_str
            } else {
                if let Ok(inner_num) = x.extract::<f64>(py) {
                    if inner_num.is_nan() {
                        String::from("")
                    } else {
                        inner_num.to_string()
                    }
                } else {
                    if let Ok(inner_date) = x.extract::<NaiveDateTime>(py) {
                        format!("{}", chrono_to_xlsx_date(inner_date))
                    } else {
                        String::from("")
                    }
                }
            }
        });

        let mut xlsx_types: Vec<String> = Vec::with_capacity(ndarray.len());

        for item in types.iter() {
            let unwrapped = item.extract::<String>().unwrap_or(String::from(""));
            xlsx_types.push(unwrapped);
        }

        let borrowed_xlsx_types = xlsx_types.iter().map(|x| x.as_str()).collect();

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_typed_worksheet(String::from("Sheet 1"));

        for row in ndarray_str.rows() {
            let bytes = row.map(|x| x.as_bytes()).to_vec();
            if let Err(e) = worksheet.write_row(bytes, &borrowed_xlsx_types) {
                panic!("{e}");
            }
        }

        if let Err(e) = worksheet.close() {
            panic!("{e}");
        }

        let final_buffer = workbook.finish().ok().unwrap();

        PyBytes::new_bound(py, &final_buffer.into_inner())
    }

    m.add_class::<PyPostgresClient>()?;

    Ok(())
}
