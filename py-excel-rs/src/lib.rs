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
use excel_rs_xlsx::typed_sheet::{TYPE_STRING, TYPE_NUMBER, TYPE_DATE};

#[pymodule]
fn _excel_rs<'py>(m: &Bound<'py, PyModule>) -> PyResult<()> {
    #[pyfn(m)]
    #[pyo3(name = "csv_to_xlsx")]
    /// Convert CSV data to XLSX format with optional formatting.
    ///
    /// Args:
    ///     buf (bytes): Input CSV data as bytes
    ///     freeze_top_row (bool, optional): If True, freezes the first row. Defaults to False.
    ///     add_auto_filter (bool, optional): If True, adds auto-filter to columns. Defaults to False.
    ///
    /// Returns:
    ///     bytes: XLSX file content as bytes
    ///
    /// Example:
    ///     >>> with open('input.csv', 'rb') as f:
    ///     ...     xlsx_data = csv_to_xlsx(f.read(), freeze_top_row=True, add_auto_filter=True)
    ///     >>> with open('output.xlsx', 'wb') as f:
    ///     ...     f.write(xlsx_data)
    fn csv_to_xlsx<'py>(
        py: Python<'py>,
        buf: Bound<'py, PyBytes>,
        freeze_top_row: Option<bool>,
        add_auto_filter: Option<bool>,
    ) -> Bound<'py, PyBytes> {
        let x = buf.as_bytes();

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_typed_worksheet(String::from("Sheet 1"));

        if freeze_top_row.unwrap_or(false) {
            worksheet.freeze_top_row();
        }
        if add_auto_filter.unwrap_or(false) {
            worksheet.add_auto_filter();
        }

        worksheet.init_sheet().expect("Failed to initialize worksheet");

        let mut reader = bytes_to_csv(x);
        let headers = get_headers(&mut reader);

        if let Some(headers) = headers {
            let headers_to_bytes = headers.iter().to_owned().collect();
            let header_types = vec![TYPE_STRING; headers.len()];
            if let Err(e) = worksheet.write_row(headers_to_bytes, &header_types) {
                panic!("{e}");
            }
        }

        if let Some(record) = get_next_record(&mut reader) {
            let row_data: Vec<&[u8]> = record.iter().to_owned().collect();
            let types = worksheet.infer_row_types(&row_data);
            if let Err(e) = worksheet.write_row(row_data, &types) {
                panic!("{e}");
            }

            while let Some(record) = get_next_record(&mut reader) {
                let row_data = record.iter().to_owned().collect();
                if let Err(e) = worksheet.write_row(row_data, &types) {
                    panic!("{e}");
                }
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
    /// Convert a 2D numpy array to XLSX format with type inference and optional formatting.
    ///
    /// Args:
    ///     list (numpy.ndarray): 2D input array
    ///     freeze_top_row (bool, optional): If True, freezes the first row. Defaults to False.
    ///     add_auto_filter (bool, optional): If True, adds auto-filter to columns. Defaults to False.
    ///
    /// Returns:
    ///     bytes: XLSX file content as bytes
    ///
    /// Notes:
    ///     - First row is treated as headers (string type)
    ///     - Types are inferred from the second row
    ///     - Supports automatic conversion of strings, numbers, and dates
    ///
    /// Example:
    ///     >>> import numpy as np
    ///     >>> data = np.array([['Name', 'Age'], ['John', 25]])
    ///     >>> xlsx_data = py_2d_to_xlsx(data, freeze_top_row=True)
    fn py_2d_to_xlsx<'py>(
        py: Python<'py>,
        list: PyReadonlyArray2<'py, PyObject>,
        freeze_top_row: Option<bool>,
        add_auto_filter: Option<bool>,
    ) -> Bound<'py, PyBytes> {
        let ndarray = list.as_array();

        let ndarray_str = ndarray.mapv(|x| {
            if let Ok(inner_str) = x.extract::<String>(py) {
                inner_str
            } else if let Ok(inner_num) = x.extract::<f64>(py) {
                if inner_num.is_nan() {
                    String::from("")
                } else {
                    inner_num.to_string()
                }
            } else if let Ok(inner_date) = x.extract::<NaiveDateTime>(py) {
                format!("{}", inner_date.format("%Y-%m-%d %r"))
            } else {
                String::from("")
            }
        });

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_typed_worksheet(String::from("Sheet 1"));

        if freeze_top_row.unwrap_or(false) {
            worksheet.freeze_top_row();
        }
        if add_auto_filter.unwrap_or(false) {
            worksheet.add_auto_filter();
        }

        worksheet.init_sheet().expect("Failed to initialize worksheet");

        if ndarray_str.nrows() > 1 {
            let data_row = ndarray_str.row(1);
            let first_data_row: Vec<&[u8]> = data_row.iter().map(|x| x.as_bytes()).collect();
            let types = worksheet.infer_row_types(&first_data_row);

            let header = ndarray_str.row(0);
            let header_row: Vec<&[u8]> = header.iter().map(|x| x.as_bytes()).collect();
            let header_types = vec![TYPE_STRING; header_row.len()];
            if let Err(e) = worksheet.write_row(header_row, &header_types) {
                panic!("{e}");
            }

            if let Err(e) = worksheet.write_row(first_data_row, &types) {
                panic!("{e}");
            }

            for i in 2..ndarray_str.nrows() {
                let row = ndarray_str.row(i);
                let row_data: Vec<&[u8]> = row.iter().map(|x| x.as_bytes()).collect();
                if let Err(e) = worksheet.write_row(row_data, &types) {
                    panic!("{e}");
                }
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
    /// Convert a 2D numpy array to XLSX format with explicit column types and optional formatting.
    ///
    /// Args:
    ///     list (numpy.ndarray): 2D input array
    ///     types (list): List of column types. Valid types are:
    ///         - 'n': Number
    ///         - 'd': Date
    ///         - 'str': String (default)
    ///     freeze_top_row (bool, optional): If True, freezes the first row. Defaults to False.
    ///     add_auto_filter (bool, optional): If True, adds auto-filter to columns. Defaults to False.
    ///
    /// Returns:
    ///     bytes: XLSX file content as bytes
    ///
    /// Example:
    ///     >>> import numpy as np
    ///     >>> data = np.array([['Name', 'Age', 'Date'],
    ///     ...                  ['John', 25, '2023-01-01']])
    ///     >>> types = ['str', 'n', 'd']
    ///     >>> xlsx_data = typed_py_2d_to_xlsx(data, types, 
    ///     ...                                 freeze_top_row=True, 
    ///     ...                                 add_auto_filter=True)
    fn typed_py_2d_to_xlsx<'py>(
        py: Python<'py>,
        list: PyReadonlyArray2<'py, PyObject>,
        types: Bound<'py, PyList>,
        freeze_top_row: Option<bool>,
        add_auto_filter: Option<bool>,
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

        let xlsx_types: Vec<&'static str> = types.iter().map(|x| {
            match x.extract::<String>().unwrap().as_str() {
                "n" => TYPE_NUMBER,
                "d" => TYPE_DATE,
                _ => TYPE_STRING
            }
        }).collect();

        let output_buffer = vec![];
        let mut workbook = WorkBook::new(Cursor::new(output_buffer));
        let mut worksheet = workbook.get_typed_worksheet(String::from("Sheet 1"));

        if freeze_top_row.unwrap_or(false) {
            worksheet.freeze_top_row();
        }
        if add_auto_filter.unwrap_or(false) {
            worksheet.add_auto_filter();
        }

        worksheet.init_sheet().expect("Failed to initialize worksheet");

        for row in ndarray_str.rows() {
            let bytes = row.map(|x| x.as_bytes()).to_vec();
            if let Err(e) = worksheet.write_row(bytes, &xlsx_types) {
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
