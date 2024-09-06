mod export_to_xlsx;
mod xlsx;

use export_to_xlsx::{
    export_ndarray_to_custom_xlsx, export_pg_client_to_custom_xlsx, export_to_custom_xlsx,
};
use native_tls::TlsConnector;
use numpy::PyReadonlyArray2;
use postgres::Client;
use postgres_native_tls::MakeTlsConnector;
use pyo3::{
    prelude::*,
    types::{PyBool, PyBytes, PyString},
};

#[pymodule]
fn _excel_rs<'py>(m: &Bound<'py, PyModule>) -> PyResult<()> {
    #[pyfn(m)]
    #[pyo3(name = "export_to_xlsx")]
    fn export_to_xlsx<'py>(py: Python<'py>, buf: Bound<'py, PyBytes>) -> Bound<'py, PyBytes> {
        let x = buf.as_bytes();
        let xlsx_bytes = match export_to_custom_xlsx(x) {
            Ok(b) => b,
            Err(e) => panic!("{e}"),
        };
        PyBytes::new_bound(py, &xlsx_bytes)
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
                    String::from("")
                }
            }
        });

        let xlsx_bytes = match export_ndarray_to_custom_xlsx(ndarray_str) {
            Ok(b) => b,
            Err(e) => panic!("{e}"),
        };

        PyBytes::new_bound(py, &xlsx_bytes)
    }

    #[pyfn(m)]
    #[pyo3(name = "pg_query_to_xlsx")]
    fn pg_query_to_xlsx<'py>(
        py: Python<'py>,
        py_query: Bound<'py, PyString>,
        py_conn_string: Bound<'py, PyString>,
        disable_strict_ssl: Bound<'py, PyBool>,
    ) -> Bound<'py, PyBytes> {
        let conn_string: String = match py_conn_string.extract() {
            Ok(s) => s,
            Err(e) => panic!("{e}"),
        };

        let query: String = match py_query.extract() {
            Ok(s) => s,
            Err(e) => panic!("{e}"),
        };

        let connector: TlsConnector;

        if disable_strict_ssl.is_true() {
            connector = TlsConnector::builder()
                .danger_accept_invalid_certs(true)
                .danger_accept_invalid_hostnames(true)
                .build()
                .ok()
                .unwrap();
        } else {
            connector = TlsConnector::new().ok().unwrap();
        }

        let connector = MakeTlsConnector::new(connector);

        let mut client = match Client::connect(&conn_string, connector) {
            Ok(c) => c,
            Err(e) => panic!("{e}"),
        };

        let xlsx_bytes = match export_pg_client_to_custom_xlsx(&query, &mut client) {
            Ok(b) => b,
            Err(e) => panic!("{e}"),
        };

        PyBytes::new_bound(py, &xlsx_bytes)
    }

    Ok(())
}
