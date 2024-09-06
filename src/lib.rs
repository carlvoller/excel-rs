mod export_to_xlsx;
mod ssl;
mod xlsx;

use export_to_xlsx::{
    export_ndarray_to_custom_xlsx, export_pg_client_to_custom_xlsx, export_to_custom_xlsx,
};
use numpy::PyReadonlyArray2;
use postgres::Client;
use pyo3::{
    prelude::*,
    types::{PyBool, PyBytes, PyString},
};
use ssl::SkipServerVerification;
use tokio_postgres_rustls::MakeRustlsConnect;

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

        let mut config = rustls::ClientConfig::builder()
            .with_root_certificates(rustls::RootCertStore::empty())
            .with_no_client_auth();

        if disable_strict_ssl.is_true() {
            config
                .dangerous()
                .set_certificate_verifier(SkipServerVerification::new())
        }

        let tls = MakeRustlsConnect::new(config);

        let mut client = match Client::connect(&conn_string, tls) {
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
