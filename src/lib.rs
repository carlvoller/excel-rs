mod export_to_xlsx;
mod xlsx;

use export_to_xlsx::export_to_custom_xlsx;
use pyo3::{prelude::*, types::PyBytes};

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

    Ok(())
}
