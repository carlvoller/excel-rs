mod export_to_xlsx;
mod xlsx;

use std::fs::File;
use std::{io::prelude::*, time::Instant};

use anyhow::Result;
pub use export_to_xlsx::export_to_custom_xlsx;

fn convert_csv_to_xlsx(filename: &str) -> Result<()> {
    let mut f = File::open(filename)?;
    let mut buffer: Vec<u8> = Vec::new();

    f.read_to_end(&mut buffer)?;

    let xlsx = export_to_custom_xlsx(&buffer)?;

    f = File::create("final.xlsx")?;
    f.write(&xlsx)?;

    Ok(())
}
fn main() {
    let now = Instant::now();
    match convert_csv_to_xlsx("original.csv") {
        Ok(_) => (),
        Err(e) => panic!("{e}"),
    }
    println!("[convert_csv_to_xlsx] Took: {:.2?}", now.elapsed());
}
