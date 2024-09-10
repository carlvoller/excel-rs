use std::{fs::File, io::{Cursor, Read, Write}};

use clap::{arg, Command};
use excel_rs_csv::{bytes_to_csv, get_headers, get_next_record};
use excel_rs_xlsx::WorkBook;

fn cli() -> Command {
    Command::new("excel-rs")
        .about("A collection of tools to work with XLSX files")
        .subcommand_required(true)
        .arg_required_else_help(true)
        .subcommand(
            Command::new("csv")
                .about("Convert a csv file to xlsx")
                .arg(arg!(--csv <FILE> "csv file to convert"))
                .arg(arg!(--out <FILE> "xlsx output file name")),
        )
}

fn main() {
    let matches = cli().get_matches();

    match matches.subcommand() {
        Some(("csv", sub_matches)) => {
            let input = sub_matches.get_one::<String>("csv").expect("required");
            let out = sub_matches.get_one::<String>("out").expect("required");

            let mut f = File::open(input).expect("input csv file not found");
            let mut data: Vec<u8> = Vec::new();

            f.read_to_end(&mut data).expect(&format!("Unable to read file {input}"));

            let output_buffer = vec![];
            let mut workbook = WorkBook::new(Cursor::new(output_buffer));
            let mut worksheet = workbook.get_worksheet(String::from("Sheet 1"));

            let mut reader = bytes_to_csv(data.as_slice());
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

            let final_buffer = workbook.finish().ok().unwrap().into_inner();

            f = File::create(out).expect(&format!("unable to write to {out}"));
            f.write(&final_buffer).expect(&format!("Failed to write to file {out}"));
        }
        _ => unreachable!("Unsupported subcommand"),
    }
}
