use std::{fs::File, io::{Cursor, Read, Write}};

use clap::{arg, Command};
use excel_rs_csv::{bytes_to_csv, get_headers, get_next_record};
use excel_rs_xlsx::{WorkBook, typed_sheet::{TYPE_STRING}};

fn cli() -> Command {
    Command::new("excel-rs")
        .about("A collection of tools to work with XLSX files")
        .subcommand_required(true)
        .arg_required_else_help(true)
        .subcommand(
            Command::new("csv")
                .about("Convert a csv file to xlsx")
                .arg(arg!(--in <FILE> "csv file to convert"))
                .arg(arg!(--out <FILE> "xlsx output file name"))
                .arg(arg!(--filter "Freeze the top row and add auto-filters")),
        )
}

fn main() {
    let matches = cli().get_matches();

    match matches.subcommand() {
        Some(("csv", sub_matches)) => {
            let input = sub_matches.get_one::<String>("in").expect("required");
            let out = sub_matches.get_one::<String>("out").expect("required");

            let apply_filter = sub_matches.get_flag("filter");

            let mut f = File::open(input).expect("input csv file not found");
            let mut data: Vec<u8> = Vec::new();

            f.read_to_end(&mut data).expect(&format!("Unable to read file {input}"));

            let output_buffer = vec![];
            let mut workbook = WorkBook::new(Cursor::new(output_buffer));
            let mut worksheet = workbook.get_typed_worksheet(String::from("Sheet 1"));

            // Apply filters first if requested
            if apply_filter {
                worksheet.freeze_top_row();
                worksheet.add_auto_filter();
            }

            // Initialize the sheet before writing any rows
            worksheet.init_sheet().expect("Failed to initialize worksheet");

            let mut reader = bytes_to_csv(data.as_slice());
            let headers = get_headers(&mut reader);

            // Write headers with string types if present
            if let Some(headers) = headers {
                let headers_to_bytes = headers.iter().to_owned().collect();
                let header_types = vec![TYPE_STRING; headers.len()];
                if let Err(e) = worksheet.write_row(headers_to_bytes, &header_types) {
                    panic!("{e}");
                }
            }

            // Get first data row to infer types
            if let Some(record) = get_next_record(&mut reader) {
                let row_data: Vec<&[u8]> = record.iter().to_owned().collect();
                // Infer types from this row
                let types = worksheet.infer_row_types(&row_data);
                // Write the row using inferred types
                if let Err(e) = worksheet.write_row(row_data, &types) {
                    panic!("{e}");
                }

                // Write remaining rows using the same types
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

            let final_buffer = workbook.finish().ok().unwrap().into_inner();

            f = File::create(out).expect(&format!("unable to write to {out}"));
            f.write(&final_buffer).expect(&format!("Failed to write to file {out}"));
        }
        _ => unreachable!("Unsupported subcommand"),
    }
}
