# excel-rs

An *extremely* fast set of Rust and Python utilities to efficiently convert CSVs to Excel XLSX files.

This library is available as a CLI tool and Python PIP package.

This library was created with the goal of being simple, lightweight, and *extremely* performant. As such, many features such as Excel formatting is not currently supported. This library gives you the quickest possible way to convert a `.csv` file to `.xlsx`.

The Python utilities also gives you the quickest possible way to export a Pandas DataFrame, Numpy 2D array or Postgres database as a `.xlsx` file.

## Python

### Installing
```bash
$ pip install py-excel-rs 
```

### Convert a pandas DataFrame to Excel:
```python
import pandas as pd
from py_excel_rs import df_to_xlsx

df = pd.read_csv("file.csv")
xlsx = df_to_xlsx(df)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

### Convert a `csv` file to Excel:
```python
from py_excel_rs import csv_to_xlsx

f = open('file.csv', 'rb')

file_bytes = f.read()
xlsx = csv_to_xlsx(file_bytes)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

### Convert Postgres response to Excel:
```python
import py_excel_rs

conn_string = "dbname=* user=* password=* host=*"
query = "SELECT * FROM table_name"
xlsx = py_excel_rs.pg_to_xlsx(query, conn_string, disable_strict_ssl=False)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

### Build Postgres Query to Excel:
```python
from py_excel_rs import ExcelPostgresBuilder, OrderBy

conn_string = "dbname=* user=* password=* host=*"
builder = ExcelPostgresBuilder(conn_str=conn_string, table_name="my_schema.my_table")

xlsx = builder.select_all()
    .exclude(["Unwanted_Column1", "Unwanted_Column2"])
    .orderBy("Usernames", OrderBy.ASCENDING)
    .execute()

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

## Command Line Tool
To install, download the latest release of `cli-excel-rs` for your platform from Github Releases [here](https://github.com/carlvoller/excel-rs/releases?q=cli-excel-rs&expanded=true).
```bash
$ wget https://github.com/carlvoller/excel-rs/releases/download/cli-0.2.0/excel-rs-linux-aarch64.zip
$ unzip excel-rs-linux-aarch64.zip
$ chmod +x ./cli-excel-rs
```
Then simply run the binary:
```bash
$ ./cli-excel-rs csv --in my_csv.csv --out my_excel.xlsx
```

If you would like the build the binary yourself, you can do so using these commands:
```bash
$ git clone https://github.com/carlvoller/excel-rs
$ cargo build --release
$ ./target/release/cli-excel-rs csv --in my_csv.csv --out my_excel.xlsx
```

## Rust
TODO: Add rust documentation

## Benchmarks
With a focus on squeezing out as much performance as possible, **py-excel-rs** is up to **45.5x** faster than `pandas` and **12.5x** faster than the fastest `xlsx` writer on pip.

**cli-excel-rs** also managed to out perform [csv2xlsx](https://github.com/mentax/csv2xlsx?tab=readme-ov-file), the most poopular csv to xlsx tool. It is up to **12x** faster given the same dataset.

These tests used a sample dataset from [DataBlist](https://www.datablist.com/learn/csv/download-sample-csv-files) that contained 1,000,000 rows and 9 columns.

Tests were conducted on an Macbook Pro M1 Max with 64GB of RAM

### Python 

#### py-excel-rs (2.89s)
```bash
$ time python test-py-excel-rs.py
python3 test-py-excel-rs.py  2.00s user 0.18s system 99% cpu 2.892 total
```

#### openpyxl (97.38s)
```bash
$ time python test-openpyxl.py
python3 test-openpyxl.py  94.48s user 2.39s system 99% cpu 1:37.38 total
```

#### pandas `to_excel()` (131.24s)
```bash
$ time python test-pandas.py
python3 test-pandas.py  127.99s user 2.75s system 99% cpu 2:11.24 total
```

#### pandas `to_excel(engine="xlsxwriter")` (82.29s)
```bash
$ time python test-pandas-xlsxwriter.py
python3 test-pandas-xlsxwriter.py  76.86s user 1.95s system 95% cpu 1:22.29 total
```

#### xlsxwriter (42.543s)
```bash
$ time python test-xlsxwriter.py
python3 test-xlsxwriter.py  41.58s user 0.81s system 99% cpu 42.543 total
```

#### pyexcelerate (35.821s)
```bash
$ time python test-pyexcelerate.py
python3 test-pyexcelerate.py  35.27s user 0.33s system 99% cpu 35.821 total
```

### Command Line Tools

#### cli-excel-rs (2.756s)
```bash
$ time ./cli-excel-rs csv --in organizations-1000000.csv --out results.xlsx
./cli-excel-rs csv --in organizations-1000000.csv --out 2.69s user 0.07s system 99% cpu 2.756 total
```

#### [csv2xlsx](https://github.com/mentax/csv2xlsx?tab=readme-ov-file)  (33.74s)
```bash
$ time ./csv2xlsx --output results.xlsx organizations-1000000.csv
./csv2xlsx --output results.xlsx organizations-1000000.csv  57.63s user 1.62s system 175% cpu 33.740 total
```

### Rust

TODO: Add Rust Benchmark comparisons to rust_xlsxwriter, etc.