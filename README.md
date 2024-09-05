# excel-rs

A set of Rust and Python utilities to efficiently convert CSVs to Excel XLSX files.

This library was created with the goal of being simple, lightweight, and *extremely* fast. As such, many features such as Excel formatting is not currently supported. This library gives you the quickest possible way to convert a `.csv` file to `.xlsx`.

## Python

### Installing
```
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
xlsx = py_excel_rs.pg_to_xlsx(query, conn_string)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

## Rust
TODO: Add rust documentation

## Benchmarks
With a focus on squeezing out as much performance as possible, **py-excel-rs** is up to **65.5x** faster than `pandas` and **17.5x** faster than the next fastest `xlsx` writer on pip.

These tests used a sample dataset from [DataBlist](https://www.datablist.com/learn/csv/download-sample-csv-files) that contained 1,000,000 rows and 9 columns.

Tests were conducted on an Macbook Pro M1 Max with 64GB of RAM

### Python 

#### py-excel-rs (2.186s)
```
$ time python test-py-excel-rs.py
python3 test-py-excel-rs.py  2.00s user 0.18s system 99% cpu 2.186 total
```

#### openpyxl (97.38s)
```
$ time python test-openpyxl.py
python3 test-openpyxl.py  94.48s user 2.39s system 99% cpu 1:37.38 total
```

#### pandas `to_excel()` (131.24s)
```
$ time python test-pandas.py
python3 test-pandas.py  127.99s user 2.75s system 99% cpu 2:11.24 total
```

#### pandas `to_excel(engine="xlsxwriter")` (82.29s)
```
$ time python test-pandas-xlsxwriter.py
python3 test-pandas-xlsxwriter.py  76.86s user 1.95s system 95% cpu 1:22.29 total
```

#### xlsxwriter (42.543s)
```
$ time python test-xlsxwriter.py
python3 test-xlsxwriter.py  41.58s user 0.81s system 99% cpu 42.543 total
```

#### pyexcelerate (35.821s)
```
$ time python test-pyexcelerate.py
python3 test-pyexcelerate.py  35.27s user 0.33s system 99% cpu 35.821 total
```


### Rust

TODO: Add Rust Benchmark comparisons to rust_xlsxwriter, etc.