# csv2xlsx-rs

A set of Rust and Python utilities to efficiently convert CSVs to Excel XLSX files.

Example usage:
```python
import pandas as pd
import py_csv2xlsx_rs

df = pd.read_csv("file.csv")
xlsx = py_csv2xlsx_rs.df_to_xlsx(df)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
```

Python utilities are typed with type hints.