import py_excel_rs
import datetime
import pandas as pd

# f = open('organizations-1000000.csv', 'rb')
# xlsx = py_excel_rs.csv_to_xlsx(f.read())


data = [[datetime.datetime.now(), "hello", 10, 10.888]]
df = pd.DataFrame(data, columns=["Date", "hi", "number1", "float2"])

xlsx = py_excel_rs.df_to_xlsx(df, should_infer_types=True)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)