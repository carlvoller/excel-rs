import pandas as pd
import numpy as np

from py_excel_rs import _excel_rs

def csv_to_xlsx(buf: bytes) -> bytes:
    return _excel_rs.export_to_xlsx(buf)

def df_to_xlsx(df: pd.DataFrame) -> bytes:
    py_list = np.vstack((df.keys().to_numpy(), df.to_numpy(dtype='object')))
    return _excel_rs.py_2d_to_xlsx(py_list)

def pg_to_xlsx(query: str, conn_string: str, disable_strict_ssl=False) -> bytes:
    return _excel_rs.pg_query_to_xlsx(query, conn_string, disable_strict_ssl)