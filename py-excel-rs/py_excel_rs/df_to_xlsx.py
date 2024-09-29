import pandas as pd
import numpy as np
from enum import Enum

from py_excel_rs import _excel_rs

from pandas.api.types import is_datetime64_any_dtype as is_datetime
from pandas.api.types import is_numeric_dtype as is_numeric

class CellTypes(Enum):
    Date = "n\" s=\"1"
    String = "str"
    Number = "n"
    Formula = "str"
    Boolean = "b"

def csv_to_xlsx(buf: bytes) -> bytes:
    return _excel_rs.csv_to_xlsx(buf)

def df_to_xlsx(df: pd.DataFrame, should_infer_types: bool = False) -> bytes:

    py_list = np.vstack((df.keys().to_numpy(), df.to_numpy(dtype='object')))

    if should_infer_types:
        df_types = []
        for x in df.dtypes:
            if is_datetime(x):
                df_types.append(CellTypes.Date)
            elif is_numeric(x):
                df_types.append(CellTypes.Number)
            else:
                df_types.append(CellTypes.String)
        return _excel_rs.typed_py_2d_to_xlsx(py_list, list(map(lambda x : x.value, df_types)))
    return _excel_rs.py_2d_to_xlsx(py_list)

def pg_to_xlsx(query: str, conn_string: str) -> bytes:
    
    client = _excel_rs.PyPostgresClient.new(conn_string)
    xlsx = client.get_xlsx_from_query(query)
    client.close()
    return xlsx