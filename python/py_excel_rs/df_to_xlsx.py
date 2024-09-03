from io import BytesIO

import pandas as pd

from py_excel_rs import _excel_rs

def csv_to_xlsx(buf: bytes) -> bytes:
    return _excel_rs.export_to_xlsx(buf)

def df_to_xlsx(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_csv(buf, index=False)

    buf.seek(0)
    return _excel_rs.export_to_xlsx(buf.read())
