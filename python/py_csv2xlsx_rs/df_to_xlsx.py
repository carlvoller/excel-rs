from io import BytesIO

import pandas as pd

from py_csv2xlsx_rs import _csv2xlsx_rs


def df_to_xlsx(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    return _csv2xlsx_rs.export_to_xlsx(buf.read())
