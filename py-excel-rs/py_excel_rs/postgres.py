from enum import Enum
from typing import Optional

from py_excel_rs import _excel_rs


class OrderBy(Enum):
    ASCENDING = "ASC"
    DESCENDING = "DESC"


class ExcelPostgresBuilder:
    _conn_str: str
    _selected: str
    _excluded: Optional[list[str]]
    _table_name: str
    _order_by: Optional[OrderBy]
    _order_by_col: Optional[str]
    _consumed: bool

    def __init__(self, conn_str: str, table_name: str):
        if (
            not conn_str
            or not table_name
            or not isinstance(conn_str, str)
            or not isinstance(table_name, str)
        ):
            raise ValueError("missing or invalid type for conn_str or table_name")

        self._consumed = False
        self._conn_str = conn_str
        self._table_name = table_name
        self._excluded = None
        self._order_by = None
        self._order_by_col = None

    def select_all(self):
        if self._consumed:
            raise RuntimeError("Cannot modify PostgresBuilder after execute()")
        self._selected = "*"
        return self

    def select(self, columns: list[str]):
        if self._consumed:
            raise RuntimeError("Cannot modify PostgresBuilder after execute()")
        self._selected = ", ".join(columns)
        return self

    def exclude(self, columns: Optional[list[str]]):
        if self._consumed:
            raise RuntimeError("Cannot modify PostgresBuilder after execute()")
        self._excluded = columns
        return self

    def order_by(self, col: Optional[str], order: Optional[OrderBy]):
        if self._consumed:
            raise RuntimeError("Cannot modify PostgresBuilder after execute()")
        self._order_by = order
        self._order_by_col = col
        return self

    def execute(self):
        if self._consumed:
            raise RuntimeError("Cannot execute PostgresBuilder after execute()")

        if self._selected is None:
            raise ValueError(
                "PostgresBuilder requires select_all() or select() to be ran once before execute()"
            )
            
        client = _excel_rs.PyPostgresClient.new(self._conn_str)
        (schema_name, table_name) = self._table_name.split(".")
        if not table_name:
            table_name = schema_name
            schema_name = ""
        
        if self._selected == "*":
            if self._excluded is not None:
                columns = [f"\"{x}\"" for x in client.get_columns(table_name, schema_name, self._excluded)]
                query = f"SELECT {', '.join(columns)} FROM {self._table_name}"
            else:
                query = f"SELECT * FROM {self._table_name}"
        else:
            parsed =  [f"'{x}'" for x in self._selected]
            query = f"SELECT {', '.join(parsed)} FROM {self._table_name}"
            
        if self._order_by is not None and self._order_by_col is not None:
            query += f" ORDER BY \"{self._order_by_col}\" {self._order_by.value}"
        
        xlsx = client.get_xlsx_from_query(query)
        client.close()
        return xlsx

