from typing import Callable  # noqa: D100
import polars as pl  # noqa: D100
import openpyxl as opy
from openpyxl.utils import get_column_interval
from openpyxl.cell.cell import Cell
from tablespam.Excel.xlsx_styles import DataStyle


def write_excel_col(
    workbook: opy.Workbook,
    sheet: str,
    data: pl.DataFrame,
    row_start: int,
    col_start: int,
    base_style: Callable[[Cell], None],
    data_styles: dict[str, DataStyle],
):
    style = None
    for data_style in data_styles:
        if data_styles[data_style].test(data):
            style = data_styles[data_style].style
            break

    col = get_column_interval(start=col_start, end=col_start)[0]
    for row in range(row_start, row_start + data.shape[0]):
        workbook[sheet][col + str(row)] = data[row - row_start, 0]
        # we first apply the base style and then add/replace type specific styles:
        base_style(workbook[sheet][col + str(row)])
        if style is not None:
            style(workbook[sheet][col + str(row)])
