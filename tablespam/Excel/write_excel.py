"""Helper functions to write data to an excel workbook."""

from typing import Callable
import polars as pl
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
    """Writes a single data column to the Excel workbook.

    Args:
        workbook (opy.Workbook): openpyxl workbook
        sheet (str): name of the sheet to which the table should be added. Defaults to 'Table'.
        data (pl.DataFrame): data frame to add to the table. Should a single column.
        row_start (int): row where the table start will start in the workbook
        col_start (int): column where the table start will start in the workbook
        base_style (Callable[[Cell], None]): style to add to all data cells
        data_styles (dict[str, DataStyle]): style to add to specific data types
    """
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
