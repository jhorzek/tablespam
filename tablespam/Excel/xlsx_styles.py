"""Styling options for tables exported to excel."""

from __future__ import annotations
from typing import Callable, cast, Literal, Optional
from dataclasses import dataclass, field
import openpyxl as opy
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_interval
from openpyxl.cell.cell import Cell
import polars as pl
from copy import copy


def set_region_style(
    sheet: Worksheet,
    style: Callable[[Cell], None],
    start_row: int | None,
    start_col: int | None,
    end_row: int | None,
    end_col: int | None,
) -> None:
    """Apply a style to a range of cells.

    Args:
        sheet (Worksheet): Worksheet to which the style is applied.
        style (Callable[[Cell], None]): Function defining the style
        start_row (int): row index at which the style should start
        start_col (int): column index at which the style should start
        end_row (int): row index at which the style should end
        end_col (int): column index at which the style should end
    """
    if any([x is None for x in [start_row, start_col, end_row, end_col]]):
        raise ValueError('One of the locations is None.')
    cols = get_column_interval(cast(int, start_col), cast(int, end_col))
    for row in range(cast(int, start_row), cast(int, end_row) + 1):
        for col in cols:
            cell = col + str(row)
            style(sheet[cell])


@dataclass
class DataStyle:
    """Data styles are styles that are applied to all columns of a specific type.

    Each DataStyle is a combination of a test and a style.

    The test is a function
    that is applied to the data column. It should check if the column is of a specific type
    and return either True or False.

    The style is a function that is applied to a single cell in an openpyxl workbook and
    adds styling to that cell.

    Example:
    >>> import polars as pl
    >>> from tablespam.Excel.xlsx_styles import DataStyle
    >>> # Define a test that checks if a single data column is of type double:
    >>> def test_double(x: pl.DataFrame):
    ...     if len(x.columns) != 1:
    ...         raise ValueError('Multiple columns passed to test.')
    ...     return all([tp in [pl.Float32, pl.Float64] for tp in x.dtypes])
    >>> style = DataStyle(
    ...     test=test_double, style=lambda c: setattr(c, 'number_format', '0.00')
    ... )
    """

    test: Callable[[pl.DataFrame], bool]
    style: Callable[[Cell], None]


@dataclass
class CellStyle:
    """Cell styles are styles that are applied to specific cells in the data.

    A cell style is defined by a list of row indexed, a list of column names, and a style. The style
    is a function that formats single cells of an openpyxl workbook.

    Example:
    >>> from tablespam.Excel.xlsx_styles import CellStyle
    >>> style = CellStyle(
    ...     rows=[1, 2],
    ...     cols=['column_1'],
    ...     style=lambda c: setattr(c, 'number_format', '0.00'),
    ... )
    """

    rows: list[int]
    cols: list[str]
    style: Callable[[Cell], None]


def default_data_styles() -> dict[str, DataStyle]:
    """Defines the default styles that are applied to different data types.

    Returns:
        dict[str, DataStyle]: dict with default styles.
    """

    def test_double(x: pl.DataFrame) -> bool:
        if len(x.columns) != 1:
            raise ValueError('Multiple columns passed to test.')
        return all([tp in [pl.Float32, pl.Float64] for tp in x.dtypes])

    return {
        'double': DataStyle(
            test=test_double,
            style=lambda c: setattr(c, 'number_format', '0.00'),
        ),
    }


BorderStyle = Literal[
    'dashDot',
    'dashDotDot',
    'dashed',
    'dotted',
    'double',
    'hair',
    'medium',
    'mediumDashDot',
    'mediumDashDotDot',
    'mediumDashed',
    'slantDashDot',
    'thick',
    'thin',
    'none',
]


def set_border(
    c: Cell,
    color: str,
    left: Optional[BorderStyle] = None,
    right: Optional[BorderStyle] = None,
    top: Optional[BorderStyle] = None,
    bottom: Optional[BorderStyle] = None,
) -> None:
    """Adds a border to a cell while retaining existing borders.

    Args:
        c (Cell): Cell to which the border should be added
        color (str): Color of the border.
        left (None | str, optional): style (thin, thick, ...) of the left border. Defaults to None.
        right (None | str, optional): style (thin, thick, ...) of the right border. Defaults to None.
        top (None | str, optional): style (thin, thick, ...) of the top border. Defaults to None.
        bottom (None | str, optional): style (thin, thick, ...) of the bottom border. Defaults to None.
    """
    # Define new border
    border = opy.styles.borders.Border(
        left=opy.styles.borders.Side(style=left, color=color)
        if left
        else c.border.left,
        right=opy.styles.borders.Side(style=right, color=color)
        if right
        else c.border.right,
        top=opy.styles.borders.Side(style=top, color=color) if top else c.border.top,
        bottom=opy.styles.borders.Side(style=bottom, color=color)
        if bottom
        else c.border.bottom,
    )
    c.border = border


# Helper functions
def default_bg_style(cell: Cell) -> None:
    """Default background style.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.fill = opy.styles.PatternFill(start_color='FFFFFF', fill_type='solid')


def vline_style(cell: Cell) -> None:
    """Default vertical line style.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    set_border(cell, color='FF000000', left='thin')


def hline_style(cell: Cell) -> None:
    """Default horizontal line style.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    set_border(cell, color='FF000000', top='thin')


def cell_title_style(cell: Cell) -> None:
    """Default title cell style.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=14, bold=True)


def cell_subtitle_style(cell: Cell) -> None:
    """Default subtitle style.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11, bold=True)


def cell_header_lhs_style(cell: Cell) -> None:
    """Default style applied to left hand side of the table header.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11, bold=True)
    cell.border = opy.styles.borders.Border(
        left=opy.styles.borders.Side(style='thin', color='FF000000'),
        bottom=opy.styles.borders.Side(style='thin', color='FF000000'),
        right=opy.styles.borders.Side(style='thin', color='FF000000'),
    )


def cell_header_rhs_style(cell: Cell) -> None:
    """Default style applied to right hand side of the table header.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11, bold=True)
    cell.border = opy.styles.borders.Border(
        left=opy.styles.borders.Side(style='thin', color='FF000000'),
        bottom=opy.styles.borders.Side(style='thin', color='FF000000'),
        right=opy.styles.borders.Side(style='thin', color='FF000000'),
    )


def cell_rownames_style(cell: Cell) -> None:
    """Default style applied to rowname cells.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11)


def cell_data_style(cell: Cell) -> None:
    """Default style applied to data cells.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11)


def cell_footnote_style(cell: Cell) -> None:
    """Default style applied to footnote cells.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11)
    cell.alignment = opy.styles.alignment.Alignment(horizontal='left')


def merged_rownames_style(cell: Cell) -> None:
    """Default style applied to merged row name cells.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.alignment = opy.styles.alignment.Alignment(vertical='top')


def footnote_style(cell: Cell) -> None:
    """Default style applied to footnote.

    Args:
        cell (Cell): Cell reference to which the style is applied
    """
    cell.font = opy.styles.Font(size=11, bold=True)
    cell.alignment = opy.styles.alignment.Alignment(horizontal='left')


@dataclass
class XlsxStyles:
    """Defines styles for different elements of the table.

    Each style element is a function that takes in a single cell of the openpyxl workbook
    and apply a style to that cell.

    The following elements can be adapted:
        merge_rownames (bool): Should adjacent rows with identical names be merged?
        merged_rownames_style (Callable[[Cell], None]): style applied to the merged rownames
        footnote_style (Callable[[Cell], None]): style applied to the table footnote
        data_styles (Callable[[Cell], None]): styles applied to the columns in the data set based on their classes (e.g., numeric, character, etc.). data_styles must be a dict of DataStyle. Note that styles will be applied in the
            order of the list, meaning that a later style may overwrite an earlier style.
        cell_styles (list[CellStyle]): an optional list with styles for selected cells in the data frame.
        bg_default (Callable[[Cell], None]): default color for the background of the table
        bg_title (Callable[[Cell], None]): background color for the title
        bg_subtitle (Callable[[Cell], None]): background color for the subtitle
        bg_header_lhs (Callable[[Cell], None]): background color for the left hand side of the table header
        bg_header_rhs (Callable[[Cell], None]): background color for the right hand side of the table header
        bg_rownames (Callable[[Cell], None]): background color for the row names
        bg_data (Callable[[Cell], None]): background color for the data
        bg_footnote (Callable[[Cell], None]): background color for the footnote
        vline (Callable[[Cell], None]): styling for all vertical lines added to the table
        hline (Callable[[Cell], None]): styling for all horizontal lines added to the table
        cell_default (Callable[[Cell], None]): default style added to cells in the table
        cell_title (Callable[[Cell], None]): style added to title cells in the table
        cell_subtitle (Callable[[Cell], None]): style added to subtitle cells in the table
        cell_header_lhs (Callable[[Cell], None]): style added to the left hand side of the header cells in the table
        cell_header_rhs (Callable[[Cell], None]): style added to the right hand side of the header cells in the table
        cell_rownames (Callable[[Cell], None]): style added to row name cells in the table
        cell_data (Callable[[Cell], None]): style added to data cells in the table
        cell_footnote (Callable[[Cell], None]): style added to footnote cells in the table
    """

    bg_default: Callable[[Cell], None] = field(default=default_bg_style)
    bg_title: Callable[[Cell], None] = field(default=default_bg_style)
    bg_subtitle: Callable[[Cell], None] = field(default=default_bg_style)
    bg_header_lhs: Callable[[Cell], None] = field(default=default_bg_style)
    bg_header_rhs: Callable[[Cell], None] = field(default=default_bg_style)
    bg_rownames: Callable[[Cell], None] = field(default=default_bg_style)
    bg_data: Callable[[Cell], None] = field(default=default_bg_style)
    bg_footnote: Callable[[Cell], None] = field(default=default_bg_style)

    vline: Callable[[Cell], None] = field(default=vline_style)
    hline: Callable[[Cell], None] = field(default=hline_style)

    cell_title: Callable[[Cell], None] = field(default=cell_title_style)
    cell_subtitle: Callable[[Cell], None] = field(default=cell_subtitle_style)
    cell_header_lhs: Callable[[Cell], None] = field(default=cell_header_lhs_style)
    cell_header_rhs: Callable[[Cell], None] = field(default=cell_header_rhs_style)
    cell_rownames: Callable[[Cell], None] = field(default=cell_rownames_style)
    cell_data: Callable[[Cell], None] = field(default=cell_data_style)
    cell_footnote: Callable[[Cell], None] = field(default=cell_footnote_style)

    merge_rownames: bool = True
    merged_rownames_style: Callable[[Cell], None] = field(default=merged_rownames_style)

    footnote_style: Callable[[Cell], None] = field(default=footnote_style)

    data_styles: dict[str, DataStyle] = field(default_factory=default_data_styles)
    cell_styles: None | list[CellStyle] = None
