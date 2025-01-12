from __future__ import annotations  # noqa: D100
from typing import Callable
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
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    cols = get_column_interval(start_col, end_col)
    for row in range(start_row, end_row + 1):
        for col in cols:
            cell = col + str(row)
            style(sheet[cell])


@dataclass
class DataStyle:
    test: Callable[[pl.DataFrame], bool]
    style: Callable[[Cell], None]


@dataclass
class CellStyle:
    rows: list[int]
    cols: list[str]
    style: Callable[[Cell], None]


def default_data_styles() -> dict[DataStyle]:
    def test_double(x: pl.DataFrame):
        if len(x.columns) != 1:
            raise ValueError('Multiple columns passed to test.')
        return all([tp in [pl.Float32, pl.Float64] for tp in x.dtypes])

    return {
        'double': DataStyle(
            test=test_double,
            style=lambda c: setattr(c, 'number_format', '0.00'),
        ),
    }


def set_border(
    c: Cell,
    color: str,
    left: None | str = None,
    right: None | str = None,
    top: None | str = None,
    bottom: None | str = None,
):
    border = copy(c.border)
    if left is not None:
        border.left = opy.styles.borders.Side(style=left, color=color)
    if right is not None:
        border.right = opy.styles.borders.Side(style=right, color=color)
    if top is not None:
        border.top = opy.styles.borders.Side(style=top, color=color)
    if bottom is not None:
        border.bottom = opy.styles.borders.Side(style=bottom, color=color)
    c.border = border


@dataclass
class XlsxStyles:
    bg_default: Callable[[Cell], None] = lambda c: setattr(
        c, 'fill', opy.styles.PatternFill(start_color='FFFFFF', fill_type='solid')
    )
    bg_title: Callable[[Cell], None] = bg_default
    bg_subtitle: Callable[[Cell], None] = bg_default
    bg_header_lhs: Callable[[Cell], None] = bg_default
    bg_header_rhs: Callable[[Cell], None] = bg_default
    bg_rownames: Callable[[Cell], None] = bg_default
    bg_data: Callable[[Cell], None] = bg_default
    bg_footnote: Callable[[Cell], None] = bg_default

    vline: Callable[[Cell], None] = lambda c: set_border(
        c=c, color='FF000000', left='thin'
    )

    hline: Callable[[Cell], None] = lambda c: set_border(
        c=c, color='FF000000', top='thin'
    )

    cell_title: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=14, bold=True))
    )
    cell_subtitle: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=11, bold=True))
    )
    cell_header_lhs: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=11, bold=True)),
        setattr(
            c,
            'border',
            opy.styles.borders.Border(
                left=opy.styles.borders.Side(style='thin', color='FF000000'),
                bottom=opy.styles.borders.Side(style='thin', color='FF000000'),
                right=opy.styles.borders.Side(style='thin', color='FF000000'),
            ),
        ),
    )

    cell_header_rhs: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=11, bold=True)),
        setattr(
            c,
            'border',
            opy.styles.borders.Border(
                left=opy.styles.borders.Side(style='thin', color='FF000000'),
                bottom=opy.styles.borders.Side(style='thin', color='FF000000'),
                right=opy.styles.borders.Side(style='thin', color='FF000000'),
            ),
        ),
    )

    cell_rownames: Callable[[Cell], None] = lambda c: setattr(
        c, 'font', opy.styles.Font(size=11)
    )
    cell_data: Callable[[Cell], None] = lambda c: setattr(
        c, 'font', opy.styles.Font(size=11)
    )
    cell_footnote: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=11)),
        setattr(c, 'alignment', opy.styles.alignment.Alignment(horizontal='left')),
    )

    merge_rownames: bool = True
    merged_rownames_style: Callable[[Cell], None] = lambda c: setattr(
        c, 'alignment', opy.styles.alignment.Alignment(vertical='top')
    )

    footnote_style: Callable[[Cell], None] = lambda c: (
        setattr(c, 'font', opy.styles.Font(size=11, bold=True)),
        setattr(c, 'alignment', opy.styles.alignment.Alignment(horizontal='left')),
    )
    data_styles: dict[DataStyle] = field(default_factory=default_data_styles)
    cell_styles: None | list[CellStyle] = None
