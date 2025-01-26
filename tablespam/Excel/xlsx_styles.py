"""Styling options for tables exported to excel."""

from __future__ import annotations
from typing import Callable
import tablespam.Excel._as_excel.styles as sty
from dataclasses import dataclass, field
import openpyxl as opy
from openpyxl.cell.cell import Cell
import polars as pl
from functools import partial


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


@dataclass
class XlsxStyles:
    """Defines styles for different elements of the table.

    Each style element is a function that takes in a single cell of the openpyxl workbook
    and apply a style to that cell.

    Args:
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

    bg_default: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_title: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_subtitle: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_header_lhs: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_header_rhs: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_rownames: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_data: Callable[[Cell], None] = field(default=sty.default_bg_style)
    bg_footnote: Callable[[Cell], None] = field(default=sty.default_bg_style)

    vline: Callable[[Cell], None] = field(default=sty.vline_style)
    hline: Callable[[Cell], None] = field(default=sty.hline_style)

    cell_title: Callable[[Cell], None] = field(default=sty.cell_title_style)
    cell_subtitle: Callable[[Cell], None] = field(default=sty.cell_subtitle_style)
    cell_header_lhs: Callable[[Cell], None] = field(default=sty.cell_header_lhs_style)
    cell_header_rhs: Callable[[Cell], None] = field(default=sty.cell_header_rhs_style)
    cell_rownames: Callable[[Cell], None] = field(default=sty.cell_rownames_style)
    cell_data: Callable[[Cell], None] = field(default=sty.cell_data_style)
    cell_footnote: Callable[[Cell], None] = field(default=sty.cell_footnote_style)

    merge_rownames: bool = True
    merged_rownames_style: Callable[[Cell], None] = field(
        default=sty.merged_rownames_style
    )

    footnote_style: Callable[[Cell], None] = field(default=sty.footnote_style)

    data_styles: dict[str, DataStyle] = field(default_factory=default_data_styles)
    cell_styles: None | list[CellStyle] = None


def style_color(primary_color: str = 'ffffff') -> XlsxStyles:
    """Provides a simple way to define a color scheme for tables.

    By default, tables have a "light" theme, where the background is white and text / lines are black.
    Based on a primary color, style_color will create tables that use the primary
    color as background for all title, header, and row name cells and adapts the
    text color based on the primary color. The automatic adaption of the
    background color is implemented based on Mark Ransom and SudoPlz at
    <https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color>


    Args:
        primary_color (str, optional): olor to be used for the title, header, and row names
            background. This must be a hex code for the color. Defaults to 'ffffff'.

    Returns:
        XlsxStyles: Style object
    """
    # We only need the hex code; remove #
    primary_color = primary_color.lstrip('#')
    text_color = sty.get_text_color(primary_color=primary_color)

    if text_color == '000000':
        line_color = '000000'
    else:
        line_color = primary_color

    def bg_fun(cell: Cell, color: str) -> None:
        """Default background style.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.fill = opy.styles.PatternFill(start_color=color, fill_type='solid')

    def vline_color(cell: Cell, color: str) -> None:
        """Default vertical line style.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to line
        """
        sty.set_border(cell, color=color, left='thin')

    def hline_color(cell: Cell, color: str) -> None:
        """Default horizontal line style.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to line
        """
        sty.set_border(cell, color=color, top='thin')

    def cell_title_style(cell: Cell, color: str) -> None:
        """Default title cell style.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.font = opy.styles.Font(size=14, bold=True, color=color)

    def cell_subtitle_style(cell: Cell, color: str) -> None:
        """Default subtitle style.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.font = opy.styles.Font(size=11, bold=True, color=color)

    def cell_header_lhs_style(cell: Cell, color: str) -> None:
        """Default style applied to left hand side of the table header.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.font = opy.styles.Font(size=11, bold=True, color=color)
        cell.border = opy.styles.borders.Border(
            left=opy.styles.borders.Side(style='thin', color=color),
            bottom=opy.styles.borders.Side(style='thin', color=color),
            right=opy.styles.borders.Side(style='thin', color=color),
        )

    def cell_header_rhs_style(cell: Cell, color: str) -> None:
        """Default style applied to right hand side of the table header.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.font = opy.styles.Font(size=11, bold=True, color=color)
        cell.border = opy.styles.borders.Border(
            left=opy.styles.borders.Side(style='thin', color=color),
            bottom=opy.styles.borders.Side(style='thin', color=color),
            right=opy.styles.borders.Side(style='thin', color=color),
        )

    def cell_rownames_style(cell: Cell, color: str) -> None:
        """Default style applied to rowname cells.

        Args:
            cell (Cell): Cell reference to which the style is applied
            color (str): Color applied to cell
        """
        cell.font = opy.styles.Font(size=11, color=color)

    styles = XlsxStyles(
        bg_default=partial(bg_fun, color='ffffff'),
        bg_title=partial(bg_fun, color=primary_color),
        bg_subtitle=partial(bg_fun, color=primary_color),
        bg_header_lhs=partial(bg_fun, color=primary_color),
        bg_header_rhs=partial(bg_fun, color=primary_color),
        bg_rownames=partial(bg_fun, color=primary_color),
        bg_data=partial(bg_fun, color='ffffff'),
        bg_footnote=partial(bg_fun, color='ffffff'),
        vline=partial(vline_color, color=line_color),
        hline=partial(hline_color, color=line_color),
        cell_title=partial(cell_title_style, color=text_color),
        cell_subtitle=partial(cell_subtitle_style, color=text_color),
        cell_header_lhs=partial(cell_header_lhs_style, color=text_color),
        cell_header_rhs=partial(cell_header_rhs_style, color=text_color),
        cell_rownames=partial(cell_rownames_style, color=text_color),
    )
    return styles
