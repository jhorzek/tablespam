"""Styling options for tables exported to excel."""

from __future__ import annotations
from typing import Callable, cast, Literal, Optional
import openpyxl as opy
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_interval
from openpyxl.cell.cell import Cell


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


def get_text_color(primary_color: str) -> str:
    """Get text color based on background color.

    Determines if the text should be black or white based on the formula
    from Mark Ransom and SudoPlz at
    <https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color>



    Args:
        primary_color (str): color to be used for the title, header, and row names background.

    Raises:
        ValueError: Error in case the color code is not for a hex

    Returns:
        str: hex color code for text
    """
    # Determines if the text should be black or white based on the formula
    # from Mark Ransom and SudoPlz at
    # <https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color>

    # We only need the hex code; remove #
    primary_color = primary_color.lstrip('#')
    if len(primary_color) != 6:
        raise ValueError(
            'Expected primary_color to be of length 6. Got {primary_color} (length {len(primary_color)}).'
        )
    # split in rgb and divide by max (255)
    r = int(primary_color[0:3], 16) / 255
    g = int(primary_color[2:4], 16) / 255
    b = int(primary_color[4:6], 16) / 255

    def convert_color(x: float) -> float:
        if x <= 0.03928:
            return x / 12.92
        else:
            return pow((x + 0.055) / 1.055, 2.4)

    luminance = (
        (0.2126 * convert_color(r))
        + (0.7152 * convert_color(g))
        + (0.0722 * convert_color(b))
    )

    if luminance <= 0.1769:
        return 'ffffff'
    return '000000'
