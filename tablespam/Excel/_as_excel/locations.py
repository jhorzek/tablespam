"""Locations for the table elements in the xlsx file."""

from __future__ import annotations
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:
    from tablespam.TableSpam import TableSpam


class Locations:
    """Provides the locations of elements in the Excel table."""

    def __init__(self, tbl: TableSpam, start_row: int, start_col: int):
        """Provides row and column indices for the table elements.

        Args:
            tbl (TableSpam): table created with tablespam
            start_row (int): row index at which the table will start
            start_col (int): column index at which the table will start
        """
        # Initialize row indices
        if tbl.title is not None:
            start_row_title = end_row_title = start_row
            start_row += 1
        else:
            start_row_title = None
            end_row_title = None

        if tbl.subtitle is not None:
            start_row_subtitle = end_row_subtitle = start_row
            start_row += 1
        else:
            start_row_subtitle = None
            end_row_subtitle = None

        start_row_header = start_row

        if tbl.header['lhs'] is not None:
            start_row += max(tbl.header['lhs'].level, tbl.header['rhs'].level) - 1
        else:
            start_row += tbl.header['rhs'].level - 1

        # Previous iteration was last entry of header rows, moving one up is the last row that belongs to the header
        end_row_header = start_row - 1

        # Begin data block at current row (one row below header)
        start_row_data = start_row
        if tbl.table_data['col_data'] is not None:
            end_row_data = start_row + len(tbl.table_data['col_data']) - 1

        # Add one row for every entry in the table
        if tbl.table_data['col_data'] is not None:
            start_row += len(tbl.table_data['col_data'])

        # Footnote is only one row long
        start_row_footnote = end_row_footnote = start_row

        # Initialize column indices
        start_col_title = start_col_subtitle = start_col_header_lhs = (
            start_col_footnote
        ) = start_col
        n_col = 0

        # Get number of columns from `width` key in table header and add to start column
        # Skip lhs if formula had no lhs
        if tbl.header['lhs'] is not None:
            n_col += tbl.header['lhs'].width + tbl.header['rhs'].width - 1
            start_col_header_rhs = start_col + tbl.header['lhs'].width
            end_col_header_lhs = start_col + tbl.header['lhs'].width - 1
        else:
            n_col += tbl.header['rhs'].width - 1
            start_col_header_rhs = start_col
            end_col_header_lhs = None
        end_col_title = end_col_subtitle = end_col_footnote = end_col_header_rhs = (
            start_col + n_col
        )

        self.rows = {
            'start_row_title': start_row_title,
            'end_row_title': end_row_title,
            'start_row_subtitle': start_row_subtitle,
            'end_row_subtitle': end_row_subtitle,
            'start_row_header': start_row_header,
            'end_row_header': end_row_header,
            'start_row_data': start_row_data,
            'end_row_data': end_row_data,
            'start_row_footnote': start_row_footnote,
            'end_row_footnote': end_row_footnote,
        }
        self.cols = {
            'start_col_title': start_col_title,
            'end_col_title': end_col_title,
            'start_col_subtitle': start_col_subtitle,
            'end_col_subtitle': end_col_subtitle,
            'start_col_header_lhs': start_col_header_lhs,
            'end_col_header_lhs': end_col_header_lhs,
            'start_col_header_rhs': start_col_header_rhs,
            'end_col_header_rhs': end_col_header_rhs,
            'start_col_footnote': start_col_footnote,
            'end_col_footnote': end_col_footnote,
        }

    def get_row(self, of: str) -> int:
        """Extract the location of a row element.

        Args:
            of (str): The element for which the location is extracted

        Raises:
            ValueError: Raises error if the location is None

        Returns:
            int: location index
        """
        if self.rows[of] is None:
            raise ValueError(f'Row {of} is None.')

        return cast(int, self.rows[of])

    def get_col(self, of: str) -> int:
        """Extract the location of a column element.

        Args:
            of (str): The element for which the location is extracted

        Raises:
            ValueError: Raises error if the location is None

        Returns:
            int: location index
        """
        if self.cols[of] is None:
            raise ValueError(f'Column {of} is None.')
        return self.cols[of]
