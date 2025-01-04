"""Functions to print the TableSpam to the console."""

from __future__ import annotations  # noqa: D100
from typing import TYPE_CHECKING, Any
import numpy as np
import polars as pl

if TYPE_CHECKING:
    from tablespam.TableSpam import TableSpam
    from tablespam.Entry import HeaderEntry


def tbl_as_string(
    tbl: TableSpam, digits: int = 2, n: int = 3, max_char: int = 30
) -> str:
    """Translates a TableSpam to a string.

    Args:
        tbl (TableSpam): TableSpam table
        digits (int, optional): Number of digits to round floats to. Defaults to 2.
        n (int, optional): number of rows from the data set to print. Defaults to 3.
        max_char (int, optional): number of characters that each cell at maximum is allows to have. Defaults to 30.

    Returns:
        str: String describing the table
    """

    if tbl.table_data['col_data'] is None:
        raise ValueError("tbl.table_data['col_data'] should not be None.")

    if tbl.header['lhs'] is not None:
        max_level = max(tbl.header['lhs'].level, tbl.header['rhs'].level)
        max_col = tbl.header['lhs'].width + tbl.header['rhs'].width
    else:
        max_level = tbl.header['rhs'].level
        max_col = tbl.header['rhs'].width

    header_table = np.full(
        (
            max_level + min(n, tbl.table_data['col_data'].height),
            max_col + (tbl.header['lhs'] is not None),
        ),
        '',
        dtype=f'<U{max_char}',
    )

    if tbl.header['lhs'] is not None:
        header_table = insert_header_entries(
            header_partial=tbl.header['lhs'],
            max_level=max_level,
            column_offset=1,
            header_table=header_table,
        )
        column_offset = tbl.header['lhs'].width + 2
    else:
        column_offset = 1

    header_table = insert_header_entries(
        header_partial=tbl.header['rhs'],
        max_level=max_level,
        column_offset=column_offset,
        header_table=header_table,
    )

    # add data
    rws = range(max_level, max_level + min(n, tbl.table_data['col_data'].height))

    if tbl.header['lhs'] is not None:
        if tbl.table_data['row_data'] is None:
            raise ValueError("tbl.table_data['row_data'] should not be None.")
        cls = range(0, tbl.table_data['row_data'].width)
        df_transformed = tbl.table_data['row_data'].head(n=n)
        # Round all floats
        for c_n in [
            c
            for c in df_transformed.columns
            if df_transformed[c].dtype in [pl.Float32, pl.Float64]
        ]:
            df_transformed = df_transformed.with_columns(pl.col(c_n).round(digits))
        # Cast all columns to string
        for c_n in df_transformed.columns:
            df_transformed = df_transformed.with_columns(pl.col(c_n).cast(pl.String))

        header_table[np.ix_(rws, cls)] = df_transformed.to_numpy()

        # add vertical line
        header_table[:, max(cls) + 1] = '|'

        cls = range(max(cls) + 2, tbl.table_data['col_data'].width + max(cls) + 2)
    else:
        cls = range(tbl.table_data['col_data'].width)

    df_transformed = tbl.table_data['col_data'].head(n=n)
    # Round all floats
    for c_n in [
        c
        for c in df_transformed.columns
        if df_transformed[c].dtype in [pl.Float32, pl.Float64]
    ]:
        df_transformed = df_transformed.with_columns(pl.col(c_n).round(digits))
    # Cast all columns to string
    for c_n in df_transformed.columns:
        df_transformed = df_transformed.with_columns(pl.col(c_n).cast(pl.String))

    header_table[np.ix_(rws, cls)] = df_transformed.to_numpy()

    # add ...
    if n < tbl.table_data['col_data'].shape[0]:
        header_table = np.vstack(
            [
                header_table,
                np.full(
                    (1, header_table.shape[1]),
                    '...',
                    dtype=f'<U{max_char}',
                ),
            ]
        )

    # Add horizontal line. We need the number of characters in
    # each column
    max_lengths = np.vectorize(len)(header_table).max(axis=0)

    # Create strings of dashes of corresponding lengths
    header_table[max_level - 1, :] = np.array(
        ['-' * length for length in max_lengths],
        dtype=f'<U{max_char}',
    )

    # set the length of each string to be the same
    header_table = np.array(
        [
            [item.ljust(max_lengths[col_idx]) for col_idx, item in enumerate(row)]
            for row in header_table
        ],
        dtype=f'<U{max_char}',
    )

    # add vertical lines to the left and right of the table
    header_table = np.hstack(
        [
            np.hstack([np.full((header_table.shape[0], 1), '|'), header_table]),
            np.full(
                (header_table.shape[0], 1),
                '|',
                dtype=f'<U{max_char}',
            ),
        ]
    )

    # actual printing
    tbl_string = ''
    if tbl.title is not None:
        tbl_string = f'{tbl_string}{tbl.title}\n'
    if tbl.subtitle is not None:
        tbl_string = f'{tbl_string}{tbl.subtitle}\n'

    for row in header_table:
        tbl_string = f"{tbl_string}\n{(' '.join(f'{x}' for x in row))}"

    tbl_string = f'{tbl_string}\n'

    if tbl.footnote is not None:
        tbl_string = f'{tbl_string}{tbl.footnote}\n'

    return tbl_string


def insert_header_entries(
    header_partial: HeaderEntry,
    max_level: int,
    column_offset: int,
    header_table: np.ndarray[tuple[Any, Any], np.dtype[Any]],
) -> np.ndarray[tuple[Any, Any], np.dtype[Any]]:
    """Insert specific entries into the table.

    Args:
        header_partial (HeaderEntry): header entry to insert
        max_level (int): maximal number of levels in the table header
        column_offset (int): offset used to specify which column to write to
        header_table (np.ndarray): array filled recursively

    Returns:
        np.ndarray: array with header
    """
    if header_partial.name != '_BASE_LEVEL_':
        # We have to take the 0-based indexing into account:
        header_table[max_level - header_partial.level - 1, column_offset - 1] = (
            header_partial.name
        )

    if header_partial.entries is not None:
        for i in range(len(header_partial.entries)):
            header_table = insert_header_entries(
                header_partial=header_partial.entries[i],
                max_level=max_level,
                column_offset=column_offset,
                header_table=header_table,
            )
            column_offset = column_offset + header_partial.entries[i].width

    return header_table
