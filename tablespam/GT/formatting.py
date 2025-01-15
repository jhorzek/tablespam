"""Formatting functions to use with great_tables."""
from typing import cast

import great_tables as gt
import polars as pl

def default_formatting(gt_tbl: gt.GT, decimals: int = 2) -> gt.GT:
    """Provides a default formatting for all columns in the great table.

    Args:
        gt_tbl (gt.GT): Great table before formatting
        decimals (int, optional): The number of decimals to round floats to. Defaults to 2.

    Returns:
        gt.GT: Great table after formatting
    """
    tbl_data = cast(pl.DataFrame, gt_tbl._tbl_data)
    for item, data_type in zip(tbl_data.columns, tbl_data.dtypes):
        if data_type in [pl.Float32, pl.Float64]:
            gt_tbl = gt_tbl.fmt_number(columns=[item], decimals=decimals)
        elif data_type in [pl.Date]:
            gt_tbl = gt_tbl.fmt_date(columns=[item])
        elif data_type in [pl.Datetime]:
            gt_tbl = gt_tbl.fmt_datetime(columns=[item])
        elif data_type in [pl.Time]:
            gt_tbl = gt_tbl.fmt_time(columns=[item])

    gt_tbl.sub_missing(missing_text='')
    return gt_tbl
