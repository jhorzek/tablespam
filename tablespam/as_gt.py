from __future__ import annotations
from typing import TYPE_CHECKING, cast, Callable, Protocol, Any

if TYPE_CHECKING:
    from tablespam.TableSpam import TableSpam
    from tablespam.Entry import HeaderEntry

import great_tables as gt
import polars as pl
from dataclasses import dataclass


@dataclass
class FlattenedEntry:
    label: str
    id: str
    level: int
    children: list[str]
    children_ids: list[str]
    children_items: list[str]


def add_gt_spanners(gt_tbl: gt.GT, tbl: TableSpam) -> gt.GT:
    flattened_tbl = flatten_table(tbl)

    if flattened_tbl["flattened_lhs"] is not None:
        gt_tbl = add_gt_spanner_partial(
            gt_tbl=gt_tbl, tbl_partial=flattened_tbl["flattened_lhs"]
        )

    if flattened_tbl["flattened_rhs"] is not None:
        gt_tbl = add_gt_spanner_partial(
            gt_tbl=gt_tbl, tbl_partial=flattened_tbl["flattened_rhs"]
        )

    return gt_tbl


def add_gt_spanner_partial(gt_tbl: gt.GT, tbl_partial: list[FlattenedEntry]) -> gt.GT:
    # The table spanners need to be added in the correct order. All children of
    # a spanner must already be in the table, otherwise we get an error.
    # The level tells us the order; we have to start with the lowest one
    if tbl_partial is None:
        raise ValueError("tbl_partial should not be None.")
    levels = list(set([tbl_part.level for tbl_part in tbl_partial]))
    levels.sort()

    # Next, we iterate over the levels and add them to the gt:
    for level in levels:
        for parent in tbl_partial:
            parent_name = parent.label

            if parent.level == level:

                tbl_data = cast(pl.DataFrame, gt_tbl._tbl_data)
                assert isinstance(gt_tbl._tbl_data, pl.DataFrame)
                item_names = [
                    item for item in parent.children_items if item in tbl_data.columns
                ]
                spanner_ids = [
                    item[1]
                    for item in zip(parent.children_items, parent.children_ids)
                    if item[0] not in tbl_data.columns
                ]

                # if we are at the base level, we do not add a spanner:
                if parent_name != "_BASE_LEVEL_":
                    assert isinstance(parent_name, str)  # required for type checking
                    assert isinstance(parent.id, str)  # required for type checking
                    gt_tbl = gt_tbl.tab_spanner(
                        label=parent_name,
                        id=parent.id,
                        columns=item_names,
                        spanners=spanner_ids,
                    )

                # If children_items and children don't match, we also need to rename elements
                needs_renaming = [
                    item
                    for item in zip(parent.children_items, parent.children)
                    if item[0] != item[1]
                ]

                if len(needs_renaming) > 0:
                    for rename in needs_renaming:
                        old_name: str = rename[0]
                        new_name: str = rename[1]
                        gt_tbl = gt_tbl.cols_label(cases=None, **{old_name: new_name})

    return gt_tbl


def flatten_table(tbl: TableSpam) -> dict[str, None | list[FlattenedEntry]]:
    if tbl.header["lhs"] is not None:
        flattened_lhs = flatten_table_partial(tbl_partial=tbl.header["lhs"])
    else:
        flattened_lhs = None

    flattened_rhs = flatten_table_partial(tbl_partial=tbl.header["rhs"])

    return {"flattened_lhs": flattened_lhs, "flattened_rhs": flattened_rhs}


def flatten_table_partial(
    tbl_partial: HeaderEntry,
    id: str = "",
    flattened: None | list[FlattenedEntry] = None,
) -> None | list[FlattenedEntry]:
    if flattened is None:
        flattened = []

    if len(tbl_partial.entries) != 0:
        flattened_entry = FlattenedEntry(
            label=tbl_partial.name,
            id=f"{id}_{tbl_partial.name}",
            level=tbl_partial.level,
            children=[entry.name for entry in tbl_partial.entries],
            children_ids=[
                f"{id}_{tbl_partial.name}_{entry.name}" for entry in tbl_partial.entries
            ],
            # For items, tablespan can store a name that is different from the actual item label to allow for renaming
            children_items=[
                entry.item_name if entry.item_name is not None else entry.name
                for entry in tbl_partial.entries
            ],
        )
        flattened.append(flattened_entry)

        for entry in tbl_partial.entries:
            flattened = flatten_table_partial(
                tbl_partial=entry, id=f"{id}_{tbl_partial.name}", flattened=flattened
            )

    return flattened


class FormattingFunction(Protocol):
    """Provides a framework for the definition of functions that are used
    to format the great table. The function must accept a gt.GT as first
    arument. All other arguments should be set with args and kwargs. See
    default_formatting for an example.
    """

    def __call__(self, gt_tbl: gt.GT, *args: Any, **kwargs: Any) -> gt.GT: ...


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

    gt_tbl.sub_missing(missing_text="")
    return gt_tbl


def add_gt_rowname_separator(
    gt_tbl: gt.GT, right_of: str, separator_style: gt.style.borders
) -> gt.GT:
    gt_tbl = gt_tbl.tab_style(
        style=separator_style, locations=gt.loc.body(columns=right_of)
    )
    return gt_tbl


def add_gt_titles(gt_tbl: gt.GT, title: str | None, subtitle: str | None) -> gt.GT:
    if title is not None:
        gt_tbl = gt_tbl.tab_header(
            title=title, subtitle=subtitle
        ).opt_align_table_header(align="left")
    return gt_tbl


def add_gt_footnote(gt_tbl: gt.GT, footnote: str) -> gt.GT:
    gt_tbl = gt_tbl.tab_source_note(footnote)
    return gt_tbl
