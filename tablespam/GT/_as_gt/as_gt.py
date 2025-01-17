from __future__ import annotations  # noqa: D100
from typing import TYPE_CHECKING, cast, Protocol, Any

if TYPE_CHECKING:
    from tablespam.TableSpam import TableSpam
    from tablespam._Formula.Entry import HeaderEntry

import great_tables as gt
import polars as pl
from dataclasses import dataclass


@dataclass
class FlattenedEntry:
    """FlattenedEntry represents a single entry in a header.

    The header elements in a TableSpam are highly nested within each other. With flatten_table,
    these nested entries can be translated into a single list with flattened entries.

    fields:
        label (str): The label of the header entry
        id (str): Unique id identifying the current entry. This is necessary to allow for multiple identical labels
        level (int): The level defines at which position (bottom to top) within the table the current entry is located.
        children (list[str]): The labels of all entries that are located directly below the current entry
        children (list[str]): The ids of all entries that are located directly below the current entry
        children_items (list[str]): Items are the columns that are found in the table (i.e., not the spanners). For items,
         the names of the items can differ from what will be shown in the table (e.g., new_name:old_name). If they differ,
         the new_name will be saved in children_items; otherwise the old_name will be saved.
    """

    label: str
    id: str
    level: int
    children: list[str]
    children_ids: list[str]
    children_items: list[str]


def add_gt_spanners(gt_tbl: gt.GT, tbl: TableSpam) -> gt.GT:
    """Add the Great Table Spanners to a table.

    Args:
        gt_tbl (gt.GT): Great Table without spanners
        tbl (TableSpam): TableSpam table, wherein the spanners are defined that should be added to the Great Table

    Returns:
        gt.GT: Great Table with spanners
    """
    flattened_tbl = flatten_table(tbl)

    if flattened_tbl['flattened_lhs'] is not None:
        gt_tbl = add_gt_spanner_partial(
            gt_tbl=gt_tbl, tbl_partial=flattened_tbl['flattened_lhs']
        )

    if flattened_tbl['flattened_rhs'] is not None:
        gt_tbl = add_gt_spanner_partial(
            gt_tbl=gt_tbl, tbl_partial=flattened_tbl['flattened_rhs']
        )

    return gt_tbl


def add_gt_spanner_partial(gt_tbl: gt.GT, tbl_partial: list[FlattenedEntry]) -> gt.GT:
    """Add the Great Table spanners for the left hand side and right hand side of the table.

    Args:
        gt_tbl (gt.GT): Great Table without spanners
        tbl_partial (list[FlattenedEntry]): List with FlattenedEntry entries representing the table spanners

    Raises:
        ValueError: Error if there are no entries in tbl_partial

    Returns:
        gt.GT: Great Table with spanners
    """
    # The table spanners need to be added in the correct order. All children of
    # a spanner must already be in the table, otherwise we get an error.
    # The level tells us the order; we have to start with the lowest one
    if tbl_partial is None:
        raise ValueError('tbl_partial should not be None.')
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
                if parent_name != '_BASE_LEVEL_':
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
    """Translate the highly nested table headers into a flat list.

    The table headers in a TableSpam are represented as highly nested elements
    in the HeaderEntry of the lhs and rhs. This function takes these nested entries
    and translates them to a list.

    Args:
        tbl (TableSpam): Table

    Returns:
        dict[str, None | list[FlattenedEntry]]: Dict with flattened entries for lhs and rhs of the header.
    """
    if tbl.header['lhs'] is not None:
        flattened_lhs = flatten_table_partial(tbl_partial=tbl.header['lhs'])
    else:
        flattened_lhs = None

    flattened_rhs = flatten_table_partial(tbl_partial=tbl.header['rhs'])

    return {'flattened_lhs': flattened_lhs, 'flattened_rhs': flattened_rhs}


def flatten_table_partial(
    tbl_partial: HeaderEntry,
    id: str = '',
    flattened: None | list[FlattenedEntry] = None,
) -> None | list[FlattenedEntry]:
    """Translates the highly nested HeaderEntries into a flat list.

    Args:
        tbl_partial (HeaderEntry): The current header entry that should be flattened
        id (str, optional): Each entry is assigned a unique ID. This is necessary for GT, where
          duplicated names would results in issues. Defaults to "".
        flattened (None | list[FlattenedEntry], optional): The function is called recursively and fills this list. Defaults to None.

    Returns:
        None | list[FlattenedEntry]: A list with flattened entries. See FlattenedEntry.
    """
    if flattened is None:
        flattened = []

    if len(tbl_partial.entries) != 0:
        flattened_entry = FlattenedEntry(
            label=tbl_partial.name,
            id=f'{id}_{tbl_partial.name}',
            level=tbl_partial.level,
            children=[entry.name for entry in tbl_partial.entries],
            children_ids=[
                f'{id}_{tbl_partial.name}_{entry.name}' for entry in tbl_partial.entries
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
                tbl_partial=entry, id=f'{id}_{tbl_partial.name}', flattened=flattened
            )

    return flattened


class FormattingFunction(Protocol):
    """Format Great Tables.

    Provides a framework for the definition of functions that are used
    to format the great table. The function must accept a gt.GT as first
    arument. All other arguments should be set with args and kwargs. See
    default_formatting for an example.
    """

    def __call__(self, gt_tbl: gt.GT, *args: Any, **kwargs: Any) -> gt.GT:
        """Format Great Tables.

        Provides a framework for the definition of functions that are used
        to format the great table. The function must accept a gt.GT as first
        arument. All other arguments should be set with args and kwargs. See
        default_formatting for an example.

        Args:
            gt_tbl (gt.GT): Great Table that should be formatted
            *args: Additional arguments used by custom function
            **kwargs: Additional arguments used by custom function

        Returns:
            gt.GT: Formatted Great Table.
        """
        ...


def add_gt_rowname_separator(
    gt_tbl: gt.GT, right_of: str, separator_style: gt.style.borders
) -> gt.GT:
    """Adds a vertical line between rownames and data.

    Args:
        gt_tbl (gt.GT): Table without vertical line
        right_of (str): name of the column to the right of which the line should be drawn
        separator_style (gt.style.borders): style used for the vertical line

    Returns:
        gt.GT: Table with vertical line
    """
    gt_tbl = gt_tbl.tab_style(
        style=separator_style, locations=gt.loc.body(columns=right_of)
    )
    return gt_tbl


def add_gt_titles(gt_tbl: gt.GT, title: str | None, subtitle: str | None) -> gt.GT:
    """Add the GT title and subtitle to an existing table.

    Args:
        gt_tbl (gt.GT): Table without titles
        title (str | None): Text to put in title
        subtitle (str | None): Text to put in subtitle

    Returns:
        gt.GT: Table with titles
    """
    if title is not None:
        gt_tbl = gt_tbl.tab_header(
            title=title, subtitle=subtitle
        ).opt_align_table_header(align='left')
    return gt_tbl


def add_gt_footnote(gt_tbl: gt.GT, footnote: str) -> gt.GT:
    """Add the GT footnote to an existing table.

    Args:
        gt_tbl (gt.GT): Table without footnote
        footnote (str): Text to put in footnote

    Returns:
        gt.GT: Table with footnote
    """
    gt_tbl = gt_tbl.tab_source_note(footnote)
    return gt_tbl
