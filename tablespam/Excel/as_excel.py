from __future__ import annotations  # noqa: D100
from typing import TYPE_CHECKING, Callable

import openpyxl as opy
from openpyxl.utils import get_column_interval
from openpyxl.cell.cell import Cell
import polars as pl
from tablespam.Excel.write_excel import write_excel_col
from tablespam.Excel.xlsx_styles import XlsxStyles, set_region_style
from tablespam.Excel.locations import get_locations


if TYPE_CHECKING:
    from tablespam.TableSpam import TableSpam
    from tablespam.Entry import HeaderEntry


def tbl_as_excel(
    tbl: TableSpam,
    workbook: opy.Workbook | None = None,
    sheet: str = 'Table',
    start_row: int = 1,
    start_col: int = 1,
    styles: XlsxStyles | None = None,
) -> opy.Workbook:
    if workbook is None:
        workbook = opy.Workbook()
    if styles is None:
        styles = XlsxStyles()

    if isinstance(sheet, str) and (sheet not in workbook.sheetnames):
        workbook.create_sheet(title=sheet)

    locations = get_locations(tbl=tbl, start_row=start_row, start_col=start_col)

    fill_background(
        tbl=tbl, workbook=workbook, sheet=sheet, locations=locations, styles=styles
    )

    write_title(
        tbl=tbl, workbook=workbook, sheet=sheet, locations=locations, styles=styles
    )

    write_header(
        workbook=workbook,
        sheet=sheet,
        header=tbl.header,
        locations=locations,
        styles=styles,
    )

    write_data(
        workbook=workbook,
        sheet=sheet,
        header=tbl.header,
        table_data=tbl.table_data,
        locations=locations,
        styles=styles,
    )

    # write_footnote(
    #    tbl=tbl, workbook=workbook, sheet=sheet, locations=locations, styles=styles
    # )

    # We create the outlines last as we may have to overwrite some border colors.
    # create_outlines(
    #    tbl=tbl, workbook=workbook, sheet=sheet, locations=locations, styles=styles
    # )

    return workbook


def fill_background(
    tbl: TableSpam,
    workbook: opy.Workbook,
    sheet: str,
    locations: dict[str, dict[str, int | None]],
    styles: XlsxStyles,
):
    sheet_ref = workbook[sheet]

    # Title
    if tbl.title is not None:
        set_region_style(
            sheet=sheet_ref,
            style=styles.bg_title,
            start_col=locations['col']['start_col_title'],
            end_col=locations['col']['end_col_title'],
            start_row=locations['row']['start_row_title'],
            end_row=locations['row']['end_row_title'],
        )

    # Subtitle
    if tbl.subtitle is not None:
        set_region_style(
            sheet=sheet_ref,
            style=styles.bg_subtitle,
            start_col=locations['col']['start_col_subtitle'],
            end_col=locations['col']['end_col_subtitle'],
            start_row=locations['row']['start_row_subtitle'],
            end_row=locations['row']['end_row_subtitle'],
        )

    # Header LHS
    if tbl.header['lhs'] is not None:
        set_region_style(
            sheet=sheet_ref,
            style=styles.bg_header_lhs,
            start_col=locations['col']['start_col_header_lhs'],
            end_col=locations['col']['end_col_header_lhs'],
            start_row=locations['row']['start_row_header'],
            end_row=locations['row']['end_row_header'],
        )

    # Header RHS
    set_region_style(
        sheet=sheet_ref,
        style=styles.bg_header_rhs,
        start_col=locations['col']['start_col_header_rhs'],
        end_col=locations['col']['end_col_header_rhs'],
        start_row=locations['row']['start_row_header'],
        end_row=locations['row']['end_row_header'],
    )

    # Rownames
    if tbl.header['lhs'] is not None:
        set_region_style(
            sheet=sheet_ref,
            style=styles.bg_rownames,
            start_col=locations['col']['start_col_header_lhs'],
            end_col=locations['col']['end_col_header_lhs'],
            start_row=locations['row']['start_row_data'],
            end_row=locations['row']['end_row_data'],
        )

    # Data
    set_region_style(
        sheet=sheet_ref,
        style=styles.bg_data,
        start_col=locations['col']['start_col_header_rhs'],
        end_col=locations['col']['end_col_header_rhs'],
        start_row=locations['row']['start_row_data'],
        end_row=locations['row']['end_row_data'],
    )

    # Footnote
    if tbl.footnote is not None:
        set_region_style(
            sheet=sheet_ref,
            style=styles.bg_footnote,
            start_col=locations['col']['start_col_footnote'],
            end_col=locations['col']['end_col_footnote'],
            start_row=locations['row']['start_row_footnote'],
            end_row=locations['row']['end_row_footnote'],
        )


def write_title(
    tbl: TableSpam,
    workbook: opy.Workbook,
    sheet: str,
    locations: dict[str, dict[str, int | None]],
    styles: XlsxStyles,
):
    if tbl.title is not None:
        loc = get_column_interval(
            start=locations['col']['start_col_title'],
            end=locations['col']['start_col_title'],
        )[0] + str(locations['row']['start_row_title'])
        workbook[sheet][loc] = tbl.title

        workbook[sheet].merge_cells(
            start_row=locations['row']['start_row_title'],
            start_column=locations['col']['start_col_title'],
            end_row=locations['row']['start_row_title'],
            end_column=locations['col']['start_col_title'],
        )
        set_region_style(
            sheet=workbook[sheet],
            style=styles.cell_title,
            start_row=locations['row']['start_row_title'],
            start_col=locations['col']['start_col_title'],
            end_row=locations['row']['start_row_title'],
            end_col=locations['col']['start_col_title'],
        )

    if tbl.subtitle is not None:
        loc = get_column_interval(
            start=locations['col']['start_col_subtitle'],
            end=locations['col']['start_col_subtitle'],
        )[0] + str(locations['row']['start_row_subtitle'])
        workbook[sheet][loc] = tbl.subtitle
        workbook[sheet].merge_cells(
            start_row=locations['row']['start_row_subtitle'],
            start_column=locations['col']['start_col_subtitle'],
            end_row=locations['row']['start_row_subtitle'],
            end_column=locations['col']['start_col_subtitle'],
        )
        set_region_style(
            sheet=workbook[sheet],
            style=styles.cell_subtitle,
            start_row=locations['row']['start_row_subtitle'],
            start_col=locations['col']['start_col_subtitle'],
            end_row=locations['row']['start_row_subtitle'],
            end_col=locations['col']['start_col_subtitle'],
        )


def merge_rownames(
    workbook: opy.Workbook,
    sheet: str,
    header: dict[str, HeaderEntry],
    locations: dict[str, dict[str, int | None]],
    styles: XlsxStyles,
):
    raise ValueError('Merging of row names not yet implemented.')


def write_header(
    workbook: opy.Workbook,
    sheet: str,
    header: dict[str, HeaderEntry],
    locations: dict[str, dict[str, int | None]],
    styles: XlsxStyles,
):
    if header['lhs'] is not None:
        max_level = max(header['lhs'].level, header['rhs'].level)

        write_header_entry(
            workbook=workbook,
            sheet=sheet,
            header_entry=header['lhs'],
            max_level=max_level,
            start_row=locations['row']['start_row_header'],
            start_col=locations['col']['start_col_header_lhs'],
            style=styles.cell_header_lhs,
        )
    else:
        max_level = header['rhs'].level

    write_header_entry(
        workbook=workbook,
        sheet=sheet,
        header_entry=header['rhs'],
        max_level=max_level,
        start_row=locations['row']['start_row_header'],
        start_col=locations['col']['start_col_header_rhs'],
        style=styles.cell_header_rhs,
    )


def write_header_entry(
    workbook: opy.Workbook,
    sheet: str,
    header_entry: HeaderEntry,
    max_level: int,
    start_row: int,
    start_col: int,
    style: Callable[[Cell], None],
):
    # write current entry name into table
    if header_entry.name != '_BASE_LEVEL_':
        loc = get_column_interval(start=start_col, end=start_col)[0] + str(
            start_row + (max_level - header_entry.level) - 1
        )
        workbook[sheet][loc] = header_entry.name

        if header_entry.width > 1:
            workbook[sheet].merge_cells(
                start_row=start_row + (max_level - header_entry.level) - 1,
                start_column=start_col,
                end_row=start_row + (max_level - header_entry.level) - 1,
                end_column=start_col + header_entry.width - 1,
            )

        set_region_style(
            sheet=workbook[sheet],
            style=style,
            start_row=start_row + (max_level - header_entry.level) - 1,
            start_col=start_col,
            end_row=start_row + (max_level - header_entry.level) - 1,
            end_col=start_col + header_entry.width - 1,
        )

    # entries may have sub-entries, that also have to be written down
    start_col_entry = start_col
    for entry in header_entry.entries:
        write_header_entry(
            workbook=workbook,
            sheet=sheet,
            header_entry=entry,
            max_level=max_level,
            start_row=start_row,
            start_col=start_col_entry,
            style=style,
        )
        start_col_entry += entry.width


def write_data(
    workbook: opy.Workbook,
    sheet: str,
    header: dict[str, HeaderEntry],
    table_data: dict[str, pl.DataFrame],
    locations: dict[str, dict[str, int | None]],
    styles: XlsxStyles,
):
    if header['lhs'] is not None:
        # Add row names and their styling
        i = 0
        for item in table_data['row_data'].columns:
            write_excel_col(
                workbook=workbook,
                sheet=sheet,
                data=table_data['row_data'].select(item),
                row_start=locations['row']['end_row_header'] + 1,
                col_start=locations['col']['start_col_header_lhs'] + i,
                base_style=styles.cell_rownames,
                data_styles=styles.data_styles,
            )
            i += 1

        if styles.merge_rownames:
            merge_rownames(
                workbook=workbook,
                sheet=sheet,
                table_data=table_data,
                locations=locations,
                styles=styles,
            )

    # Write the actual data itself
    i = 0
    for item in table_data['col_data'].columns:
        write_excel_col(
            workbook=workbook,
            sheet=sheet,
            data=table_data['col_data'].select(item),
            row_start=locations['row']['end_row_header'] + 1,
            col_start=locations['col']['start_col_header_rhs'] + i,
            base_style=styles.cell_data,
            data_styles=styles.data_styles,
        )
        i += 1

    # Apply custom styles
    if styles.cell_styles is not None:
        for sty in styles.cell_styles:
            if not set(sty.cols).issubset(set(table_data['col_data'].columns)):
                raise ValueError(
                    f"Trying to style an element that was not found in the data: {[c for c in sty.cols if c not in table_data["col_data"].columns]}."
                )
            if any(sty.rows > table_data['col_data'].shape[0]):
                raise ValueError(
                    'Trying to style a row outside of the range of the data.'
                )
        for col in sty.cols:
            for row in sty.rows:
                set_region_style(
                    sheet=workbook[sheet],
                    style=sty.style,
                    start_row=row,
                    start_col=col,
                    end_row=row,
                    end_col=col,
                )
