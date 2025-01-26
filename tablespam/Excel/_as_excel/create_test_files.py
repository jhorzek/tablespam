import openpyxl.workbook
from tablespam import TableSpam  # noqa: D100
from tablespam.Excel.xlsx_styles import CellStyle, XlsxStyles, DataStyle, style_color
import polars as pl
import openpyxl


def create_test_files_cars(
    target_dir: str | None = None,
) -> dict[str, openpyxl.workbook.Workbook]:
    """Create test excel files for internal tests.

    Args:
        target_dir (str|None): Target directory. When set to None (default) only returns a dict with the results
        Returns:
            dict[str, openpyxl.workbook.Workbook]: dict with results.
    """
    results = {}

    cars = pl.DataFrame(
        {
            'mpg': [21.0, 21.0, 22.8, 21.4, 18.7, 18.1, 14.3, 24.4, 22.8, 19.2],
            'cyl': [6, 6, 4, 6, 8, 6, 8, 4, 4, 6],
            'disp': [
                160.0,
                160.0,
                108.0,
                258.0,
                360.0,
                225.0,
                360.0,
                146.7,
                140.8,
                167.6,
            ],
            'hp': [110, 110, 93, 110, 175, 105, 245, 62, 95, 123],
            'drat': [3.90, 3.90, 3.85, 3.08, 3.15, 2.76, 3.21, 4.08, 3.92, 3.92],
            'wt': [
                2.620,
                2.875,
                2.320,
                3.215,
                3.440,
                3.460,
                3.570,
                3.190,
                3.150,
                3.440,
            ],
            'qsec': [
                16.46,
                17.02,
                18.61,
                19.44,
                17.02,
                20.22,
                15.84,
                20.00,
                22.90,
                18.30,
            ],
            'vs': [0, 0, 1, 1, 0, 1, 0, 1, 1, 1],
            'am': [1, 1, 1, 0, 0, 0, 0, 0, 0, 0],
            'gear': [4, 4, 4, 3, 3, 3, 3, 4, 4, 4],
            'carb': [4, 4, 1, 1, 2, 1, 4, 2, 2, 4],
        }
    )

    summarized_table = (
        cars.group_by(['cyl', 'vs'], maintain_order=True)
        .agg(
            [
                pl.len().alias('N'),
                pl.col('hp').mean().alias('mean_hp'),
                pl.col('hp').std().alias('sd_hp'),
                pl.col('wt').mean().alias('mean_wt'),
                pl.col('wt').std().alias('sd_wt'),
            ]
        )
        .sort(pl.col('cyl'), pl.col('vs'))
    )

    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                     N +
                     (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                     (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars'] = tbl.as_excel()

    if target_dir is not None:
        results['cars'].save(f'{target_dir}/cars.xlsx')

    results['cars_color_1'] = tbl.as_excel(styles=style_color('008080'))
    results['cars_color_2'] = tbl.as_excel(styles=style_color('FFFFC5'))

    if target_dir is not None:
        results['cars_color_1'].save(f'{target_dir}/cars_color_1.xlsx')
        results['cars_color_2'].save(f'{target_dir}/cars_color_2.xlsx')

    # Complex merging of rownames
    summarized_table_merge = summarized_table
    summarized_table_merge = summarized_table_merge.with_columns(pl.lit(1).alias('vs'))
    summarized_table_merge = summarized_table_merge.with_columns(
        pl.when(pl.arange(0, summarized_table_merge.height) == 0)
        .then(0)
        .otherwise(pl.col('vs'))
        .alias('vs')
    )
    summarized_table_merge = summarized_table_merge.with_columns(pl.lit(1).alias('N'))

    tbl_merge = TableSpam(
        data=summarized_table_merge,
        formula="""Cylinder:cyl + Engine:vs + N ~
                            (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                            (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_complex_merge'] = tbl_merge.as_excel()

    if target_dir is not None:
        results['cars_complex_merge'].save(f'{target_dir}/cars_complex_merge.xlsx')

    # offset
    results['cars_offset'] = tbl.as_excel(start_row=3, start_col=5)
    if target_dir is not None:
        results['cars_offset'].save(f'{target_dir}/cars_offset.xlsx')

    # custom cell styles
    results['cars_cell_styles'] = tbl.as_excel(
        styles=XlsxStyles(
            cell_styles=[
                CellStyle(
                    rows=[2, 3],
                    cols=['mean_hp'],
                    style=lambda c: setattr(c, 'font', openpyxl.styles.Font(bold=True)),
                ),
                CellStyle(
                    rows=[1, 4],
                    cols=['mean_wt', 'sd_wt'],
                    style=lambda c: setattr(c, 'font', openpyxl.styles.Font(bold=True)),
                ),
            ]
        )
    )

    if target_dir is not None:
        results['cars_cell_styles'].save(f'{target_dir}/cars_cell_styles.xlsx')

    def test_double(x: pl.DataFrame) -> bool:
        if len(x.columns) != 1:
            raise ValueError('Multiple columns passed to test.')
        return all([tp in [pl.Float32, pl.Float64] for tp in x.dtypes])

    # custom data type styles
    results['cars_data_styles'] = tbl.as_excel(
        styles=XlsxStyles(
            data_styles={
                'double': DataStyle(
                    test=test_double,
                    style=lambda c: setattr(c, 'font', openpyxl.styles.Font(bold=True)),
                ),
            }
        )
    )

    if target_dir is not None:
        results['cars_data_styles'].save(f'{target_dir}/cars_data_styles.xlsx')

    # Additional spanners
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                      (Results = N +
                          (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                          (`Weight` = Mean:mean_wt + SD:sd_wt))""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_additional_spanners'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_additional_spanners'].save(
            f'{target_dir}/cars_additional_spanners.xlsx'
        )

    # Spanner where we need additional lines
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                      (Results = N +
                          (`Inner result` = (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                            (`Weight` = Mean:mean_wt + SD:sd_wt)))""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_additional_spanners_left_right'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_additional_spanners_left_right'].save(
            f'{target_dir}/cars_additional_spanners_left_right.xlsx'
        )

    # no row names
    tbl = TableSpam(
        data=summarized_table,
        formula="""1 ~
                     (Results = N +
                        (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                        (`Weight` = Mean:mean_wt + SD:sd_wt))""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_no_row_names'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_no_row_names'].save(f'{target_dir}/cars_no_row_names.xlsx')

    # no titles
    tbl = TableSpam(
        data=summarized_table,
        formula="""1 ~
                      (Results = N +
                          (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                          (`Weight` = Mean:mean_wt + SD:sd_wt))""",
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_no_titles'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_no_titles'].save(f'{target_dir}/cars_no_titles.xlsx')

    # no titles, no footnote
    tbl = TableSpam(
        data=summarized_table,
        formula="""1 ~
                      (Results = N +
                          (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                          (`Weight` = Mean:mean_wt + SD:sd_wt))""",
    )

    results['cars_no_titles_no_footnote'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_no_titles_no_footnote'].save(
            f'{target_dir}/cars_no_titles_no_footnote.xlsx'
        )

    ## Test data with missing row names. See https://github.com/jhorzek/tablespan/issues/40

    summarized_table_NA = summarized_table.with_columns(
        pl.when(pl.arange(0, summarized_table.height) == 0)
        .then(None)
        .otherwise(pl.col('cyl'))
        .alias('cyl')
    ).with_columns(
        pl.when(pl.arange(0, summarized_table.height) == 1)
        .then(None)
        .otherwise(pl.col('vs'))
        .alias('vs')
    )

    tbl = TableSpam(
        data=summarized_table_NA,
        formula="""Cylinder:cyl + Engine:vs ~
                      N +
                      (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                      (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        title='Motor Trend Car Road Tests',
        subtitle='A table created with tablespan',
        footnote='Data from the infamous mtcars data set.',
    )

    results['cars_missing_rownames'] = tbl.as_excel()

    if target_dir is not None:
        results['cars_missing_rownames'].save(
            f'{target_dir}/cars_missing_rownames.xlsx'
        )

    return results
