from tablespam import TableSpam
import polars as pl
import great_tables as gt
import re


def compare_tables(tbl_1: gt.GT, tbl_2: gt.GT) -> bool:
    tbl_1_html = tbl_1.as_raw_html(inline_css=True)
    # the following is taken from https://github.com/rstudio/gt/blob/1e4bae1af102c171a19316bca512db4260592645/tests/testthat/test-as_raw_html.R#L6
    # and removes the unique id:
    tbl_1_html = re.sub(r"id=\"[a-z]*?\"", "", tbl_1_html)

    tbl_2_html = tbl_2.as_raw_html(inline_css=True)
    # the following is taken from https://github.com/rstudio/gt/blob/1e4bae1af102c171a19316bca512db4260592645/tests/testthat/test-as_raw_html.R#L6
    # and removes the unique id:
    tbl_2_html = re.sub(r"id=\"[a-z]*?\"", "", tbl_2_html)

    return tbl_1_html == tbl_2_html


cars = pl.DataFrame(
    {
        "mpg": [21.0, 21.0, 22.8, 21.4, 18.7, 18.1, 14.3, 24.4, 22.8, 19.2],
        "cyl": [6, 6, 4, 6, 8, 6, 8, 4, 4, 6],
        "disp": [160.0, 160.0, 108.0, 258.0, 360.0, 225.0, 360.0, 146.7, 140.8, 167.6],
        "hp": [110, 110, 93, 110, 175, 105, 245, 62, 95, 123],
        "drat": [3.90, 3.90, 3.85, 3.08, 3.15, 2.76, 3.21, 4.08, 3.92, 3.92],
        "wt": [2.620, 2.875, 2.320, 3.215, 3.440, 3.460, 3.570, 3.190, 3.150, 3.440],
        "qsec": [16.46, 17.02, 18.61, 19.44, 17.02, 20.22, 15.84, 20.00, 22.90, 18.30],
        "vs": [0, 0, 1, 1, 0, 1, 0, 1, 1, 1],
        "am": [1, 1, 1, 0, 0, 0, 0, 0, 0, 0],
        "gear": [4, 4, 4, 3, 3, 3, 3, 4, 4, 4],
        "carb": [4, 4, 1, 1, 2, 1, 4, 2, 2, 4],
    }
)

summarized_table = cars.group_by(["cyl", "vs"]).agg(
    [
        pl.len().alias("N"),
        pl.col("hp").mean().alias("mean_hp"),
        pl.col("hp").std().alias("sd_hp"),
        pl.col("wt").mean().alias("mean_wt"),
        pl.col("wt").std().alias("sd_wt"),
    ]
)


def test_cars():
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                    N +
                    (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                    (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        title="Motor Trend Car Road Tests",
        subtitle="A table created with tablespan",
        footnote="Data from the infamous mtcars data set.",
    ).as_gt()
    expected = (
        gt.GT(summarized_table, groupname_col=None)
        .tab_header(
            title="Motor Trend Car Road Tests",
            subtitle="A table created with tablespan",
        )
        .opt_align_table_header(align="left")
        .tab_source_note(source_note="Data from the infamous mtcars data set.")
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Horse Power",
            columns=["mean_hp", "sd_hp"],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_style(
            style=gt.style.borders(sides="right", weight=gt.px(1), color="gray"),
            locations=gt.loc.body(columns=["vs"]),
        )
        .cols_label(
            cyl="Cylinder",
            vs="Engine",
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .fmt_number(columns=["mean_hp", "sd_hp", "mean_wt", "sd_wt"], decimals=2)
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)


def test_cars_no_autostyle():
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                    N +
                    (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                    (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        title="Motor Trend Car Road Tests",
        subtitle="A table created with tablespan",
        footnote="Data from the infamous mtcars data set.",
    ).as_gt(auto_format=False)
    expected = (
        gt.GT(summarized_table, groupname_col=None)
        .tab_header(
            title="Motor Trend Car Road Tests",
            subtitle="A table created with tablespan",
        )
        .opt_align_table_header(align="left")
        .tab_source_note(source_note="Data from the infamous mtcars data set.")
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Horse Power",
            columns=["mean_hp", "sd_hp"],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_style(
            style=gt.style.borders(sides="right", weight=gt.px(1), color="gray"),
            locations=gt.loc.body(columns=["vs"]),
        )
        .cols_label(
            cyl="Cylinder",
            vs="Engine",
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)


def test_cars_additional_spanners():
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                     (Results = N +
                        (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                        (`Weight` = Mean:mean_wt + SD:sd_wt))""",
        title="Motor Trend Car Road Tests",
        subtitle="A table created with tablespan",
        footnote="Data from the infamous mtcars data set.",
    ).as_gt()

    expected = (
        gt.GT(summarized_table, groupname_col=None)
        .tab_header(
            title="Motor Trend Car Road Tests",
            subtitle="A table created with tablespan",
        )
        .opt_align_table_header(align="left")
        .tab_source_note(source_note="Data from the infamous mtcars data set.")
        .tab_spanner(
            label="Mean",
            id="__BASE_LEVEL__Results_Horse Power_Mean",
            columns=["mean_hp"],
        )
        .tab_spanner(
            label="Standard Deviation",
            id="__BASE_LEVEL__Results_Horse Power_Standard Deviation",
            columns=["sd_hp"],
        )
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Results_Horse Power",
            spanners=[
                "__BASE_LEVEL__Results_Horse Power_Mean",
                "__BASE_LEVEL__Results_Horse Power_Standard Deviation",
            ],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Results_Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_spanner(
            label="Results",
            id="__BASE_LEVEL__Results",
            columns=["N"],
            spanners=[
                "__BASE_LEVEL__Results_Horse Power",
                "__BASE_LEVEL__Results_Weight",
            ],
        )
        .tab_style(
            style=gt.style.borders(sides="right", weight=gt.px(1), color="gray"),
            locations=gt.loc.body(columns=["vs"]),
        )
        .cols_label(
            cyl="Cylinder",
            vs="Engine",
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .fmt_number(columns=["mean_hp", "sd_hp", "mean_wt", "sd_wt"], decimals=2)
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)


def test_cars_no_rownames():
    tbl = TableSpam(
        data=summarized_table,
        formula="""1 ~
                     (Results = N +
                        (`Horse Power` = (Mean = Mean:mean_hp) + (`Standard Deviation` = SD:sd_hp)) +
                        (`Weight` = Mean:mean_wt + SD:sd_wt))""",
        title="Motor Trend Car Road Tests",
        subtitle="A table created with tablespan",
        footnote="Data from the infamous mtcars data set.",
    ).as_gt()

    summarized_table_no_rownames = summarized_table.drop(["cyl", "vs"])

    expected = (
        gt.GT(summarized_table_no_rownames, groupname_col=None)
        .tab_header(
            title="Motor Trend Car Road Tests",
            subtitle="A table created with tablespan",
        )
        .opt_align_table_header(align="left")
        .tab_source_note(source_note="Data from the infamous mtcars data set.")
        .tab_spanner(
            label="Mean",
            id="__BASE_LEVEL__Results_Horse Power_Mean",
            columns=["mean_hp"],
        )
        .tab_spanner(
            label="Standard Deviation",
            id="__BASE_LEVEL__Results_Horse Power_Standard Deviation",
            columns=["sd_hp"],
        )
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Results_Horse Power",
            spanners=[
                "__BASE_LEVEL__Results_Horse Power_Mean",
                "__BASE_LEVEL__Results_Horse Power_Standard Deviation",
            ],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Results_Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_spanner(
            label="Results",
            id="__BASE_LEVEL__Results",
            columns=["N"],
            spanners=[
                "__BASE_LEVEL__Results_Horse Power",
                "__BASE_LEVEL__Results_Weight",
            ],
        )
        .cols_label(
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .fmt_number(columns=["mean_hp", "sd_hp", "mean_wt", "sd_wt"], decimals=2)
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)


def test_cars_no_title():
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                    N +
                    (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                    (`Weight` = Mean:mean_wt + SD:sd_wt)""",
        footnote="Data from the infamous mtcars data set.",
    ).as_gt()
    expected = (
        gt.GT(summarized_table, groupname_col=None)
        .opt_align_table_header(align="left")
        .tab_source_note(source_note="Data from the infamous mtcars data set.")
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Horse Power",
            columns=["mean_hp", "sd_hp"],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_style(
            style=gt.style.borders(sides="right", weight=gt.px(1), color="gray"),
            locations=gt.loc.body(columns=["vs"]),
        )
        .cols_label(
            cyl="Cylinder",
            vs="Engine",
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .fmt_number(columns=["mean_hp", "sd_hp", "mean_wt", "sd_wt"], decimals=2)
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)


def test_cars_no_footnote():
    tbl = TableSpam(
        data=summarized_table,
        formula="""Cylinder:cyl + Engine:vs ~
                    N +
                    (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                    (`Weight` = Mean:mean_wt + SD:sd_wt)""",
    ).as_gt()
    expected = (
        gt.GT(summarized_table, groupname_col=None)
        .opt_align_table_header(align="left")
        .tab_spanner(
            label="Horse Power",
            id="__BASE_LEVEL__Horse Power",
            columns=["mean_hp", "sd_hp"],
        )
        .tab_spanner(
            label="Weight",
            id="__BASE_LEVEL__Weight",
            columns=["mean_wt", "sd_wt"],
        )
        .tab_style(
            style=gt.style.borders(sides="right", weight=gt.px(1), color="gray"),
            locations=gt.loc.body(columns=["vs"]),
        )
        .cols_label(
            cyl="Cylinder",
            vs="Engine",
            mean_hp="Mean",
            sd_hp="SD",
            mean_wt="Mean",
            sd_wt="SD",
        )
        .fmt_number(columns=["mean_hp", "sd_hp", "mean_wt", "sd_wt"], decimals=2)
        .sub_missing(missing_text="")
    )
    assert compare_tables(tbl, expected)
