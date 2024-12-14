from tablespam.Formulas import Formula
import polars as pl
import great_tables as gt
from tablespam.as_gt import *
class TableSpam:
    def __init__(self,
                 data: pl.DataFrame,
                 formula: str,
                 title: str | None = None, 
                 subtitle: str | None = None,
                 footnote : str | None = None):
        self.data = data

        form = Formula(formula=formula)
        variables = form.get_variables()
        self.table_data = {
            "row_data": select_data(self.data, variables["lhs"]),
            "col_data": select_data(self.data, variables["rhs"])
        }

        self.title = title
        self.subtitle = subtitle
        self.footnote = footnote
        self.header = form.get_entries()
    
    def as_gt(self,
              groupname_col = None,
                    separator_style = gt.style.borders(sides = ["right"],
                                                      color = "gray"), 
                    auto_format = True, 
                    decimals: int = 2,
                    **kwargs) -> gt.GT:
        """
        Translates a table created with `tablespam` into a `gt` table. 

        The `tablespan` library does not provide built-in support for rendering tables as HTML. 
        However, with `as_gt`, a `tablespan` table can be converted into a `gt` table, 
        which supports HTML and LaTeX output. For more details on `gt`, see 
        <https://gt.rstudio.com/>.

        Args:
            groupname_col (str, optional): Column names to group data. Refer to the 
                `gt` documentation for details.
            separator_style (str, optional): Style of the vertical line separating row 
                names from data.
            auto_format (bool, optional): Whether the table should be formatted automatically. 
                Defaults to True.
            **kwargs: Additional keyword arguments passed to the `gt` function.

        Returns:
            GtTable: A `gt` table object that can be further customized using the `gt` package.

        Example:
        ```python
        cars = pl.DataFrame({
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
            "carb": [4, 4, 1, 1, 2, 1, 4, 2, 2, 4]
        })

        summarized_table = (
            cars
            .group_by(["cyl", "vs"])
            .agg([
                pl.len().alias("N"),
                pl.col("hp").mean().alias("mean_hp"),
                pl.col("hp").std().alias("sd_hp"),
                pl.col("wt").mean().alias("mean_wt"),
                pl.col("wt").std().alias("sd_wt"),
            ])
        )

        tbl = TableSpam(data = summarized_table,
                        formula = '''Cylinder:cyl + Engine:vs ~
                        N +
                        (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                        (`Weight` = Mean:mean_wt + SD:sd_wt)''',
                        title = "Motor Trend Car Road Tests",
                        subtitle = "A table created with tablespan",
                        footnote = "Data from the infamous mtcars data set.")

        tbl.as_gt().show()
        ```
        """
        if self.header["lhs"] is not None:
            data_set = pl.concat([self.table_data["row_data"],
                                self.table_data["col_data"]],
                                how = 'horizontal')
        else:
            data_set = self.table_data["col_data"]

        # Create the gt-like table (assuming `gt` functionality is implemented)
        gt_tbl = gt.GT(data=data_set, 
                    groupname_col=groupname_col, 
                    **kwargs)

        gt_tbl = add_gt_spanners(gt_tbl=gt_tbl, tbl=tbl)

        if tbl.header["lhs"] is not None:
            rowname_headers = tbl.table_data["row_data"].columns
            gt_tbl = add_gt_rowname_separator(
                gt_tbl=gt_tbl,
                right_of=rowname_headers[-1],  # Use the last row header
                separator_style=separator_style
            )

        # Add titles and subtitles if present
        if tbl.title is not None or tbl.subtitle is not None:
            gt_tbl = add_gt_titles(
                gt_tbl=gt_tbl,
                title=tbl.title,
                subtitle=tbl.subtitle
            )

        # Add footnotes if present
        if tbl.footnote is not None:
            gt_tbl = add_gt_footnote(
                gt_tbl=gt_tbl,
                footnote=tbl.footnote
            )

        # Apply auto-formatting if requested
        if auto_format:
            gt_tbl = add_automatic_formatting(gt_tbl,
                                    decimals = decimals)
            gt_tbl = gt_tbl.sub_missing(missing_text="")

        return gt_tbl

def select_data(data, variables):
    if variables is None:
        return None
    return data.select(variables)

cars = pl.DataFrame({
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
    "carb": [4, 4, 1, 1, 2, 1, 4, 2, 2, 4]
})

summarized_table = (
    cars
    .group_by(["cyl", "vs"])
    .agg([
        pl.len().alias("N"),
        pl.col("hp").mean().alias("mean_hp"),
        pl.col("hp").std().alias("sd_hp"),
        pl.col("wt").mean().alias("mean_wt"),
        pl.col("wt").std().alias("sd_wt"),
    ])
)

tbl = TableSpam(data = summarized_table,
                   formula = "1 ~ N + (`Horse Power` = Mean:mean_hp + SD:sd_hp) + (`Weight` = Mean:mean_wt + SD:sd_wt)")
tbl.header["lhs"]

