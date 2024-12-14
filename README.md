# tablespam

> Create satisficing tables in Python the formula way.

`tablespam` is a re-implementation of the [`tablespan`](https://github.com/jhorzek/tablespan) in Python.

The objective of `tablespam` is to provide a “good enough” approach to
creating tables by leveraging formulas.

`tablespam` builds on the awesome packages XYZ and
[`great_tables`](https://posit-dev.github.io/great-tables/articles/intro.html), which allows tables created with
`tablespam` to be exported to the following formats:

2.  **HTML** (using [`gt`](https://gt.rstudio.com/))
3.  **LaTeX** (using [`gt`](https://gt.rstudio.com/))
4.  **RTF** (using [`gt`](https://gt.rstudio.com/))

## Example

```python
from tablespam import TableSpam
import polars as pl

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
                formula = """Cylinder:cyl + Engine:vs ~
                N +
                (`Horse Power` = Mean:mean_hp + SD:sd_hp) +
                (`Weight` = Mean:mean_wt + SD:sd_wt)""",
                title = "Motor Trend Car Road Tests",
                subtitle = "A table created with tablespan",
                footnote = "Data from the infamous mtcars data set.")

tbl.as_gt().show()
```