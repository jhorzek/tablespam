from .TableSpam import TableSpam  # noqa: D104

"""Create satisficing tables with tablespam.

Tablespam is a very basic package with a sole objective: To simplify creating tables
that are "good enough" for many purposes and that can easily be exported to a variety of
formats. To this end, tablespam leverages the awesome packages great_tables 
(https://posit-dev.github.io/great-tables/articles/intro.html) and openpyxl 
(https://openpyxl.readthedocs.io/en/stable/).
"""
# Define the exports for the package
__all__ = ['TableSpam']
