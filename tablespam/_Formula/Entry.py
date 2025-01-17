"""Entries defines the header entry structure for tables."""

from typing import Self  # noqa: D100


class HeaderEntry:
    """Single entry in the table header.

    A table header consists of one or more header entries.
    For example, the following table has 1 lhs header entry (Species) and 6
    rhs header entries (Sepal and Petal at level 2 and Length, Width, Length, Width at level 1).
    ```
    |         |      Sepal     |      Petal     |
    | Species | Length | Width | Length | Width |
    |:--------|-------:|------:|-------:|------:|
    | setosa  |    5.1 |   3.5 |    1.4 |   0.2 |
    ```
    Each of the entries (e.g., Sepal) is represented by a HeaderEntry.
    """

    name: str
    item_name: str
    entries: list[Self]
    width: int
    level: int

    def __init__(self, name: str, item_name: str | None = None):
        """Initialize a HeaderEntry.

        The HeaderEntry represents a single entry in the header of the table. Each HeaderEntry
        may have sub-entries. In that case, the current HeaderEntry is a spanner. If there are no
        sub-entries, the current HeaderEntry should be an item that is also found in the data set.

        Args:
            name (str): The name of the current entry
            item_name (str | None, optional): If the entry is an item (i.e., it is also found in the data set), the name that is shown in the table
              can be diffferent from the actual item_name. For example, the item_name may be mean_hp, but the name in the table is "Mean Horse Power". Defaults to None.
        """
        if item_name is None:
            item_name = name
        self.name = name
        self.item_name = item_name
        self.entries = []

    def add_entry(self, entry: Self) -> None:
        """Add an entry to a HeaderEntry.

        A HeaderEntry can be a single item (in which case there are no sub-entries)
        or a spanner (in which case there can be one or more sub-entries).
        `add_entry` adds sub-entries to the spanners.

        ```
        |         |      Sepal     |      Petal     | <- Each with 2 sub-entries
        | Species | Length | Width | Length | Width | <- No sub-entries
        |:--------|-------:|------:|-------:|------:|
        | setosa  |    5.1 |   3.5 |    1.4 |   0.2 |
        ```
        Args:
            entry (Self): A sub-entry to be added. This sub-entry must itself be a HeaderEntry.
        """
        self.entries.append(entry)

    def set_width(self, width: int) -> None:
        """Specify the width of the header entry.

        The width of an entry specifies the number of columns wihtin the
        data set that the entry spans.

        ```
        |         |          Sepal          | <- Sepal has width=3
        | Species | Length | Width | Color  | <- each width=1
        |:--------|-------:|------:|-------:|
        | setosa  |    5.1 |   3.5 | "red"  |
        ```

        Args:
            width (int): _description_
        """
        self.width = width

    def set_level(self, level: int) -> None:
        """Specify the level of the header entry.

        The level is the vertical location, with level=1 being the level
        that is at the bottom of the table.

        ```
        |         |      Sepal     |      Petal     | <- level 2
        | Species | Length | Width | Length | Width | <- level 1
        |:--------|-------:|------:|-------:|------:|
        | setosa  |    5.1 |   3.5 |    1.4 |   0.2 |
        ```

        Args:
            level (int): Level at which the header is located.
        """
        self.level = level

    def __eq__(self, other: object) -> bool:
        """Compare HeaderEntries.

        Args:
            other (object): Any object that should be compared to the current HeaderEntry.

        Returns:
            bool: True if both objects are equal, False otherwise.
        """
        if not isinstance(other, HeaderEntry):
            return False
        equal = all([self.name == other.name, self.item_name == other.item_name])
        if hasattr(self, 'width'):
            if hasattr(other, 'width'):
                equal = equal & (self.width == other.width)
            else:
                equal = False
        if hasattr(self, 'level'):
            if hasattr(other, 'level'):
                equal = equal & (self.level == other.level)
            else:
                equal = False

        for i in range(len(self.entries)):
            equal = equal & self.entries[i].__eq__(other.entries[i])
        return equal
