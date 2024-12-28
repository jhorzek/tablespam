from typing import Self

class HeaderEntry():
    name: str
    item_name: str
    entries: list[Self]
    width: int
    level: int

    def __init__(self, name: str, item_name: str | None = None):
        if item_name is None:
            item_name = name
        self.name = name
        self.item_name = item_name
        self.entries = []

    def add_entry(self, entry: Self) -> None:
        self.entries.append(entry)

    def set_width(self, width: int) -> None:
        self.width = width
    
    def set_level(self, level: int) -> None:
        self.level = level

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, HeaderEntry):
            return False
        equal = all([self.name == other.name,
                   self.item_name == other.item_name])
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
    
