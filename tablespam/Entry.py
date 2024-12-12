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

    def add_entry(self, entry: Self):
        self.entries.append(entry)

    def set_width(self, width: int):
        self.width = width
    
    def set_level(self, level: int):
        self.level = level
    
