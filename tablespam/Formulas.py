"""Formulas are a rudimentary and limited implementation of an R-style formula syntax for TableSpam."""

from tablespam.Entry import HeaderEntry
import pyparsing as pyp
from typing import Union

RecursiveList = list[Union['RecursiveList', str]]


class Formula:
    """Provides an R-formula like syntax to create tables."""

    def __init__(self, formula: str):
        """Initialize a Formula object.

        The Formula class provides an R-formula like syntax to create tables.

        The formula defines the table headers and is inspired by the R package `tables`.
        For example, `Species ~ Sepal_Length + Sepal_Width` defines a table with `Species` as the
        row names and `Sepal_Length` and `Sepal_Width` as columns.

        You can add spanner labels as follows:

        `Species ~ (Sepal = Sepal_Length + Sepal_Width) + (Petal = Petal_Length + Petal_Width)`

        You can also nest spanners, e.g.,
        `Species ~ (Sepal = (Length = Sepal_Length) + (Width = Sepal_Width))`.
        When exporting tables, you may want to rename some of your columns. For example,
        you may want to rename `Sepal_Length` and `Petal_Length` to `Length` and
        `Sepal_Width` and `Petal_Width` to `Width`. With `tablespam`, you can rename the
        items in the header using `new_name:old_name`.
        For example,
        `Species ~ (Sepal = Length:Sepal_Length + Width:Sepal_Width) + (Petal = Length:Petal_Length + Width:Petal_Width)`.
        Finally, to create a table without row names, use:
        `1 ~ (Sepal = Length:Sepal_Length + Width:Sepal_Width) + (Petal = Length:Petal_Length + Width:Petal_Width)`

        References:
        - tables: Murdoch D (2024). tables: Formula-Driven Table Generation. R package version 0.9.31, <https://dmurdoch.github.io/tables/>

        Args:
            formula (str): The formula used to describe the table. In general,
            the formula should have a left hand side and a right hand side, separated
            by a ~ (e.g., "a + b ~ c + d").

        >>> f = Formula('a + b ~ c + d')
        >>> f.parse_formula()
        [['a', 'b'], ['c', 'd']]
        """
        self.formula = formula
        self.expression = define_parser()

    def parse_formula(self) -> RecursiveList:
        """Parse_formula breaks down the formula into its elements to make it usable for creating the table.

        Returns:
            RecursiveList: nested lists with the elements of the table.
        """
        parsed_formula = self.expression.parseString(self.formula).asList()
        return parsed_formula

    def get_entries(self) -> dict:
        """Extracts the entries found in a table.

        Returns:
            dict: dict with entries for the lhs and rhs of the table.
        """
        parsed_formula = self.parse_formula()
        if parsed_formula[0] == '1':
            lhs = None
        else:
            lhs = create_entries(parsed_formula[0])
            lhs = add_header_width(lhs)
            lhs = add_header_level(lhs)
        rhs = create_entries(parsed_formula[1])
        rhs = add_header_width(rhs)
        rhs = add_header_level(rhs)

        return {'lhs': lhs, 'rhs': rhs}

    def get_variables(self) -> dict[str, list[str]]:
        """Extract the names of the variables found in the formula.

        The names should also be found in the data set.

        Returns:
            dict[str, list[str]]: The dictionary will have the names found on the
            left hand side (lhs) and right hand side (rhs) of the formula.
        """
        entries = self.get_entries()
        lhs = extract_variables(entries['lhs'])
        rhs = extract_variables(entries['rhs'])
        return {'lhs': lhs, 'rhs': rhs}


def create_entries(
    entry_list: RecursiveList | list[str] | str, depth: int | None = None
) -> HeaderEntry:
    """Create header entries.

    The entries are of type HeaderEntry and may contain other entries nested within them.

    Args:
        entry_list (RecursiveList): Current entry list. This list will be filled recursively
        depth (int | None, optional): current depht of the entry. Defaults to None.

    Raises:
        ValueError: Error in case of missing spanner name
        ValueError: Error in case of parsing issues

    Returns:
        list: the header entries
    """
    if depth is None:
        depth = 1
        header_entry = HeaderEntry(name='_BASE_LEVEL_', item_name='_BASE_LEVEL_')
    elif (depth > 1) & ((entry_list[1] != '=') | (len(entry_list) < 3)):
        raise ValueError(f'Expected a spanner name in {entry_list}.')
    else:
        # spanner names could still contain backticks; those are only removed
        # for variables. Therefore, we remove them here:
        if isinstance(entry_list[0], str):
            spanner_name = entry_list[0].strip('`')
        else:
            raise ValueError('Incorrect type; expected str.')
        # We can now drop the first two entries as those are the spanner name
        # and the equal sign. Everything else should be actual entries.
        entry_list = entry_list[2:]

        header_entry = HeaderEntry(name=spanner_name, item_name=spanner_name)

    for entry in entry_list:
        if isinstance(entry, str):
            # It's a variable
            variable = split_variable(entry)
            sub_entry = HeaderEntry(
                name=variable['name'], item_name=variable['item_name']
            )
            header_entry.add_entry(sub_entry)
        elif isinstance(entry, list):
            header_entry.add_entry(create_entries(entry, depth=depth + 1))
        else:
            raise ValueError(f'Could not parse {entry_list}.')
    return header_entry


def extract_variables(
    entry_list: HeaderEntry, variables: list[str] | None = None
) -> list[str]:
    """Get the names of the variables found in a formula.

    The variables are the items that we expect to also be in the data set.

    Args:
        entry_list (list[str]): list with entries for the header
        variables (list[str] | None, optional): list that gets filled with the names
        of the items. Defaults to None.

    Returns:
        list[str]: list with names of items
    """
    if entry_list is None:
        return None
    if variables is None:
        variables = []
    if len(entry_list.entries) == 0:
        variables.append(entry_list.item_name)
    else:
        for entry in entry_list.entries:
            extract_variables(entry, variables=variables)
    return variables


def split_variable(var: str) -> dict[str, str]:
    """Split item label and item name.

    Variables can be either a single statement (e.g. x2) or a name and
    an item (e.g., name:item). This function splits the variable into
    name and item.

    Args:
        var (str): variable

    Raises:
        ValueError: in case of more than one : within the string

    Returns:
        dict[str, str]: name and item_name
    """
    if var == '1':
        return {'name': var, 'item_name': var}
    variable = pyp.Word(
        pyp.alphas + '_', pyp.alphas + pyp.nums + '_'
    ) | pyp.QuotedString('`', escChar='\\', unquote_results=True)
    variable_with_name = variable + pyp.Suppress(':') + variable
    if variable_with_name.matches(var):
        result = variable_with_name.parseString(var).asList()
        if len(result) == 2:
            return {'name': result[0], 'item_name': result[1]}
        else:
            raise ValueError(f'Expected two elements, got {result}')
    else:
        result = variable.parseString(var).asList()
        return {'name': result[0], 'item_name': result[0]}


def define_variable() -> pyp.core.Combine:
    """Internal function defining the pattern that variables can have for pyparsing.

    Returns:
        pyp.core.Combine: pyparsing pattern for variables
    """
    # Match regular variable names:
    #  Variable names must start with a letter or underscore, followed
    #  by any combination of letters, numbers, and underscores.
    single_variable = pyp.Word(pyp.alphas + '_', pyp.alphas + pyp.nums + '_')

    # We also want to allow for labels with spaces and special characters in them. This is
    # mostly required for renaming columns:
    #  Any variable in `` will be seen as one variable
    quoted_variable = pyp.QuotedString('`', escChar='\\', unquoteResults=False)
    base_variable = single_variable | quoted_variable

    # Match colon-separated variable patterns
    variable = pyp.Combine(base_variable + pyp.Optional(':' + base_variable))

    return variable


def define_operators() -> pyp.core.ParserElement:
    """Internal function describing the pyparsing pattern for operators (+, :) used in the formulas.

    Returns:
        pyp.core.ParserElement: pyparsing definition of operators
    """
    # Operator definitions
    # We only need the following operators:
    #  = separates the name of a spanner from the spanner content
    #  + separates elements within a spanner
    equal = pyp.one_of('=')
    plus = pyp.Suppress('+')
    operator = equal | plus

    return operator


def define_parser() -> pyp.core.ParserElement:
    """Internal function defining the full syntax for the formula parser used to decipher the R-style formula.

    Returns:
        pyp.core.ParserElement: pyparsing definition
    """
    # Define a forward-declared grammar
    expr = pyp.Forward()

    # Additionally, we need braces that define groups which will form a spanner.
    # Note that the group itself may contain the expression again, so we have a
    # recursive algorithm
    term = define_variable() | pyp.Group(pyp.Suppress('(') + expr + pyp.Suppress(')'))
    expr <<= term + pyp.ZeroOrMore(define_operators() + term)  # Recursive expression
    full_expression = (
        ('1' | pyp.Group(expr).setResultsName('lhs'))
        + pyp.Suppress('~')
        + pyp.Group(expr).setResultsName('rhs')
        + pyp.StringEnd()
    )

    return full_expression


def add_header_width(parsed_partial: HeaderEntry) -> HeaderEntry:
    """Adds width information to each element in a table header.

    Tablespam represents headers as nested lists. This function calculates how
    wide each header entry must be by determining how many root elements each
    parent element spans. For example, if a header 'x' spans two elements 'x1' and 'x2':

        |    x    |
        | x1 | x2 |

    The function updates each header entry with its corresponding width. Here, the width of
    x is 2, and the width of x1 and x2 is 1.

    Args:
        parsed_partial (dict): The left-hand side or right-hand side of the parsed table.

    Returns:
        dict or None: The parsed_partial with additional width fields, or None if parsed_partial is None.
    """
    if parsed_partial is None:
        return None

    # In case of single level set width to 1 (current parsed_partial is not a spanner)
    if len(parsed_partial.entries) == 0:
        parsed_partial.set_width(1)
        return parsed_partial

    # Recursively go through all nested spanners and add to width
    parsed_partial.set_width(0)
    for entry_index in range(len(parsed_partial.entries)):
        parsed_partial.entries[entry_index] = add_header_width(
            parsed_partial.entries[entry_index]
        )
        parsed_partial.set_width(
            parsed_partial.width + parsed_partial.entries[entry_index].width
        )

    return parsed_partial


def add_header_level(parsed_partial: HeaderEntry) -> HeaderEntry:
    """Adds level information to each element in a table header.

    Tablespam represents headers as nested lists. This function calculates at which
    level each header entry must be by determining how many root elements each
    parent element spans. For example, if a header 'x' spans two elements 'x1' and 'x2':

        |  other  | <- level 3
        |    x    | <- level 2
        | x1 | x2 | <- level 1

    Args:
        parsed_partial (_type_): left or right hand side of the header

    Raises:
        ValueError: Error if no level could be defined

    Returns:
        _type_: the parsed_partial with added level info
    """
    if parsed_partial is None:
        return None

    # Level 1 is the level of the headers that are closest
    # to the data. When we have an empty entries list,
    # we reached that level.
    if len(parsed_partial.entries) == 0:
        parsed_partial.set_level(1)
        return parsed_partial

    parsed_partial.set_level(0)
    for entry_index in range(len(parsed_partial.entries)):
        # Bascially, what we are doing here is updating the level of one
        # subentry at a time
        parsed_partial.entries[entry_index] = add_header_level(
            parsed_partial.entries[entry_index]
        )
        # After updating the subentries, we check if the current level of the entry
        # is lower than the subentry + 1. If this is the case, we need to update the level
        # of the entry.
        parsed_partial.set_level(
            max(parsed_partial.level, parsed_partial.entries[entry_index].level + 1)
        )

    if parsed_partial.level == 0:
        raise ValueError(f'Could not set a level for {parsed_partial}')

    return parsed_partial
