from tablespam.Entry import HeaderEntry
import pyparsing as pyp

class Formula:
    def __init__(self, formula: str):
        self.formula = formula
        self.expression = define_parser()
    
    def parse_formula(self) -> list[str]:
        parsed_formula = self.expression.parseString(self.formula).asList()
        return(parsed_formula)
    
    def get_entries(self) -> dict:
        parsed_formula = self.parse_formula()
        if parsed_formula[0] == "1":
            lhs = None
        else:
            lhs = create_entries(parsed_formula[0])
        rhs = create_entries(parsed_formula[1])

        return {"lhs": lhs, "rhs": rhs}
    
    def get_variables(self) -> dict[str, list[str]]:
        entries = self.get_entries()
        lhs = extract_variables(entries["lhs"])
        rhs = extract_variables(entries["rhs"])
        return {"lhs": lhs, "rhs": rhs}

def create_entries(entry_list: list[str], depth: int | None = None) -> list:
    if depth is None:
        depth = 1
        header_entry = HeaderEntry(name = "_BASE_LEVEL_", 
                                   item_name = "_BASE_LEVEL_")
    elif (depth > 1) & ((entry_list[1] != "=") | len(entry_list) < 3):
        raise ValueError(f"Expected a spanner name in {entry_list}.")
    else:
        spanner_name = entry_list[0]
        # We can now drop the first two entries as those are the spanner name
        # and the equal sign. Everything else should be actual entries.
        entry_list = entry_list[2:]

        header_entry = HeaderEntry(name = spanner_name, 
                                   item_name = spanner_name)
        
    header_entry.set_level(depth)

    for entry in entry_list:
        if isinstance(entry, str):
            # It's a variable
            variable = split_variable(entry)
            sub_entry = HeaderEntry(name = variable["name"],
                        item_name = variable["item_name"])
            sub_entry.set_level(depth + 1)
            header_entry.add_entry(sub_entry)
        elif isinstance(entry, list):
            header_entry.add_entry(create_entries(entry, depth = depth + 1))
        else:
            raise ValueError(f"Could not parse {entry_list}.")
    return header_entry

def extract_variables(entry_list: list[str],
                      variables: list[str] | None = None) -> list[str]:
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
    if var == "1":
        return({"name": var, "item_name": var})
    variable = (pyp.Word(pyp.alphas + "_", pyp.alphas + pyp.nums + "_") | 
                pyp.QuotedString("`", escChar='\\', unquote_results=True))
    variable_with_name = variable + pyp.Suppress(":") + variable
    if variable_with_name.matches(var):
        result = variable_with_name.parseString(var).asList()
        if(len(result) == 2):
            return({"name": result[0], "item_name": result[1]})
        else:
            raise ValueError(f"Expected two elements, got {result}")
    else:
        result = variable.parseString(var).asList()
        return({"name": result[0], "item_name": result[0]})
    

def define_variable() -> pyp.core.Combine:
    # Match regular variable names:
    #  Variable names must start with a letter or underscore, followed
    #  by any combination of letters, numbers, and underscores. 
    single_variable = pyp.Word(pyp.alphas + "_", pyp.alphas + pyp.nums + "_")

    # We also want to allow for labels with spaces and special characters in them. This is 
    # mostly required for renaming columns:
    #  Any variable in `` will be seen as one variable
    quoted_variable = pyp.QuotedString("`", escChar='\\', unquoteResults=False)
    base_variable = single_variable | quoted_variable

    # Match colon-separated variable patterns
    variable = pyp.Combine(base_variable + pyp.Optional(":" + base_variable))

    return variable

def define_operators() -> pyp.core.MatchFirst:
    # Operator definitions
    # We only need the following operators:
    #  = separates the name of a spanner from the spanner content
    #  + separates elements within a spanner
    equal = pyp.one_of("=")
    plus = pyp.Suppress("+")
    operator = equal | plus

    return operator

def define_parser():
    # Define a forward-declared grammar
    expr = pyp.Forward()

    # Additionally, we need braces that define groups which will form a spanner.
    # Note that the group itself may contain the expression again, so we have a
    # recursive algorithm
    term = define_variable() | pyp.Group(pyp.Suppress('(') + expr + pyp.Suppress(')'))
    expr <<= term + pyp.ZeroOrMore(define_operators() + term) # Recursive expression
    full_expression = ("1" | pyp.Group(expr).setResultsName("lhs"))+ pyp.Suppress("~") + pyp.Group(expr).setResultsName("rhs")

    return full_expression

form = Formula(formula="Some + (A = b1 + `this is a  variable`:name) ~ other + (d = (e = f:h)) + g + h")
parsed = form.parse_formula()

type(split_variable("klsdjf:`skjdf sdf`"))
entries = create_entries(parsed[1])
print(entries)
