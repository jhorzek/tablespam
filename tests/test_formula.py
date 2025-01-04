from tablespam.Formulas import Formula
from tablespam.Entry import HeaderEntry
import pytest
import pyparsing


def test_valid_formulas():
    f = Formula("   x ~ y ")
    assert f.parse_formula() == [["x"], ["y"]]
    assert f.get_variables() == {"lhs": ["x"], "rhs": ["y"]}
    lhs = HeaderEntry(name="_BASE_LEVEL_", item_name="_BASE_LEVEL_")
    lhs.set_level(2)
    lhs.set_width(1)
    lhs.add_entry(HeaderEntry(name="x", item_name="x"))
    lhs.entries[0].set_level(1)
    lhs.entries[0].set_width(1)
    rhs = HeaderEntry(name="_BASE_LEVEL_", item_name="_BASE_LEVEL_")
    rhs.set_level(2)
    rhs.set_width(1)
    rhs.add_entry(HeaderEntry(name="y", item_name="y"))
    rhs.entries[0].set_level(1)
    rhs.entries[0].set_width(1)
    assert f.get_entries() == {"lhs": lhs, "rhs": rhs}

    f = Formula("x1 + x2 ~ y1 + y2 + y3")
    assert f.parse_formula() == [["x1", "x2"], ["y1", "y2", "y3"]]
    assert f.get_variables() == {"lhs": ["x1", "x2"], "rhs": ["y1", "y2", "y3"]}

    f = Formula("x1 + x2 ~ a:y1 + b:y2 + y3")
    assert f.parse_formula() == [["x1", "x2"], ["a:y1", "b:y2", "y3"]]
    assert f.get_variables() == {"lhs": ["x1", "x2"], "rhs": ["y1", "y2", "y3"]}

    f = Formula("x1 + x2 ~ a:`y1 lks+2|/` + b:y2 + y3")
    assert f.parse_formula() == [["x1", "x2"], ["a:`y1 lks+2|/`", "b:y2", "y3"]]
    assert f.get_variables() == {"lhs": ["x1", "x2"], "rhs": ["y1 lks+2|/", "y2", "y3"]}

    f = Formula("1 ~ a:`y1 lks+2|/` + b:y2 + y3")
    assert f.parse_formula() == ["1", ["a:`y1 lks+2|/`", "b:y2", "y3"]]
    assert f.get_variables() == {"lhs": None, "rhs": ["y1 lks+2|/", "y2", "y3"]}

    f = Formula("x1 + x2 ~ (sp1 = a:`y1 lks+2|/` + (sp2 = b:y2 + y3))")
    assert f.parse_formula() == [
        ["x1", "x2"],
        [["sp1", "=", "a:`y1 lks+2|/`", ["sp2", "=", "b:y2", "y3"]]],
    ]
    assert f.get_variables() == {"lhs": ["x1", "x2"], "rhs": ["y1 lks+2|/", "y2", "y3"]}
    lhs = HeaderEntry(name="_BASE_LEVEL_", item_name="_BASE_LEVEL_")
    lhs.set_level(2)
    lhs.set_width(2)
    lhs.add_entry(HeaderEntry(name="x1", item_name="x1"))
    lhs.entries[0].set_level(1)
    lhs.entries[0].set_width(1)
    lhs.add_entry(HeaderEntry(name="x2", item_name="x2"))
    lhs.entries[1].set_level(1)
    lhs.entries[1].set_width(1)
    rhs = HeaderEntry(name="_BASE_LEVEL_", item_name="_BASE_LEVEL_")
    rhs.set_level(4)
    rhs.set_width(3)
    rhs.add_entry(HeaderEntry(name="sp1", item_name="sp1"))
    rhs.entries[0].set_level(3)
    rhs.entries[0].set_width(3)
    rhs.entries[0].add_entry(HeaderEntry(name="a", item_name="y1 lks+2|/"))
    rhs.entries[0].entries[0].set_level(1)
    rhs.entries[0].entries[0].set_width(1)
    rhs.entries[0].add_entry(HeaderEntry(name="sp2", item_name="sp2"))
    rhs.entries[0].entries[1].set_level(2)
    rhs.entries[0].entries[1].set_width(2)
    rhs.entries[0].entries[1].add_entry(HeaderEntry(name="b", item_name="y2"))
    rhs.entries[0].entries[1].entries[0].set_level(1)
    rhs.entries[0].entries[1].entries[0].set_width(1)
    rhs.entries[0].entries[1].add_entry(HeaderEntry(name="y3", item_name="y3"))
    rhs.entries[0].entries[1].entries[1].set_level(1)
    rhs.entries[0].entries[1].entries[1].set_width(1)
    assert f.get_entries() == {"lhs": lhs, "rhs": rhs}


def test_invalid_formulas():
    with pytest.raises(pyparsing.exceptions.ParseException):
        f = Formula("x1 + x2 + a:y1 + b:y2 + y3")
        assert f.parse_formula()

    # incorrect symbol (*)
    with pytest.raises(pyparsing.exceptions.ParseException):
        f = Formula("x1 + x2 ~ a*y1 + b:y2 + y3")
        assert f.parse_formula()

    # spanner without name
    with pytest.raises(ValueError):
        f = Formula("(x1 + x2) ~ y1 + b:y2 + y3")
        # parsing should still work
        f.parse_formula()
        # but extracting the entries should throw an error.
        assert f.get_entries()

    # too many ~
    with pytest.raises(pyparsing.exceptions.ParseException):
        f = Formula("x1 + x2 ~~ y1 + b:y2 + y3")
        assert f.parse_formula()


test_invalid_formulas()
