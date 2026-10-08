import math
import os

import pytest

from sheets import errors, ops
from sheets.formula import (adjust_structure, normalize, rename_sheet,
                            shift_formula, toggle_absolute)
from sheets.numfmt import (edit_text, format_general, format_value,
                           parse_input, serial_to_datetime)
from sheets.refs import col_index, col_name, key, parse_addr, parse_range
from sheets.workbook import DEFAULT_STYLE, new_workbook


def make(cells, wb=None):
    wb = wb or new_workbook()
    sh = wb.sheets[0]
    for a, text in cells.items():
        r, c = parse_addr(a)
        sh.set_state(key(r, c), ops.input_state(sh, r, c, text))
    wb.recalc()
    return wb, sh


def v(sh, a):
    return sh.value(*parse_addr(a))


# ---------------------------------------------------------------- refs

def test_col_names():
    assert col_name(0) == "A"
    assert col_name(25) == "Z"
    assert col_name(26) == "AA"
    assert col_name(16383) == "XFD"
    for i in (0, 25, 26, 700, 16383):
        assert col_index(col_name(i)) == i


def test_parse_range():
    assert parse_range("A1:B3") == (0, 0, 2, 1)
    assert parse_range("B3:A1") == (0, 0, 2, 1)
    assert parse_range("C:C")[1] == 2
    assert parse_range("2:4")[0] == 1
    assert parse_range("ZZZZ1") is None


# ---------------------------------------------------------------- formulas

@pytest.mark.parametrize("formula,expected", [
    ("=1+2*3", 7),
    ("=(1+2)*3", 9),
    ("=-2^2", 4),
    ("=2^3^2", 64),
    ("=10%", 0.1),
    ('="a"&"b"&1', "ab1"),
    ("=1=1", True),
    ('="A"="a"', True),
    ("=3>2", True),
    ('=1<"a"', True),
    ("=SUM(1,2,3)", 6),
    ("=ROUND(2.675,2)", 2.68),
    ("=ROUND(-2.5,0)", -3),
    ("=ROUNDDOWN(3.99,0)", 3),
    ("=ROUNDUP(3.01,0)", 4),
    ("=MOD(-3,2)", 1),
    ("=INT(-1.5)", -2),
    ('=IF(1>2,"yes","no")', "no"),
    ("=IF(FALSE,1)", False),
    ('=IFERROR(1/0,"x")', "x"),
    ("=ISERROR(1/0)", True),
    ('=LEFT("hello",2)', "he"),
    ('=MID("hello",2,3)', "ell"),
    ('=LEN("hello")', 5),
    ('=UPPER("abc")', "ABC"),
    ('=PROPER("hello world")', "Hello World"),
    ('=TRIM("  a   b ")', "a b"),
    ('=SUBSTITUTE("a-b-c","-","+")', "a+b+c"),
    ('=SUBSTITUTE("a-b-c","-","+",2)', "a-b+c"),
    ('=FIND("l","hello")', 3),
    ('=SEARCH("L*o","hello")', 3),
    ('=TEXT(0.256,"0.0%")', "25.6%"),
    ('=TEXT(1234567.891,"#,##0.00")', "1,234,567.89"),
    ('=TEXT(DATE(2024,3,5),"mmm d, yyyy")', "Mar 5, 2024"),
    ('=TEXT(DATE(2024,3,5),"dddd")', "Tuesday"),
    ("=DATE(2024,13,1)=DATE(2025,1,1)", True),
    ("=YEAR(DATE(2024,2,29))", 2024),
    ("=DAY(EOMONTH(DATE(2023,2,10),0))", 28),
    ("=MONTH(EDATE(DATE(2024,1,31),1))", 2),
    ('=DATEDIF(DATE(2020,1,15),DATE(2024,3,1),"Y")', 4),
    ("=WEEKDAY(DATE(2024,3,5))", 3),
    ('=VALUE("1,234.5")', 1234.5),
    ('=CONCATENATE("a",1,TRUE)', "a1TRUE"),
    ('=TEXTJOIN(", ",TRUE,"a","","b")', "a, b"),
    ("=AND(TRUE,1)", True),
    ("=OR(FALSE,0)", False),
    ("=CHOOSE(2,\"a\",\"b\")", "b"),
    ('=SWITCH(2,1,"one",2,"two")', "two"),
    ("=PMT(0.05/12,360,200000)", -1073.6432460242797),
    ("=SUM({1,2;3,4})", 10),
    ("=MAX(1,5,3)", 5),
    ("=MEDIAN(1,2,3,4)", 2.5),
    ("=ABS(-3)", 3),
    ("=SQRT(16)", 4),
    ("=POWER(2,10)", 1024),
])
def test_formula_values(formula, expected):
    wb, sh = make({"A1": formula})
    got = v(sh, "A1")
    if isinstance(expected, float):
        assert got == pytest.approx(expected)
    else:
        assert got == expected


@pytest.mark.parametrize("formula,code", [
    ("=1/0", "#DIV/0!"),
    ('=1+"x"', "#VALUE!"),
    ("=NOSUCHFUNC(1)", "#NAME?"),
    ("=SQRT(-1)", "#NUM!"),
    ("=NA()", "#N/A"),
    ("=VLOOKUP(99,{1,2},1,FALSE)", "#N/A"),
])
def test_formula_errors(formula, code):
    wb, sh = make({"A1": formula})
    assert v(sh, "A1") == errors.XLError(code)


def test_ranges_and_lookups():
    cells = {"A1": "Name", "B1": "Qty", "C1": "Price"}
    data = [("apple", 3, 1.5), ("banana", 5, 0.25), ("cherry", 7, 3), ("date", 2, 4)]
    for i, (n, q, p) in enumerate(data, start=2):
        cells[f"A{i}"], cells[f"B{i}"], cells[f"C{i}"] = n, str(q), str(p)
    cells.update({
        "E1": '=VLOOKUP("cherry",A2:C5,3,FALSE)',
        "E2": "=SUMPRODUCT(B2:B5,C2:C5)",
        "E3": '=SUMIF(A2:A5,"b*",B2:B5)',
        "E4": '=COUNTIF(B2:B5,">=3")',
        "E5": '=INDEX(C2:C5,MATCH("date",A2:A5,0))',
        "E6": '=XLOOKUP("banana",A2:A5,B2:B5)',
        "E7": '=XLOOKUP("zzz",A2:A5,B2:B5,"none")',
        "E8": '=SUMIFS(B2:B5,C2:C5,">1",A2:A5,"<>date")',
        "E9": "=AVERAGE(B:B)",
        "E10": "=COUNTA(A:A)",
        "E11": "=COUNT(A1:C5)",
        "E12": '=AVERAGEIF(B2:B5,">2")',
        "E13": "=VLOOKUP(4,B2:C5,2)",
        "E14": "=MATCH(6,{1,3,5,7})",
        "E15": "=SUM(B2:B5*C2:C5)",
        "E16": "=ROWS(A2:C5)*COLUMNS(A2:C5)",
        "E17": "=LARGE(B2:B5,2)",
        "E18": "=SUBTOTAL(9,B2:B5)",
    })
    wb, sh = make(cells)
    assert v(sh, "E1") == 3
    assert v(sh, "E2") == pytest.approx(3 * 1.5 + 5 * 0.25 + 7 * 3 + 2 * 4)
    assert v(sh, "E3") == 5
    assert v(sh, "E4") == 3
    assert v(sh, "E5") == 4
    assert v(sh, "E6") == 5
    assert v(sh, "E7") == "none"
    assert v(sh, "E8") == 3 + 7
    assert v(sh, "E9") == pytest.approx(17 / 4)
    assert v(sh, "E10") == 5
    assert v(sh, "E11") == 8
    assert v(sh, "E12") == pytest.approx(5)
    assert v(sh, "E13") == 1.5  # approx match: largest <= 4 in unsorted col finds 3
    assert v(sh, "E14") == 3
    assert v(sh, "E16") == 12
    assert v(sh, "E17") == 5
    assert v(sh, "E18") == 17


def test_recalc_chain_and_dependents():
    wb, sh = make({"A1": "1", "A2": "=A1+1", "A3": "=A2*10", "B1": "=SUM(A1:A3)"})
    assert v(sh, "A3") == 20
    assert v(sh, "B1") == 23
    sh.set_state(key(0, 0), ops.input_state(sh, 0, 0, "5"))
    wb.recalc()
    assert v(sh, "A3") == 60
    assert v(sh, "B1") == 71


def test_long_chain_no_recursion_error():
    wb = new_workbook()
    sh = wb.sheets[0]
    sh.set_value(0, 0, 1.0)
    for r in range(1, 20000):
        sh.set_formula(r, 0, f"=A{r}+1")
    wb.recalc()
    assert sh.value(19999, 0) == 20000


def test_circular():
    wb, sh = make({"A1": "=B1", "B1": "=A1"})
    assert v(sh, "A1") == errors.CIRC or v(sh, "B1") == errors.CIRC


def test_cross_sheet_and_rename():
    wb, sh = make({"A1": "10"})
    s2 = wb.add_sheet("Data 2")
    s2.set_value(0, 0, 5.0)
    sh.set_formula(1, 0, "='Data 2'!A1*2")
    wb.recalc()
    assert v(sh, "A2") == 10
    wb.rename_sheet(s2, "Other")
    assert sh.formula_text(1, 0) == "=Other!A1*2"
    s2.set_value(0, 0, 7.0)
    wb.recalc()
    assert v(sh, "A2") == 14


def test_insert_delete_rows_adjust_formulas():
    wb, sh = make({"A1": "1", "A2": "2", "A3": "3", "B1": "=SUM(A1:A3)", "B2": "=A3"})
    sh.insert_rows(1, 2)  # before row 2
    assert sh.formula_text(0, 1) == "=SUM(A1:A5)"
    assert sh.formula_text(3, 1) == "=A5"
    assert sh.value(0, 1) == 6
    sh.delete_rows(4, 1)  # delete the row holding 3 (A5)
    assert sh.formula_text(3, 1) == "=#REF!"
    assert sh.formula_text(0, 1) == "=SUM(A1:A4)"
    assert sh.value(0, 1) == 3


def test_shift_formula():
    assert shift_formula("=A1+$B$2+C$3+$D4", 2, 1) == "=B3+$B$2+D$3+$D6"
    assert shift_formula("=A1", -1, 0) == "=#REF!"
    assert shift_formula("=SUM(A:A)", 5, 1) == "=SUM(B:B)"


def test_normalize_and_toggle():
    assert normalize("=sum(a1:b2)+if(true,1,2)") == "=SUM(A1:B2)+IF(TRUE,1,2)"
    t, cur = toggle_absolute("=A1+B2", 2)
    assert t == "=$A$1+B2"
    assert rename_sheet("=Sheet1!A1+'Sheet1'!B2", "Sheet1", "My Data") == "='My Data'!A1+'My Data'!B2"


def test_adjust_structure_ranges():
    assert adjust_structure("=SUM(A2:A10)", "S", "S", "row", 4, -2) == "=SUM(A2:A8)"
    assert adjust_structure("=SUM(A2:A10)", "S", "S", "row", 0, -3) == "=SUM(A1:A7)"
    assert adjust_structure("=Other!A5", "S", "Other", "row", 0, 1) == "=Other!A6"
    assert adjust_structure("=A5", "S", "Other", "row", 0, 1) == "=A5"


# ---------------------------------------------------------------- number formats

@pytest.mark.parametrize("value,fmt,expected", [
    (1234.5, "General", "1234.5"),
    (0.1 + 0.2, "General", "0.3"),
    (1 / 3, "General", "0.333333333"),
    (123456789012, "General", "1.23457E+11"),
    (1234.5, "#,##0.00", "1,234.50"),
    (-1234.5, "#,##0.00", "-1,234.50"),
    (-1234.5, "#,##0.00;(#,##0.00)", "(1,234.50)"),
    (0.256, "0%", "26%"),
    (0.5, "#.00", ".50"),
    (5, "000", "005"),
    (1234567, "#,##0,", "1,235"),
    (12345.678, "0.00E+00", "1.23E+04"),
    (1234.5, '"$"#,##0.00', "$1,234.50"),
    (1234.5, "$#,##0.00_);($#,##0.00)", "$1,234.50 "),
    (45356, "yyyy-mm-dd", "2024-03-05"),
    (45356, "m/d/yyyy", "3/5/2024"),
    (45356.75, "h:mm AM/PM", "6:00 PM"),
    (45356.75, "hh:mm:ss", "18:00:00"),
    (1.5, "[h]:mm", "36:00"),
    (5551234567, "(000) 000-0000", "(555) 123-4567"),
    (0, "0.00;-0.00;\"zero\"", "zero"),
    ("text", "@", "text"),
])
def test_format_value(value, fmt, expected):
    assert format_value(value, fmt)[0] == expected


@pytest.mark.parametrize("text,value,fmt", [
    ("42", 42.0, None),
    ("1.50", 1.5, "0.00"),
    ("00123", "00123", None),
    ("1,234", 1234.0, "#,##0"),
    ("$1,234.50", 1234.5, "$#,##0.00"),
    ("12%", 0.12, "0%"),
    ("(5)", -5.0, None),
    ("TRUE", True, None),
    ("2024-03-05", 45356.0, "yyyy-mm-dd"),
    ("3/5/2024", 45356.0, "m/d/yyyy"),
    ("hello", "hello", None),
    ("12345678901234567890", "12345678901234567890", None),
    ("1e3", 1000.0, None),
])
def test_parse_input(text, value, fmt):
    got, gfmt = parse_input(text)
    assert got == value
    assert gfmt == fmt


def test_dates_roundtrip():
    d = serial_to_datetime(45356)
    assert (d.year, d.month, d.day) == (2024, 3, 5)
    assert edit_text(45356.0, "yyyy-mm-dd") == "3/5/2024"
    assert edit_text(0.5, "h:mm") == "12:00"


# ---------------------------------------------------------------- ops

def test_fill_series():
    wb, sh = make({"A1": "1", "A2": "3", "B1": "Item 1", "C1": "Mon", "D1": "=A1*2", "E1": "3/5/2024"})
    states = ops.fill_states(sh, (0, 0, 1, 0), (0, 0, 4, 0))
    wb.apply_states(sh, states)
    assert [sh.value(r, 0) for r in range(5)] == [1, 3, 5, 7, 9]
    wb.apply_states(sh, ops.fill_states(sh, (0, 1, 0, 1), (0, 1, 2, 1)))
    assert sh.value(2, 1) == "Item 3"
    wb.apply_states(sh, ops.fill_states(sh, (0, 2, 0, 2), (0, 2, 2, 2)))
    assert sh.value(2, 2) == "Wed"
    wb.apply_states(sh, ops.fill_states(sh, (0, 3, 0, 3), (0, 3, 2, 3)))
    assert sh.formula_text(2, 3) == "=A3*2"
    assert sh.value(2, 3) == 10
    wb.apply_states(sh, ops.fill_states(sh, (0, 4, 0, 4), (0, 4, 1, 4)))
    assert sh.value(1, 4) == 45357


def test_sort():
    wb, sh = make({"A1": "Name", "B1": "Score", "A2": "bob", "B2": "5", "A3": "al", "B3": "9",
                   "A4": "cy", "B4": "", "A5": "di", "B5": "1", "C2": "=B2*2", "C3": "=B3*2",
                   "C4": "=B4*2", "C5": "=B5*2"})
    states = ops.sort_states(sh, (0, 0, 4, 2), [(1, False)], header=True)
    wb.apply_states(sh, states)
    assert [sh.value(r, 0) for r in range(1, 5)] == ["al", "bob", "di", "cy"]
    assert sh.formula_text(1, 2) == "=B2*2"
    assert sh.value(1, 2) == 18


def test_clipboard_text_parse():
    rows = ops.parse_clipboard_text('a\tb\r\n"multi\nline"\t"q""x"\r\n')
    assert rows == [["a", "b"], ["multi\nline", 'q"x']]


def test_paste_shifts_formulas():
    wb, sh = make({"A1": "1", "A2": "2", "B1": "=A1*10"})
    clip = ops.copy_rect(sh, (0, 1, 0, 1))
    wb.apply_states(sh, ops.paste_states(sh, clip, (1, 1, 1, 1)))
    assert sh.formula_text(1, 1) == "=A2*10"
    assert sh.value(1, 1) == 20


def test_current_region_and_autosum():
    wb, sh = make({"A1": "1", "A2": "2", "A3": "3", "B2": "x"})
    assert ops.current_region(sh, 1, 0) == (0, 0, 2, 1)
    assert ops.autosum_range(sh, 3, 0) == (0, 0, 2, 0)


def test_find_replace():
    wb, sh = make({"A1": "apple pie", "A2": "Apple", "A3": "=1+1"})
    assert ops.find_matches(sh, "apple") == [(0, 0), (1, 0)]
    assert ops.find_matches(sh, "apple", match_case=True) == [(0, 0)]
    assert ops.find_matches(sh, "apple", whole=True) == [(1, 0)]
    st = ops.replace_states(sh, [(0, 0), (1, 0)], "apple", "pear")
    wb.apply_states(sh, st)
    assert sh.value(0, 0) == "pear pie"
    assert sh.value(1, 0) == "pear"


def test_filter_hidden():
    wb, sh = make({"A1": "k", "A2": "x", "A3": "y", "A4": "x", "A5": ""})
    sh.autofilter = (0, 0, 4, 0)
    sh.filters = {0: {"values": {"x"}, "blanks": False}}
    assert ops.compute_filter_hidden(sh) == {2, 4}


@pytest.mark.parametrize("value,fmt,expected", [
    (45356.6, "# ?/?", "45356 3/5"),
    (0.5, "# ?/?", " 1/2"),
    (1.3, "# ?/16", "1 5/16"),
    (-1.75, "# ?/?", "-1 3/4"),
    (1234567.891, '0.0,,"M"', "1.2M"),
    (1234567, "#,##0,", "1,235"),
    (1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', " $1,234.50 "),
])
def test_more_formats(value, fmt, expected):
    assert format_value(value, fmt)[0] == expected
