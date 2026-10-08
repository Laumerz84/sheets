"""Excel error values.  They are exceptions so evaluation code can raise them,
and also plain values stored in cells."""


class XLError(Exception):
    _instances = {}

    def __new__(cls, code):
        inst = cls._instances.get(code)
        if inst is None:
            inst = super().__new__(cls)
            inst.code = code
            cls._instances[code] = inst
        return inst

    def __init__(self, code):
        super().__init__(code)

    def __repr__(self):
        return f"XLError({self.code!r})"

    def __str__(self):
        return self.code

    def __eq__(self, other):
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self):
        return hash(("xlerr", self.code))

    def __reduce__(self):
        return (XLError, (self.code,))


NULL = XLError("#NULL!")
DIV0 = XLError("#DIV/0!")
VALUE = XLError("#VALUE!")
REF = XLError("#REF!")
NAME = XLError("#NAME?")
NUM = XLError("#NUM!")
NA = XLError("#N/A")
CIRC = XLError("#CIRC!")
SPILL = XLError("#SPILL!")
CALC = XLError("#CALC!")

ALL_CODES = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
             "#CIRC!", "#SPILL!", "#CALC!", "#GETTING_DATA"]


def from_code(code):
    code = code.upper()
    return XLError(code) if code in ALL_CODES else None
