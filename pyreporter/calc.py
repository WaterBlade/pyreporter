import math


def wrapper_number(number):
    if isinstance(number, float):
        i = 0
        while number < 10*i:
            i -= 1
        return Constant(number, precision=abs(i)+3)
    elif isinstance(number, int):
        return Constant(number, precision=0)
    else:
        return number


class Expression:
    def __init__(self, left=None, right=None):
        self.left = wrapper_number(left)
        self.right = wrapper_number(right)

    def copy(self):
        left = self.left.copy() if self.left is not None else None
        right = self.right.copy() if self.right is not None else None
        return self.__init__(left, right)

    def copy_result(self):
        return self.copy()

    def calc(self):
        pass

    def write_to(self, doc):
        pass

    def __add__(self, other):
        return Add(self, other)

    def __radd__(self, other):
        return Add(other, self)

    def __sub__(self, other):
        return Sub(self, other)

    def __rsub__(self, other):
        return Sub(other, self)

    def __mul__(self, other):
        return Mul(self, other)

    def __rmul__(self, other):
        return Mul(other, self)

    def __truediv__(self, other):
        return Div(self, other)

    def __rtruediv__(self, other):
        return Div(other, self)

    def __pow__(self, other):
        return Pow(self, other)

    def __rpow__(self, other):
        return Pow(other, self)


class Add(Expression):
    def calc(self):
        return self.left.calc() + self.right.calc()

    def write_to(self, doc):
        doc.write_add(self.left, self.right)


class Sub(Expression):
    def calc(self):
        return self.left.calc() - self.right.calc()

    def write_to(self, doc):
        doc.write_sub(self.left, self.right)


class Mul(Expression):
    def calc(self):
        return self.left.calc() * self.right.calc()

    def write_to(self, doc):
        doc.write_mul(self.left, self.right)


class Div(Expression):
    def calc(self):
        return self.left.calc() / self.right.calc()

    def write_to(self, doc):
        doc.write_div(self.left, self.right)


class Pow(Expression):
    def calc(self):
        return math.pow(self.left.calc(), self.right.calc())

    def write_to(self, doc):
        if isinstance(self.left, Variable) and self.left.subscript is not None:
            left = self.left.copy()
            left.subscript = None
            doc.write_pow_with_sub(left, Variable(self.left.subscript), self.right)
        else:
            doc.write_pow(self.left, self.right)


class Radical(Expression):
    def calc(self):
        return math.pow(self.left.calc(), 1 / self.right.calc())

    def write_to(self, doc):
        doc.write_radical(self.left, self.right)


class ToDegree(Expression):
    def calc(self):
        return math.degrees(self.left.calc())

    def write_to(self, doc):
        self.left.write_to(doc)


class ToRadian(Expression):
    def calc(self):
        return math.radians(self.left.calc())

    def write_to(self, doc):
        self.left.write_to(doc)


class Sin(Expression):
    def calc(self):
        return math.sin(self.left.calc())

    def write_to(self, doc):
        doc.write_sin(self.left)


class Cos(Expression):
    def calc(self):
        return math.cos(self.left.calc())

    def write_to(self, doc):
        doc.write_cos(self.left)


class Tan(Expression):
    def calc(self):
        return math.tan(self.left.calc())

    def write_to(self, doc):
        doc.write_tan(self.left)


class Cot(Expression):
    def calc(self):
        return 1 / math.tan(self.left.calc())

    def write_to(self, doc):
        doc.write_cot(self.left)


class ASin(Expression):
    def calc(self):
        return math.asin(self.left.calc())

    def write_to(self, doc):
        doc.write_arcsin(self.left)


class ACos(Expression):
    def calc(self):
        return math.acos(self.left.calc())

    def write_to(self, doc):
        doc.write_arccos(self.left)


class ATan(Expression):
    def calc(self):
        return math.atan(self.left.calc())

    def write_to(self, doc):
        doc.write_arctan(self.left)


class ACot(Expression):
    def calc(self):
        return math.pi - math.atan(self.left.calc())

    def write_to(self, doc):
        doc.write_arccot(self.left)


# parenthesis: ()
class Pr(Expression):
    def calc(self):
        return self.left.calc()

    def write_to(self, doc):
        doc.write_parenthesis(self.left)


# square bracket: []
class Sq(Expression):
    def calc(self):
        return self.left.calc()

    def write_to(self, doc):
        doc.write_bracket(self.left)


# brace: {}
class Br(Expression):
    def calc(self):
        return self.left.calc()

    def write_to(self, doc):
        doc.write_brace(self.left)


class Variable(Expression):
    def __init__(self, symbol, subscript=None, value=None, unit=None, precision=2):
        super().__init__()
        self.symbol = symbol
        self.value = value
        self.subscript = subscript
        self.precision = precision
        self.unit = unit

    def copy(self):
        return Variable(self.symbol, self.subscript, self.precision, self.unit)

    def copy_result(self):
        return Constant(self.value, self.precision)

    def calc(self):
        if self.value is None:
            raise ValueError(f'Variable {self.symbol} has no value!')
        else:
            return self.value

    def write_to(self, doc):
        if self.subscript is None:
            doc.write_variable(self.symbol)
        else:
            doc.write_subscript_variable(V(self.symbol), V(self.subscript))


class Constant(Variable):
    def __init__(self, value, precision=2):
        p_format = f'%.{precision}f'
        super().__init__(p_format % value, value=value, precision=precision)


class Unit(Variable):
    def __init__(self, symbol):
        super().__init__(symbol)

    def write_to(self, writer):
        writer.write_unit(self.symbol)


V = Variable
