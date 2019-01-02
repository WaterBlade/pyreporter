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
            # TODO: 有下标的指数运算需要特别考虑
            pass
        else:
            doc.write_pow(self.left, self.right)


class Radical(Expression):
    def calc(self):
        return math.pow(self.left.calc(), 1 / self.right.calc())

    def write_to(self, doc):
        doc.write_radical(self.left, self.right)


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
        doc.write_variable(self.symbol)

    def assign(self, value, unit=None):
        if unit is not None:
            factor = self.unit.convert_factor(unit)
            self.value = value * factor
        else:
            self.value = value


class Constant(Variable):
    def __init__(self, value, precision=2):
        p_format = f'%.{precision}f'
        super().__init__(p_format % value, value=value, precision=precision)


class UnitBase:
    def __init__(self, symbol: Expression=Variable(''), factor=1.0, base=(0, 0, 0)):
        self.symbol = symbol
        self.factor = factor
        self.base = base

    def set_symbol(self, symbol):
        self.symbol = symbol
        return self

    def scale(self, factor, symbol):
        return UnitBase(symbol, self.factor*factor, tuple(self.base))

    def __mul__(self, other):
        base = tuple(i+j for i, j in zip(self.base, other.base))
        factor = self.factor * other.factor
        symbol = self.symbol * other.symbol
        return UnitBase(symbol, factor, base)

    def __rmul__(self, other):
        return self.__mul__(other)

    def __truediv__(self, other):
        base = tuple(i-j for i, j in zip(self.base, other.base))
        factor = self.factor / other.factor
        symbol = self.symbol / other.symbol
        return UnitBase(symbol, factor, base)

    def __pow__(self, index):
        base = tuple(i*index for i in self.base)
        factor = math.pow(self.factor, index)
        symbol = self.symbol**2
        return UnitBase(symbol, factor, base)

    def convert_factor(self, unit):
        if all(i == j for i, j in zip(self.base, unit.base)):
            return self.factor / unit.factor
        else:
            raise RuntimeError('Unit Base Unmatched!')


V = Variable


class Unit:
    kg = UnitBase(V('kg'), 1.0, (1, 0, 0))
    m = UnitBase(V('m'), 1.0, (0, 1, 0))
    s = UnitBase(V('s'), 1.0, (0, 0, 1))

    cm = UnitBase(V('m'), 0.01, (0, 1, 0))
    mm = UnitBase(V('mm'), 0.001, (0, 1, 0))

    g = UnitBase(V('g'), 0.001, (1, 0, 0))

    N = (kg*m/s**2).set_symbol(V('N'))
    kN = N.scale(1000, V('kN'))
    MN = N.scale(1000000, V('MN'))

    Pa = (N/m**2).set_symbol(V('Pa'))
    kPa = (kN/m**2).set_symbol(V('kPa'))
    MPa = (MN/m**2).set_symbol(V('MPa'))


class Assignment:
    def __init__(self, var, expression):
        self.variable = var  # type: Variable
        self.expression = expression  # type: Expression

    def calc(self):
        self.variable.value = self.expression.calc()

    def write_to(self, doc):
        doc.write_assignment(self.variable, self.expression)
