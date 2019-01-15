from .reporter import (MathRun, MathComposite, MathMultiLine, MathMultiLineBrace,
                       SymbolNotes)
from . import reporter as rpt
import math


def wrapper_number(number):
    if isinstance(number, float):
        i = 0
        while number < 10*i:
            i -= 1
        return Number(number, precision=abs(i) + 3)
    elif isinstance(number, int):
        return Number(number, precision=0)
    else:
        return number


def has_contained(cols, item):
    for x in cols:
        if x is item:
            return True
    return False


class Expression:
    def __init__(self, left=None, right=None):
        self.left = wrapper_number(left)
        self.right = wrapper_number(right)

    def copy(self):
        left = self.left.copy() if self.left is not None else None
        right = self.right.copy() if self.right is not None else None
        return type(self)(left, right)

    def copy_result(self):
        left = self.left.copy_result() if self.left is not None else None
        right = self.right.copy_result() if self.right is not None else None
        return type(self)(left, right)

    def calc(self):
        pass

    def to_report(self):
        pass

    def get_variable_list(self):
        s = self.left.get_variable_list()

        if self.right is not None:
            for var in self.right.get_variable_list():
                if not has_contained(s, var):
                    s.append(var)

        return s

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

    def __lt__(self, other):
        return LesserThan(self, other)

    def __le__(self, other):
        return LesserOrEqual(self, other)

    def __eq__(self, other):
        return Equal(self, other)

    def __ne__(self, other):
        return NotEqual(self, other)

    def __gt__(self, other):
        return GreaterThan(self, other)

    def __ge__(self, other):
        return GreaterOrEqual(self, other)


class Add(Expression):
    def calc(self):
        return self.left.calc() + self.right.calc()

    def to_report(self):
        return rpt.Add(self.left.to_report(), self.right.to_report())


class Sub(Expression):
    def calc(self):
        return self.left.calc() - self.right.calc()

    def to_report(self):
        return rpt.Sub(self.left.to_report(), self.right.to_report())


class Mul(Expression):
    def calc(self):
        return self.left.calc() * self.right.calc()

    def to_report(self):
        return rpt.Mul(self.left.to_report(), self.right.to_report())


class Div(Expression):
    def calc(self):
        return self.left.calc() / self.right.calc()

    def to_report(self):
        return rpt.Div(self.left.to_report(), self.right.to_report())


class FlatDiv(Expression):
    def calc(self):
        return self.left.calc() / self.right.calc()

    def to_report(self):
        return rpt.FlatDiv(self.left.to_report(), self.right.to_report())


class Pow(Expression):
    def calc(self):
        return math.pow(self.left.calc(), self.right.calc())

    def to_report(self):
        return rpt.Pow(self.left.to_report(), self.right.to_report())


class Radical(Expression):
    def calc(self):
        return math.pow(self.left.calc(), 1 / self.right.calc())

    def to_report(self):
        return rpt.Radical(self.left.to_report(), self.right.to_report())


class LesserThan(Expression):
    def calc(self):
        return self.left.calc() < self.right.calc()

    def to_report(self):
        return rpt.LesserThan(self.left.to_report(), self.right.to_report())


class LesserOrEqual(Expression):
    def calc(self):
        return self.left.calc() <= self.right.calc()

    def to_report(self):
        return rpt.LesserOrEqual(self.left.to_report(), self.right.to_report())


class Equal(Expression):
    def calc(self):
        return self.left.calc() == self.right.calc()

    def to_report(self):
        return rpt.Equal(self.left.to_report(), self.right.to_report())


class NotEqual(Expression):
    def calc(self):
        return self.left.calc() != self.right.calc()

    def to_report(self):
        return rpt.NotEqual(self.left.to_report(), self.right.to_report())


class GreaterThan(Expression):
    def calc(self):
        return self.left.calc() > self.right.calc()

    def to_report(self):
        return rpt.GreaterThan(self.left.to_report(), self.right.to_report())


class GreaterOrEqual(Expression):
    def calc(self):
        return self.left.calc() >= self.right.calc()

    def to_report(self):
        return rpt.GreaterOrEqual(self.left.to_report(), self.right.to_report())


class ToDegree(Expression):
    def calc(self):
        return math.degrees(self.left.calc())

    def to_report(self):
        return self.left.to_report()


class ToRadian(Expression):
    def calc(self):
        return math.radians(self.left.calc())

    def to_report(self):
        return self.left.to_report()


class Sin(Expression):
    def calc(self):
        return math.sin(self.left.calc())

    def to_report(self):
        return rpt.Sin(self.left.to_report())


class Cos(Expression):
    def calc(self):
        return math.cos(self.left.calc())

    def to_report(self):
        return rpt.Cos(self.left.to_report())


class Tan(Expression):
    def calc(self):
        return math.tan(self.left.calc())

    def to_report(self):
        return rpt.Tan(self.left.to_report())


class Cot(Expression):
    def calc(self):
        return 1 / math.tan(self.left.calc())

    def to_report(self):
        return rpt.Cot(self.left.to_report())


class ASin(Expression):
    def calc(self):
        return math.asin(self.left.calc())

    def to_report(self):
        return rpt.ASin(self.left.to_report())


class ACos(Expression):
    def calc(self):
        return math.acos(self.left.calc())

    def to_report(self):
        return rpt.ACos(self.left.to_report())


class ATan(Expression):
    def calc(self):
        return math.atan(self.left.calc())

    def to_report(self):
        return rpt.ATan(self.left.to_report())


class ACot(Expression):
    def calc(self):
        return math.pi - math.atan(self.left.calc())

    def to_report(self):
        return rpt.ACot(self.left.to_report())


# parenthesis: ()
class Pr(Expression):
    def calc(self):
        return self.left.calc()

    def to_report(self):
        return rpt.Pr(self.left.to_report())


# square bracket: []
class Sq(Expression):
    def calc(self):
        return self.left.calc()

    def to_report(self):
        return rpt.Sq(self.left.to_report())


# brace: {}
class Br(Expression):
    def calc(self):
        return self.left.calc()

    def to_report(self):
        return rpt.Br(self.left.to_report())


class Variable(Expression):
    def __init__(self, symbol, subscript=None, value=None, unit=None, precision=2, inform=None):
        super().__init__()
        self.symbol = symbol
        self.value = value
        self.subscript = subscript
        self.precision = precision
        self.unit = unit
        self.inform = inform

    def get_variable_list(self):
        return [self]

    def copy(self):
        return Variable(self.symbol, self.subscript, self.precision, self.unit)

    def copy_result(self):
        return Number(self.value, self.precision)

    def calc(self):
        if self.value is None:
            raise ValueError(f'Variable {self.symbol} has no value!')
        else:
            return self.value

    def to_report(self):
        return rpt.Variable(self.symbol, self.subscript)


class FractionVariable(Variable):
    def __init__(self, symbol, subscript, num, den):
        super().__init__(symbol=symbol, subscript=subscript)
        self.num = num
        self.den = den

    def copy_result(self):
        return Number(self.num, precision=0) / Number(self.den, precision=0)


class Constant(Variable):
    def __init__(self, symbol, subscript=None, value=0, precision=2):
        super().__init__(symbol, subscript, value, precision)


class Number(Variable):
    def __init__(self, value, precision=0):
        if precision == 0:
            data = '%d' % value
        else:
            fmt = f'%.{precision}f'
            data = fmt % value
        super().__init__(data, value=value, precision=precision)

    def get_variable_list(self):
        return []


class Unit(Variable):
    def __init__(self, symbol):
        super().__init__(symbol)

    def to_report(self):
        return rpt.Unit(self.symbol)


class FormulaBase:
    def get_variable_list(self):
        pass

    def get_definition(self):
        pass

    def get_procedure(self):
        pass


class Formula(FormulaBase):
    def __init__(self, var, expression):
        self.variable = var  # type: Variable
        self.expression = expression  # type: Expression

    def calc(self):
        self.variable.value = self.expression.calc()
        return self.variable.value

    def get_variable_list(self):
        s = [self.variable]
        for var in self.expression.get_variable_list():
            if not has_contained(s, var):
                s.append(var)
        return s

    def get_definition(self):
        var = self.variable.to_report()
        exp = self.expression.to_report()
        return MathComposite(var, MathRun('&=', sty='p'), exp)

    def get_procedure(self):
        result = self.expression.copy_result().to_report()
        express = self.expression.to_report()
        var = self.variable.to_report()
        var_ret = self.variable.copy_result().to_report()
        if self.variable.unit is not None:
            unit = self.variable.unit.to_report()
            return MathComposite(var,
                                 MathRun('&=', sty='p'), express,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), var_ret,
                                 unit)
        else:
            return MathComposite(var,
                                 MathRun('&=', sty='p'), express,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), var_ret)


class PiecewiseFormula(FormulaBase):
    def __init__(self, var, expression_list, condition_list):
        self.variable = var  # type: Variable
        self.expression_list = expression_list  # type: List[Expression]
        self.condition_list = condition_list  # type: List[Expression]

        self.expression = None
        self.condition = None

    def calc(self):
        for exp, cond in zip(self.expression_list, self.condition_list):
            if cond.calc():
                self.variable.value = exp.calc()
                self.expression = exp
                self.condition = cond
                return self.variable.value
        return None

    def get_variable_list(self):
        s = [self.variable]
        for exp, cond in zip(self.expression_list, self.condition_list):
            for var in exp.get_variable_list():
                if not has_contained(s, var):
                    s.append(var)
            for var in cond.get_variable_list():
                if not has_contained(s, var):
                    s.append(var)
        return s

    def get_definition(self):
        exps = list()

        for exp, cond in zip(self.expression_list, self.condition_list):
            exp = exp.to_report()
            cond = cond.to_report()
            exps.append(MathComposite(exp, MathRun('&,'), cond))

        var = self.variable.to_report()

        return MathComposite(var,
                             MathRun('&=', sty='p'),
                             MathMultiLineBrace(MathMultiLine(exps)))

    def get_procedure(self):
        result = self.expression.copy_result().to_report()
        express = self.expression.to_report()
        var = self.variable.to_report()
        var_ret = self.variable.copy_result().to_report()
        if self.variable.unit is not None:
            unit = self.variable.unit.to_report()
            return MathComposite(var,
                                 MathRun('&=', sty='p'), express,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), var_ret,
                                 unit)
        else:
            return MathComposite(var,
                                 MathRun('&=', sty='p'), express,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), var_ret)


class Calculator:
    def __init__(self):
        self.formula_list = list()

    def get_target_variable(self):
        return self.formula_list[0].variable

    def add(self, formula):
        self.formula_list.append(formula)

    def calc(self):
        for formula in self.formula_list[::-1]:
            formula.calc()

        return self.formula_list[0].variable.value

    def get_variable_list(self):
        s = list()
        for formula in self.formula_list:
            for var in formula.get_variable_list():
                if not has_contained(s, var):
                    s.append(var)
        return s

    def get_definition(self):
        return MathMultiLine([formula.get_definition() for formula in self.formula_list])

    def get_procedure(self):
        reversed = self.formula_list[::-1]
        return MathMultiLine([formula.get_procedure() for formula in reversed])

    def get_symbol_note(self):
        note_list = list()
        for var in self.get_variable_list():
            if var.unit is not None:
                unit = var.unit.to_report()
            else:
                unit = None
            note_list.append([var.to_report(), var.inform, unit])

        return SymbolNotes(note_list)


class TrailSolver(Calculator):
    def set_unknown(self, unknown_var):
        self.unknown_variable = unknown_var

    def solve(self, target_value, left=0.001, right=100, tol=1e-5, max_iter=100):
        target = self.get_target_variable()
        unknown_var = self.unknown_variable

        def eq(x):
            unknown_var.value = x
            self.calc()
            return target_value - target.value

        mid = (left + right) / 2
        y_mid = eq(mid)
        y_left = eq(left)
        y_right = eq(right)
        while abs(y_mid) > tol:
            if max_iter <= 0:
                return None

            if y_left * y_mid > 0:
                left = mid
                y_left = eq(left)
            elif y_right * y_mid > 0:
                right = mid
                y_right = eq(right)
            else:
                return None
            mid = (left + right) / 2
            y_mid = eq(mid)

            max_iter -= 1

        return mid
