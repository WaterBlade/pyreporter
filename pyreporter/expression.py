import math
from typing import List
from collections import OrderedDict
from .re_reporter import Math, MathDefinition, MathNote, MathObject


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


class Expression(MathObject):
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

    def visit(self, visitor):
        pass

    def get_variable_dict(self)->OrderedDict:
        d = self.left.get_variable_dict()
        if self.right is not None:
            d.update(self.right.get_variable_dict())
        return d

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

    def visit(self, visitor):
        return visitor.visit_add(self.left, self.right)


class Sub(Expression):
    def calc(self):
        return self.left.calc() - self.right.calc()

    def visit(self, visitor):
        return visitor.visit_sub(self.left, self.right)


class Mul(Expression):
    def calc(self):
        return self.left.calc() * self.right.calc()

    def visit(self, visitor):
        return visitor.visit_mul(self.left, self.right)


class Div(Expression):
    def calc(self):
        return self.left.calc() / self.right.calc()

    def visit(self, visitor):
        return visitor.visit_div(self.left, self.right)


class FlatDiv(Expression):
    def calc(self):
        return self.left.calc() / self.right.calc()

    def visit(self, visitor):
        return visitor.visit_flat_div(self.left, self.right)


class Pow(Expression):
    def calc(self):
        return math.pow(self.left.calc(), self.right.calc())

    def visit(self, visitor):
        if isinstance(self.left, Variable) and self.left.subscript is not None:
            left = self.left.copy()
            left.subscript = None
            return visitor.visit_pow_with_sub(left, Variable(self.left.subscript), self.right)
        else:
            return visitor.visit_pow(self.left, self.right)


class Radical(Expression):
    def calc(self):
        return math.pow(self.left.calc(), 1 / self.right.calc())

    def visit(self, visitor):
        return visitor.visit_radical(self.left, self.right)


class LesserThan(Expression):
    def calc(self):
        return self.left.calc() < self.right.calc()

    def visit(self, visitor):
        return visitor.visit_lesser_than(self.left, self.right)


class LesserOrEqual(Expression):
    def calc(self):
        return self.left.calc() <= self.right.calc()

    def visit(self, visitor):
        return visitor.visit_lesser_or_equal(self.left, self.right)


class Equal(Expression):
    def calc(self):
        return self.left.calc() == self.right.calc()

    def visit(self, visitor):
        return visitor.visit_equal(self.left, self.right)


class NotEqual(Expression):
    def calc(self):
        return self.left.calc() != self.right.calc()

    def visit(self, visitor):
        return visitor.visit_not_equal(self.left, self.right)


class GreaterThan(Expression):
    def calc(self):
        return self.left.calc() > self.right.calc()

    def visit(self, visitor):
        return visitor.visit_greater_than(self.left, self.right)


class GreaterOrEqual(Expression):
    def calc(self):
        return self.left.calc() >= self.right.calc()

    def visit(self, visitor):
        return visitor.visit_greater_or_equal(self.left, self.right)


class ToDegree(Expression):
    def calc(self):
        return math.degrees(self.left.calc())

    def visit(self, visitor):
        return self.left.visit(visitor)


class ToRadian(Expression):
    def calc(self):
        return math.radians(self.left.calc())

    def visit(self, visitor):
        return self.left.visit(visitor)


class Sin(Expression):
    def calc(self):
        return math.sin(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_sin(self.left)


class Cos(Expression):
    def calc(self):
        return math.cos(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_cos(self.left)


class Tan(Expression):
    def calc(self):
        return math.tan(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_tan(self.left)


class Cot(Expression):
    def calc(self):
        return 1 / math.tan(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_cot(self.left)


class ASin(Expression):
    def calc(self):
        return math.asin(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_arcsin(self.left)


class ACos(Expression):
    def calc(self):
        return math.acos(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_arccos(self.left)


class ATan(Expression):
    def calc(self):
        return math.atan(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_arctan(self.left)


class ACot(Expression):
    def calc(self):
        return math.pi - math.atan(self.left.calc())

    def visit(self, visitor):
        return visitor.visit_arccot(self.left)


# parenthesis: ()
class Pr(Expression):
    def calc(self):
        return self.left.calc()

    def visit(self, visitor):
        return visitor.visit_parenthesis(self.left)


# square bracket: []
class Sq(Expression):
    def calc(self):
        return self.left.calc()

    def visit(self, visitor):
        return visitor.visit_bracket(self.left)


# brace: {}
class Br(Expression):
    def calc(self):
        return self.left.calc()

    def visit(self, visitor):
        return visitor.visit_brace(self.left)


class Variable(Expression):
    def __init__(self, symbol, subscript=None, value=None, unit=None, precision=2, inform=None):
        super().__init__()
        self.symbol = symbol
        self.value = value
        self.subscript = subscript
        self.precision = precision
        self.unit = unit
        self.inform = inform

    def __hash__(self):
        return hash(self.symbol)

    def get_variable_dict(self):
        return OrderedDict({id(self): self})

    def copy(self):
        return Variable(self.symbol, self.subscript, self.precision, self.unit)

    def copy_result(self):
        return Number(self.value, self.precision)

    def get_evaluation(self):
        if self.unit is None:
            return MathLine(self.copy(), '=', self.copy_result())
        else:
            return MathLine(self.copy(), '=', self.copy_result(), self.unit)

    def calc(self):
        if self.value is None:
            raise ValueError(f'Variable {self.symbol} has no value!')
        else:
            return self.value

    def visit(self, visitor):
        if self.subscript is None:
            return visitor.visit_variable(self.symbol)
        else:
            return visitor.visit_subscript_variable(Variable(self.symbol), Variable(self.subscript))


class FractionVariable(Variable):
    def __init__(self, symbol, subscript=None, value=None, unit=None, inform=None):
        super().__init__(symbol=symbol, subscript=subscript, value=value, unit=unit, precision=0, inform=inform)

    def copy_result(self):
        den = 1 / self.value
        return 1 / Number(den, precision=0)


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

    def get_variable_dict(self):
        return OrderedDict()


class Unit(Variable):
    def __init__(self, symbol):
        super().__init__(symbol)

    def visit(self, visitor):
        return visitor.visit_unit(self.symbol)


class MathText(MathObject):
    def __init__(self, text):
        self.text = text

    def visit(self, visitor):
        return visitor.visit_math_text(self.text)


class MathLine(MathObject):
    def __init__(self, *items):
        self._list = list()
        for item in items:
            self.add(item)

    def add(self, item):
        if isinstance(item, str):
            self._list.append(MathText(item))
        elif isinstance(item, Expression):
            self._list.append(item)
        else:
            raise TypeError('Unknown math type % s' % type(item))

    def visit(self, visitor):
        return visitor.visit_math_line(self._list)


class MultiLine(MathObject):
    def __init__(self, *exps, included: str=None):
        self._list = list()
        for exp in exps:
            self.add(exp)
        self.included = included

    def add(self, item):
        if isinstance(item, MathLine):
            self._list.append(item)
        else:
            raise TypeError('Unknown math type % s' % type(item))

    def visit(self, visitor):
        return visitor.visit_multi_line(self._list, self.included)


class FormulaBase:
    def get_variable_dict(self):
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

    def get_variable_dict(self):
        d = OrderedDict({id(self.variable): self.variable})
        d.update(self.expression.get_variable_dict())
        return d

    def get_definition(self):
        return MathLine(self.variable, '&=', self.expression)

    def get_procedure(self):
        if self.variable.unit is not None:
            return MathLine(self.variable,
                            '&=', self.expression,
                            '=', self.expression.copy_result(),
                            '=', self.variable.copy_result(),
                            self.variable.unit)
        else:
            return MathLine(self.variable,
                            '&=', self.expression,
                            '=', self.expression.copy_result(),
                            '=', self.variable.copy_result())


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

    def get_variable_dict(self):
        d = OrderedDict({id(self): self})
        for exp, cond in zip(self.expression_list, self.condition_list):
            d.update(exp.get_variable_dict())
            d.update(cond.get_variable_dict())
        return d

    def get_definition(self):
        multi = MultiLine(included='left')

        for exp, cond in zip(self.expression_list, self.condition_list):
            multi.add(MathLine(exp, '&,', cond))

        return MathLine(self.variable, '&=', multi)

    def get_procedure(self):
        if self.variable.unit is not None:
            return MathLine(self.variable,
                            '&=', self.expression,
                            '=', self.expression.copy_result(),
                            '=', self.variable.copy_result(),
                            self.variable.unit)
        else:
            return MathLine(self.variable,
                            '&=', self.expression,
                            '=', self.expression.copy_result(),
                            '=', self.variable.copy_result())


class Calculator:
    def __init__(self):
        self.formula_list = list()  # type: List[Formula]

    def get_target_variable(self):
        return self.formula_list[0].variable

    def add(self, formula):
        self.formula_list.append(formula)

    def calc(self):
        for formula in self.formula_list[::-1]:
            formula.calc()

        return self.formula_list[0].variable.value

    def get_variable_dict(self):
        d = OrderedDict()
        for formula in self.formula_list:
            d.update(formula.get_variable_dict())
        return d

    def get_definition(self):
        return MathDefinition(MultiLine(*(formula.get_definition() for formula in self.formula_list)))

    def get_procedure(self):
        reversed = self.formula_list[::-1]
        return Math(MultiLine(*(formula.get_procedure() for formula in reversed)))

    def get_note(self):
        d = self.get_variable_dict()
        return MathNote(d.values())


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
