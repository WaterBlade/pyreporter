import math
from typing import List
from collections import OrderedDict

__all__ = ['Variable', 'FractionVariable', 'Number', 'Unit',
           'Formula', 'PiecewiseFormula', 'Calculator', 'TrailSolver', 'Equation',
           'FlatDiv', 'Sin', 'ASin', 'Cos', 'ACos', 'Tan', 'ATan', 'Cot', 'ACot',
           'Radical', 'Pr', 'Sq', 'Br', 'Sum']


def wrapper_number(number):
    if isinstance(number, float) or isinstance(number, int):
        return Number(number)
    else:
        return number


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

    def visit(self, visitor):
        pass

    def get_variable_dict(self) -> OrderedDict:
        d = self.left.get_variable_dict()
        if self.right is not None:
            d.update(self.right.get_variable_dict())
        return d

    def __neg__(self):
        return Negative(self)

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


class Negative(Expression):
    def calc(self):
        return - self.left.calc()

    def visit(self, visitor):
        return visitor.visit_negative(self.left)


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

    def reverse(self):
        return GreaterOrEqual(self.left, self.right)


class LesserOrEqual(Expression):
    def calc(self):
        return self.left.calc() <= self.right.calc()

    def visit(self, visitor):
        return visitor.visit_lesser_or_equal(self.left, self.right)

    def reverse(self):
        return GreaterThan(self.left, self.right)


class Equal(Expression):
    def calc(self):
        return self.left.calc() == self.right.calc()

    def visit(self, visitor):
        return visitor.visit_equal(self.left, self.right)

    def reverse(self):
        return NotEqual(self.left, self.right)


class NotEqual(Expression):
    def calc(self):
        return self.left.calc() != self.right.calc()

    def visit(self, visitor):
        return visitor.visit_not_equal(self.left, self.right)

    def reverse(self):
        return Equal(self.left, self.right)


class GreaterThan(Expression):
    def calc(self):
        return self.left.calc() > self.right.calc()

    def visit(self, visitor):
        return visitor.visit_greater_than(self.left, self.right)

    def reverse(self):
        return LesserOrEqual(self.left, self.right)


class GreaterOrEqual(Expression):
    def calc(self):
        return self.left.calc() >= self.right.calc()

    def visit(self, visitor):
        return visitor.visit_greater_or_equal(self.left, self.right)

    def reverse(self):
        return LesserThan(self.left, self.right)


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
        return visitor.visit_square_bracket(self.left)


# brace: {}
class Br(Expression):
    def calc(self):
        return self.left.calc()

    def visit(self, visitor):
        return visitor.visit_brace(self.left)


class Variable(Expression):
    def __init__(self, symbol, subscript=None, value=None, unit=None, precision='auto', inform=None):
        super().__init__()
        self.symbol = symbol
        self.value = value
        self.subscript = subscript
        self.precision = precision
        if unit is not None and isinstance(unit, Div):
            unit = FlatDiv(unit.left, unit.right)
        self.unit = unit
        self.inform = inform

    def __repr__(self):
        if self.subscript is None:
            return f'{self.symbol}'
        else:
            return f'{self.symbol}-{self.subscript}'

    def __hash__(self):
        return id(self)

    def set(self, value):
        if isinstance(value, Expression):
            value = value.calc()
        self.value = value

    def get_variable_dict(self):
        return OrderedDict({id(self): self})

    def copy(self):
        return Variable(self.symbol, self.subscript, self.unit, self.precision, self.inform)

    def copy_result(self):
        return Number(self.value, self.precision)

    def calc(self):
        if self.value is None:
            raise ValueError(f'Variable {self} has no value!')
        else:
            return self.value

    def visit(self, visitor):
        return visitor.visit_variable(self.symbol, self.subscript)


class FractionVariable(Variable):
    def __init__(self, symbol, subscript=None, value=None, unit=None, inform=None):
        super().__init__(symbol=symbol, subscript=subscript, value=value, unit=unit, inform=inform)

    def copy_result(self):
        den = round(1 / self.value)
        return 1 / Number(den)


class Number(Variable):
    def __init__(self, value, precision=None):
        assert isinstance(value, float) or isinstance(value, int)
        super().__init__('number', value=value, precision=precision)

    def copy(self):
        return Number(self.value, self.precision)

    def visit(self, visitor):
        return visitor.visit_number(self.value, self.precision)

    def copy_result(self):
        return self.copy()

    def get_variable_dict(self):
        return OrderedDict()


class Unit(Variable):
    def __init__(self, symbol):
        super().__init__(symbol)

    def visit(self, visitor):
        return visitor.visit_unit(self.symbol)


class Sum(Expression):
    def __init__(self, exp, serial_variable_list: list):
        super().__init__(exp)
        for serial in serial_variable_list[1:]:
            assert len(serial) == len(serial_variable_list[0])
        self.serial_variable_list = serial_variable_list
        self.serial_length = len(serial_variable_list[0])
        self.expanded = None

    def calc(self):
        ret = 0
        for i in range(self.serial_length):
            for var in self.serial_variable_list:
                var.set_current(i)
            ret += self.left.calc()
        return ret

    def copy_result(self):
        ret = None
        for i in range(self.serial_length):
            for var in self.serial_variable_list:
                var.set_current(i)
            if ret is None:
                ret = self.left.copy_result()
            else:
                ret += self.left.copy_result()
        return ret

    def visit(self, visitor):
        return visitor.visit_sum(self.left)


class SerialVariable(Variable):
    def __init__(self, symbol, subscript=None, value=None, index='i', unit=None, precision='auto', inform=None):
        super().__init__(symbol=symbol, subscript=subscript, value=value,
                         unit=unit, precision=precision, inform=inform)
        self.index = index
        self._variable_list = list()
        self._curr = None

    def new(self, inform=None):
        index = len(self._variable_list) + 1
        v = VariableInSerial(serial=self, symbol=self.symbol, subscript=self.subscript,
                             index=index, unit=self.unit, precision=self.precision, inform=inform)
        self._variable_list.append(v)
        return v

    def __getitem__(self, item):
        return self._variable_list[item]

    def __len__(self):
        return len(self._variable_list)

    def calc(self):
        return self._curr.calc()

    def set_current(self, index):
        self._curr = self._variable_list[index]

    def copy_result(self):
        return self._curr.copy_result()

    def visit(self, visitor):
        return visitor.visit_serial_variable(self.symbol, self.subscript, self.index)


class VariableInSerial(Variable):
    def __init__(self, *, serial, symbol, subscript, index, unit, precision, inform):
        self.root = serial
        self.index = index
        super().__init__(symbol=symbol, subscript=subscript, unit=unit, precision=precision, inform=inform)

    def get_variable_dict(self):
        if self.inform is None:
            return self.root.get_variable_dict()
        else:
            return OrderedDict({id(self): self})

    def visit(self, visitor):
        return visitor.visit_serial_variable(self.symbol, self.subscript, self.index)


class FormulaBase:
    def get_variable_dict(self):
        pass

    def get_definition(self):
        pass

    def get_procedure(self):
        pass


class Formula(FormulaBase):
    def __init__(self, var, expression, long=False):
        self.variable = var  # type: Variable
        self.expression = expression  # type: Expression
        self.long = long

    def calc(self):
        self.variable.value = self.expression.calc()
        return self.variable.value

    def visit(self, visitor):
        return visitor.visit_formula(self.variable, self.expression, self.long)


class PiecewiseFormula(FormulaBase):
    def __init__(self, var, expression_list, condition_list, long: List[bool]=None):
        self.variable = var  # type: Variable
        assert len(expression_list) == len(condition_list)
        self.expression_list = expression_list  # type: List[Expression]
        self.condition_list = condition_list  # type: List[Expression]

        if long is None:
            long = [False] * len(self.expression_list)
        else:
            assert len(long) == len(self.expression_list)
        self.long_list = long
        self.long = False

        self.expression = None

    def calc(self):
        for exp, cond, long in zip(self.expression_list, self.condition_list, self.long_list):
            if cond.calc():
                self.variable.value = exp.calc()
                self.expression = exp
                self.long = long
                return self.variable.value
        return None

    def visit(self, visitor):
        return visitor.visit_piecewise_formula(self.variable, self.expression,
                                               self.expression_list, self.condition_list,
                                               self.long)


class ConditionFormula(Formula):
    def __init__(self, variable, expression, condition, long=False):
        super().__init__(variable, expression, long)
        self.condition = condition

    def calc(self):
        if self.condition.calc():
            return super().calc()

    def visit(self, visitor):
        return visitor.visit_condition_formula(self.variable, self.expression, self.condition.calc(), self.long)


class Calculator:
    def __init__(self, sequence=True):
        self.formula_list = list()  # type: List[Formula]
        self.sequence = sequence

    def add(self, formula):
        self.formula_list.append(formula)

    def calc(self):
        if self.sequence:
            list_ = self.formula_list[::-1]
        else:
            list_ = self.formula_list

        for formula in list_:
            formula.calc()

        return self.formula_list[0].variable.value

    def visit(self, visitor):
        visitor.visit_calculator(self.formula_list, sequence=self.sequence)


class TrailSolver(Calculator):
    def __init__(self):
        super().__init__()
        self.unknown_variable = None

    def get_target_variable(self):
        return self.formula_list[0].variable

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


class Equation:
    def __init__(self, exp, long=False):
        assert (isinstance(exp, Equal)
                or isinstance(exp, NotEqual)
                or isinstance(exp, GreaterThan)
                or isinstance(exp, GreaterOrEqual)
                or isinstance(exp, LesserThan)
                or isinstance(exp, LesserOrEqual))
        self.equation = exp
        self.left = Variable('left')
        self.right = Variable('right')
        self.result_equation = None  # type: Expression
        self.satisfied = True
        self.long = long

    def calc(self):
        self.satisfied = self.equation.calc()
        self.left.set(self.equation.left.calc())
        self.right.set(self.equation.right.calc())
        return self.satisfied

    def visit(self, visitor):
        # TODO: need fix
        if not self.satisfied:
            result_equation = self.equation.reverse()
        else:
            result_equation = self.equation

        if isinstance(result_equation, Equal):
            result_sign = 'equal'
        elif isinstance(result_equation, NotEqual):
            result_sign = 'not equal'
        return visitor.visit_equation(self.equation, self.satisfied,
                                      self.left, self.right,
                                      result_equation,
                                      result_sign)