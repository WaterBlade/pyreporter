from .expression import Expression, Variable
from typing import List
from .reporter import (MathRun, MathComposite, MathMultiLine, MathMultiLineBrace,
                       SymbolNote)


class FormulaBase:
    def get_variable_set(self):
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

    def get_variable_set(self):
        return {self.variable}.union(self.expression.get_variable_set())

    def get_definition(self):
        return MathComposite(self.variable, MathRun('&=', sty='p'), self.expression)

    def get_procedure(self):
        result = self.expression.copy_result()
        if self.variable.unit is not None:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), self.expression,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result(),
                                 self.variable.unit)
        else:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), self.expression,
                                 MathRun('=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result())


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

    def get_variable_set(self):
        s = {self.variable}
        for exp, cond in zip(self.expression_list, self.condition_list):
            s = s.union(exp.get_variable_set())
            s = s.union(cond.get_variable_set())
        return s

    def get_definition(self):
        exps = list()
        for exp, cond in zip(self.expression_list, self.condition_list):
            exps.append(MathComposite(exp, MathRun('&,'), cond))
        return MathComposite(self.variable,
                             MathRun('&=', sty='p'),
                             MathMultiLineBrace(MathMultiLine(exps)))

    def get_procedure(self):
        result = self.expression.copy_result()
        # TODO: Just copy from the formula, need alter
        if self.variable.unit is not None:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result(),
                                 self.variable.unit)
        else:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result())


class FormulaSystem(FormulaBase):
    def __init__(self, *formula):
        self.formula_list = list(formula)  # type: List[Formula]

    def get_target_variable(self):
        return self.formula_list[0].variable

    def add(self, formula):
        self.formula_list.append(formula)

    def calc(self):
        for formula in self.formula_list[::-1]:
            formula.calc()

        return self.formula_list[0].variable.value

    def get_variable_set(self):
        s = set()
        for formula in self.formula_list:
            s = s.union(formula.get_variable_set())
        return s

    def get_definition(self):
        return MathMultiLine([formula.get_definition() for formula in self.formula_list])

    def get_procedure(self):
        reversed = self.formula_list[::-1]
        return MathMultiLine([formula.get_procedure() for formula in reversed])


class Calculator:
    def __init__(self):
        self.formula_system = FormulaSystem()

    def add(self, formula):
        self.formula_system.add(formula)

    def calc(self):
        return self.formula_system.calc()

    def get_definition(self):
        return self.formula_system.get_definition()

    def get_procedure(self):
        return self.formula_system.get_procedure()

    def get_symbol_note(self):
        return SymbolNote(self.formula_system.get_variable_set())


class TrailSolver(Calculator):
    def set_unknown(self, unknown_var):
        self.unknown_variable = unknown_var

    def solve(self, target_value, left=0.001, right=100, tol=1e-5, max_iter=100):
        target = self.formula_system.get_target_variable()
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
