from .reporter import (MathRun, MathComposite, MathMultiLine, MathMultiLineBrace,
                       SymbolNote)
from .expression import Expression, Variable
import math


def has_contained(cols, item):
    for x in cols:
        if x is item:
            return True
    return False


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
            for var in exp.get_variable_dict():
                if not has_contained(s, var):
                    s.append(var)
            for var in cond.get_variable_dict():
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
            for var in formula.get_variable_dict():
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
