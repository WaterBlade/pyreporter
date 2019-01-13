from .expression import Expression, Variable
from typing import List
from .reporter import MathRun, MathComposite, MathMultiLine, MathMultiLineBrace


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
                                 MathRun('&=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result(),
                                 self.variable.unit)
        else:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), result,
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
            s.union(exp.get_variable_set())
            s.union(cond.get_variable_set())
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

    def calc(self):
        length = len(self.formula_list)
        for i in range(length):
            self.formula_list[length - i - 1].calc()

        return self.formula_list[0].variable.value

    def get_variable_set(self):
        s = set()
        for formula in self.formula_list:
            s.union(formula.get_variable_set())
        return s

    def get_definition(self):
        return [formula.get_definition for formula in self.formula_list]

    def get_procedure(self):
        pass


class Calculator:
    def get_definition(self):
        pass

    def get_procedure(self):
        pass

    def get_symbol_note(self):
        pass


class BisectSolver:
    pass