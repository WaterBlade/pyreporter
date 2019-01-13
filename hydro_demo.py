from pyreporter import (Report,
                        Formula, PiecewiseFormula,
                        Calculator, TrailSolver)
from pyreporter import expression as exp

V = exp.Variable
U = exp.Unit
C = exp.Constant
N = exp.Number

Q = V('Q', inform='渡槽的过水流量', unit=U('m')**3/U('s'), precision=3)
A = V('A', inform='槽身过水断面闽面积', unit=U('m')**2, precision=3)
R = V('R', inform='水力半径', unit=U('m'), precision=3)
i = V('i', inform='槽底比降', precision=5)
n = V('n', inform='槽身过水断面的槽壁糙率', precision=3)
chi = V('χ', inform='湿周', unit=U('m'), precision=3)

h = V('h', inform='水深', unit=U('m'), precision=3)
b = V('b', inform='槽身底宽', unit=U('m'), precision=3)

f1 = Formula(Q, 1/n*A*R**(N(2)/3)*i**(N(1)/2))
f2 = Formula(R, A/chi)
f3 = Formula(chi, b+2*h)
f4 = Formula(A, b*h)

calc = Calculator()

calc.add(f1)
calc.add(f2)
calc.add(f3)
calc.add(f4)

solver = TrailSolver()
solver.add(f1)
solver.add(f2)
solver.add(f3)
solver.add(f4)
solver.set_unknown(h)


if __name__ == '__main__':
    b.value = 2.4
    h.value = 3.0
    i.value = 1/2000
    n.value = 0.014
    calc.calc()
    ret1 = calc.get_procedure()

    print(solver.solve(12))
    ret2 = solver.get_procedure()
    rep = Report()
    rep.set_default_cover()
    rep.add_heading('计算公式', 1)
    rep.add_math_definition(calc.get_definition())
    rep.add_symbol_note(calc.get_symbol_note())
    rep.add_heading('计算过程', 1)
    rep.add_math_procedure(ret1)
    rep.add_math_procedure(ret2)
    rep.save('hydro_demo.docx')
