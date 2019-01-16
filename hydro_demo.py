from pyreporter.expression import (Formula, PiecewiseFormula,
                                   Calculator, TrailSolver,
                                   Variable, FractionVariable, Unit, Constant, Number, FlatDiv)
from pyreporter.re_reporter import Report, DefaultCover

V = Variable
FV = FractionVariable
U = Unit
C = Constant
N = Number

Q = V('Q', inform='渡槽的过水流量', unit=FlatDiv(U('m')**3,U('s')), precision=3)
A = V('A', inform='槽身过水断面闽面积', unit=U('m')**2, precision=3)
R = V('R', inform='水力半径', unit=U('m'), precision=3)
i = FV('i', inform='槽底比降')
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
    rep.set_cover(DefaultCover())
    rep.add_heading('计算公式', 1)
    rep.add(calc.get_definition())
    rep.add(calc.get_note())
    rep.add_heading('计算过程', 1)
    rep.add(ret1)
    rep.add(ret2)
    rep.add_paragraph('有计算可知，试算得到的水深', h.get_evaluation())
    rep.save('hydro_demo.docx')
