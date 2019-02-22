from pyreporter.calculator import (Formula, PiecewiseFormula,
                                   Calculator, TrailSolver, Pr,
                                   Variable, SerialVariable, FractionVariable, Unit, Number, FlatDiv, Sum)
from pyreporter import Report, DefaultCover, Math, Procedure, Definition, Note, VariableValue
import math

V = Variable
SV = SerialVariable
FV = FractionVariable
U = Unit
N = Number
Value = VariableValue


Ra = V('R', 'a', inform='单桩竖向承载力特征值', unit=U('kN'))
K = V('K', value=2, inform='安全系数')
Quk = V('Q', 'uk', inform='单桩极限承载力标准值', unit=U('kN'))
Qsk = V('Q', 'sk', inform='总极限吃的阻力标准值', unit=U('kN'))
Qpk = V('Q', 'pk', inform='总极限端阻力标准值', unit=U('kN'))
qsik = SV('q', 'sk', inform='桩侧土极限测阻力标准值', unit=U('kPa'))
li = SV('l', inform='桩侧土对应的桩长', unit=U('m'))
ψsi = V('ψ', 'si', inform='大直径桩侧阻力')
ψp = V('ψ', 'p', inform='大直径桩端阻力')
qpk = V('q', 'pk', inform='极限端阻力标准值', unit=U('kPa'))
u = V('u', inform='桩身周长', unit=U('m'))
Ap = V('A', 'p', inform='桩端横截面面积', unit=U('m')**2)
d = V('d', inform='桩身直径', unit=U('m'))
Pi = V('π', value=math.pi, inform='圆周率')

R_calc = Calculator()
R_calc.add(Formula(Ra, 1/K*Quk))
R_calc.add(Formula(Quk, Qsk + Qpk))
R_calc.add(Formula(Qsk, u*Sum(ψsi*qsik*li, [qsik, li])))
R_calc.add(Formula(Qpk, ψp * qpk * Ap))

P_calc = Calculator(sequence=False)
P_calc.add(Formula(ψsi, Pr(0.8/d)**(N(1)/3)))
P_calc.add(Formula(ψp, Pr(0.8/d)**(N(1)/3)))

G_calc = Calculator()
G_calc.add(Formula(u, Pi * d))
G_calc.add(Formula(Ap, Pi * d ** 2 / 4))


if __name__ == '__main__':
    d.set(1.2)
    qpk.set(1500)

    G_calc.calc()
    P_calc.calc()
    R_calc.calc()

    rep = Report()
    rep.set_cover(DefaultCover('康苏水库工程', '溢洪洞出口桩基计算', '水  工', '技  施'))

    rep.add_heading('计算基本参数及结果', 1)
    rep.add_heading('基本参数', 2)
    rep.add_paragraph('桩身直径：', Value(d))
    rep.add_paragraph('极限端阻力：', Value(qpk))

    rep.add_heading('计算结果', 2)
    rep.add_paragraph('单桩竖向承载力特征值：', Value(Ra))
    rep.add_paragraph('单桩竖向极限承载力标准值：', Value(Quk))

    rep.add_heading('计算过程', 1)
    rep.add_heading('计算公式', 2)
    rep.add_paragraph('根据《建筑桩基技术规范》5.2.2与5.3.6，大直径桩的单桩竖向承载力特征值及极限承载力标准值的计算公式如下：')
    rep.add(Definition(R_calc))
    rep.add(Note(R_calc))

    rep.add_heading('计算过程', 2)
    rep.add_paragraph('基本几何参数计算如下：')
    rep.add(Procedure(G_calc))
    rep.add_paragraph('根据规范5.3.6，端阻系数和侧阻系数计算如下：')
    rep.add(Procedure(P_calc))
    rep.add_paragraph('最终计算成果如下：')
    rep.add(Procedure(R_calc))

    rep.save('康苏溢洪洞出口消能建筑物桩计算2.docx')


