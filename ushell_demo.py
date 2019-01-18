from pyreporter import Report, DefaultCover, Definition, Procedure, Note
from pyreporter.calculator import (Variable, FractionVariable, Number, Unit,
                                   Pr, FlatDiv, Sq,
                                   Formula, Calculator)

V = Variable
FV = FractionVariable
N = Number
U = Unit

X1_calc = Calculator()
Dp_calc = Calculator()
Dh_calc = Calculator()
Dw_calc = Calculator()
Dt_calc = Calculator()
P_calc = Calculator()

X1 = V('X', '1', inform='多余未知力')
δ11 = V('δ', '11', inform='单位多余未知力作用时对应的水平变位')
Δ1P = V('Δ', '1P', inform='所有荷载引起的水平变位')
Δ1G0 = V('Δ', V('1G', '0'), inform='附加集中力引起的水平变位')
Δ1M0 = V('Δ', V('1M', '0'), inform='附加弯矩引起的水平变位')
Δ1h = V('Δ', '1h', inform='槽壳自重引起的水平变位')
Δ1W = V('Δ', '1W', inform='水压力引起的水平变位')
Δ1τ = V('Δ', '1τ', inform='剪应力引起的水平变位')
R = V('R', inform='槽壳中心半径', unit=U('m'))
E = V('E', inform='混凝土弹性模量', unit=U('kPa'))
It = V('I', 't', inform='槽壳横向惯性矩', unit=U('m') ** 4)
β = V('β', inform='计算中间参数，直段高度与半径比值')
h = V('h', inform='圆心轴至拉杆中心高度', unit=U('m'))
t = V('t', inform='槽壳厚度', unit=U('m'))
λ = V('λ', inform='计算中间参数，形心距与半径比值')
K = V('K', inform='形心轴至圆形轴距离', unit=U('m'))

X1_calc.add(Formula(X1, Δ1P / δ11))
X1_calc.add(Formula(Δ1P, Δ1G0 + Δ1M0 + Δ1h + Δ1W + Δ1τ))
X1_calc.add(Formula(δ11, R ** 3 / (E * It) * Pr(0.333 * β ** 3 + 1.571 * β ** 2 + 2 * β + 0.785)))

P_calc.add(Formula(λ, K / R))
P_calc.add(Formula(β, h / R))
P_calc.add(Formula(It, 1 * t ** 3 / 12))

G0 = V('G', '0', inform='附加集中力', unit=U('kN'))
M0 = V('M', '0', inform='附加弯矩', unit=U('kN') * U('m'))
γh = V('γ', 'h', inform='混凝土重度', unit=FlatDiv(U('kN'), U('m') ** 3))
γ = V('γ', inform='水的重度', unit=FlatDiv(U('kN'), U('m') ** 3))
h1 = V('h', '1', inform='圆形轴至水面高度', unit=U('m'))
R0 = V('R', '0', inform='槽壳内径', unit=U('m'))
J = V('J', inform='横截面对形心轴的惯性矩', unit=U('m') ** 4)
q = V('q', inform='每米槽壳长度内的所有荷载之和', unit=U('kN'))
T = V('T', inform='直段上的总剪力', unit=U('kN'))
T1 = V('T', '1', inform='直段顶部加大部分剪力', unit=U('kN'))
T2 = V('T', '2', inform='直段加大部分以下的直段剪力', unit=U('kN'))
a = V('a', inform='顶梁外挑距离', unit=U('m'))

Dp_calc.add(Formula(Δ1G0, -((G0 * R ** 3) / (E * It) * Pr(0.571 * β + 0.5)), long=True))
Dp_calc.add(Formula(Δ1M0, (M0 * R ** 2) / (E * It) * Pr(0.5 * β ** 2 + 1.57 * β + 1), long=True))
Dh_calc.add(Formula(Δ1h, -((γh * t * R ** 4) / (E * It) * Pr(0.571 * β ** 2 + 0.929 * β + 0.393)), long=True))
Dw_calc.add(Formula(Δ1W,
                   - (γ / (E * It) *
                      Pr(-N(0.008, 3) * h ** 5
                         + N(0.04, 2) * h ** 4 * h1
                         - 0.082 * h ** 3 * h1 ** 2
                         + 0.083 * h ** 2 * h1 ** 3)
                      - (γ * R) / (E * It) *
                      Sq(h1 ** 3 * Pr(0.262 * h + 0.167 * R)
                         + h1 ** 2 * R * Pr(0.5 * h + 0.393 * R)
                         + h1 * R0 * R * Pr(0.57 * h + 0.5 * R)
                         + R0 ** 2 * R * Pr(0.215 * h + 0.197 * R))), long=True))
Dt_calc.add(Formula(Δ1τ,
                   1 / (E * It) * ((q * t) / J) * R ** 6 * Pr(0.214 * β - 0.294 * λ * β - 0.265 * λ + 0.197)
                   + (T * R ** 3) / (E * It) * Pr(0.571 * β + 0.5)
                   + T1 / (E * It) * ((a * R ** 2) / 2) * Pr(0.5 * β ** 2 + 1.57 * β + 1), long=True))

if __name__ == '__main__':
    rep = Report()
    rep.set_cover(DefaultCover())
    rep.add_heading('横向结构计算', level=1)

    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，基本计算公式如下：')
    rep.add(Definition(X1_calc))
    rep.add(Note(X1_calc))

    rep.add_paragraph('附加力引起的水平位移计算公式如下：')
    rep.add(Definition(Dp_calc))
    rep.add(Note(Dp_calc))

    rep.add_paragraph('槽壳自重引起的水平位移计算公式如下：')
    rep.add(Definition(Dh_calc))
    rep.add(Note(Dh_calc))

    rep.add_paragraph('水压力引起的水平位移计算公式如下：')
    rep.add(Definition(Dw_calc))
    rep.add(Note(Dw_calc))

    rep.add_paragraph('剪应力引起的水平位移计算公式如下：')
    rep.add(Definition(Dt_calc))
    rep.add(Note(Dt_calc))

    rep.add_paragraph('计算中间参数定义如下：')
    rep.add(Definition(P_calc))
    rep.add(Note(P_calc))

    rep.save('ushell_demo.docx')
