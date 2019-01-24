from pyreporter import (Report, DefaultCover, Paragraph, Content, Math,
                        Definition, Procedure, Note, VariableValue)
from pyreporter.calculator import (Variable, FractionVariable, SerialVariable, Number, Unit,
                                   Pr, FlatDiv, Sq,
                                   Formula, PiecewiseFormula, Calculator,
                                   Cos, ACos, Sin, ASin, ATan, Sum, Radical)
import math

V = Variable
SV = SerialVariable
FV = FractionVariable
N = Number
U = Unit
Value = VariableValue

# 常数
Pi = V('π', value=math.pi, inform='圆周率')

# 截面参数计算
t = V('t', inform='槽壳厚度', unit=U('m'))
t0 = V('t', '0', inform='槽底加厚厚度', unit=U('m'))
d0 = V('d', '0', inform='槽底加厚宽度', unit=U('m'))
R0 = V('R', '0', inform='槽壳内径', unit=U('m'))
f = V('f', inform='直段高度', unit=U('m'))
a = V('a', inform='顶梁外挑距离', unit=U('m'))
b = V('b', inform='顶梁直段高度', unit=U('m'))
c = V('c', inform='顶梁斜段高度', unit=U('m'))

B = V('B', inform='顶梁换算高度', unit=U('m'))
R = V('R', inform='槽壳中心半径', unit=U('m'))
R1 = V('R', '1', inform='槽壳外径', unit=U('m'))
φq = V('φ', 'q', inform='切点对应的圆心角', unit=U('rad'))
s0 = V('s', '0', inform='槽底加厚斜段宽度', unit=U('m'))

Pa_calc = Calculator(sequence=False)
Pa_calc.add(Formula(B, (b+c)/2))
Pa_calc.add(Formula(R, R0+t/2))
Pa_calc.add(Formula(R1, R0+t))
Pa_calc.add(Formula(φq,
                    Pi/2 - ATan(d0/(2*Pr(R1+t0))) - ACos(R1/Radical(Pr(d0/2)**2+Pr(R1+t0)**2, 2)),
                    long=True))
Pa_calc.add(Formula(s0, R1*Cos(φq)-d0/2))


# 分块计算
Iti = SV('I', 't', unit=U('m')**4, inform='各分块惯性矩')
Ai = SV('A', unit=U('m')**2, inform='各分块截面宽度')
yci = SV('y', 'c', unit=U('m'), inform='各分块截面重心至槽顶距离')
Split_calc = Calculator(sequence=False)

It1 = Iti.new('顶梁外挑矩形部分惯性矩')
A1 = Ai.new('顶梁外挑矩形部分截面面积')
yc1 = yci.new('顶梁外挑矩形部分重心至槽顶距离')

Split_calc.add(Formula(yc1, b / 2))
Split_calc.add(Formula(A1, 2 * a * b))
Split_calc.add(Formula(It1, 2 * (a * b ** 3 / 12)))

It2 = Iti.new('顶梁外挑三角形部分惯性矩')
A2 = Ai.new('顶梁外挑三角形部分截面面积')
yc2 = yci.new('顶梁外挑三角形部分重心至槽顶距离')

Split_calc.add(Formula(yc2, b + c / 3))
Split_calc.add(Formula(A2, 2 * (a * c / 2)))
Split_calc.add(Formula(It2, 2 * (a * c ** 3 / 36)))

It3 = Iti.new('直段惯性矩')
A3 = Ai.new('直段截面面积')
yc3 = yci.new('直段重心至槽顶距离')

Split_calc.add(Formula(yc3, f / 2))
Split_calc.add(Formula(A3, 2 * f * t))
Split_calc.add(Formula(It3, 2 * (t * f ** 3 / 12)))

It4 = Iti.new('弧段半圆惯性矩')
A4 = Ai.new('弧段半圆面积')
yc4 = yci.new('弧段半圆重心至槽顶距离')

Split_calc.add(Formula(yc4, f + 4 * R1 / (3 * Pi)))
Split_calc.add(Formula(A4, Pi * R1 ** 2 / 2))
Split_calc.add(Formula(It4, Pr(9 * Pi ** 2 - 64) * R1 ** 4 / (72 * Pi)))

It5 = Iti.new('弧段空腔半圆惯性矩')
A5 = Ai.new('弧段空腔半圆面积')
yc5 = yci.new('弧段空腔半圆重心至槽顶距离')

Split_calc.add(Formula(yc5, f + 4 * R0 / (3 * Pi)))
Split_calc.add(Formula(A5, -(Pi * R0 ** 2 / 2)))
Split_calc.add(Formula(It5, -(Pr(9 * Pi ** 2 - 64) * R0 ** 4 / (72 * Pi))))

It6 = Iti.new('槽底加厚矩形部分惯性矩')
A6 = Ai.new('槽底加厚矩形部分面积')
yc6 = yci.new('槽底加厚矩形部分重心至槽顶距离')

Split_calc.add(Formula(yc6, f + R1 + t0 / 2))
Split_calc.add(Formula(A6, t0 * d0))
Split_calc.add(Formula(It6, d0 * t0 ** 3 / 12))

It7 = Iti.new('槽底加厚三角形部分惯性矩')
A7 = Ai.new('槽底加厚三角形部分面积')
yc7 = yci.new('槽底加厚三角形部分重心至槽顶距离')

Split_calc.add(Formula(yc7, f + R1))
Split_calc.add(Formula(A7, 2 * (t0 * s0 / 2)))
Split_calc.add(Formula(It7, 2 * (s0 * t0 ** 3 / 36)))


# 截面特性计算
A = V('A', inform='横截面的总面积', unit=U('m') ** 2)
y1 = V('y', '1', inform='形心轴到槽身顶部距离', unit=U('m'))
y2 = V('y', '2', inform='形心轴到槽身底部距离', unit=U('m'))
K = V('K', inform='形心轴至圆形轴距离', unit=U('m'))
I = V('I', inform='横截面对形心轴的惯性矩', unit=U('m') ** 4)
φ0 = V('φ', '0', inform='形心轴位置对应的夹角', unit=U('rad'))
Sl = V('S', 'l', inform='受拉区面积对截面形心轴的静面矩', unit=U('m') ** 3)


Geo_calc = Calculator(sequence=False)
Geo_calc.add(Formula(A, Sum(Ai, [Ai])))
Geo_calc.add(Formula(y1, Sum(Ai * yci, [Ai, yci]) / A, long=True))
Geo_calc.add(Formula(y2, f + R1 + t0 - y1))
Geo_calc.add(Formula(K, y1 - f))
Geo_calc.add(Formula(I, Sum(Ai * Pr(y1 - yci) ** 2, [Ai, yci]) + Sum(Iti, [Iti]), long=True))
Geo_calc.add(Formula(φ0, ACos(K / R)))
Geo_calc.add(Formula(Sl,
                     2 * t * R ** 2 * Pr(Sin(φ0) - φ0 * Cos(φ0))
                     + t0 * d0 * Pr(y2 - t0 / 2)
                     + 2 * t0 * s0 * Pr(y2 - t0) / 2,
                     long=True))

# U形槽纵向内力计算
q = V('q', inform='每米槽壳长度内的所有荷载之和', unit=U('kN')/U('m'))
L = V('L', inform='槽身计算跨径', unit=U('m'))
M = V('M', inform='简支梁跨中弯矩', unit=U('kN') * U('m'))
Q = V('Q', inform='简支梁梁端剪力', unit=U('kN'))

F_calc = Calculator(sequence=False)
F_calc.add(Formula(M, q * L ** 2 / 8))
F_calc.add(Formula(Q, q * L / 2))

# U形槽纵向结构配筋计算
Zl = V('Z', 'l', inform='简支梁计算受拉区的总拉力', unit=U('kN'))
Ag = V('A', 'g', inform='受拉钢筋总面积', unit=U('mm') ** 2)
fy = V('f', 'y', inform='钢筋受拉强度设计值', unit=U('N') / U('mm') ** 2)
Kg = V('K', 'g', inform='结构安全系数')

St_calc = Calculator()
St_calc.add(Formula(Ag, Kg * (1000 * Zl / fy)))
St_calc.add(Formula(Zl, M / I * Sl))

# U形槽正截面抗裂验算
σl = V('σ', 'l', inform='正截面最大拉应力', unit=U('kPa'))
ft = V('f', 'c', inform='混凝土抗拉强度')

# U形槽横向结构计算
# 剪力计算
T = V('T', inform='直段上的总剪力', unit=U('kN'))
T1 = V('T', '1', inform='直段顶部加大部分剪力', unit=U('kN'))
T2 = V('T', '2', inform='直段加大部分以下的直段剪力', unit=U('kN'))

T_calc = Calculator()
T_calc.add(Formula(T, T1 + T2))
T_calc.add(Formula(T1, q / I * Pr(y1 * B / 2 - B ** 3 / 6) * Pr(t + a), long=True))
T_calc.add(Formula(T2,
                   q / I * Sq(t * y1 * Pr(f ** 2 / 2 - B * f + B ** 2 / 2)
                              - t * Pr(f ** 3 / 6 - B ** 2 * f / 2 + B ** 3 / 3)
                              + Pr(t + a) * Pr(y1 * B - B ** 2 / 2) * Pr(f - B)),
                   long=True))

X1_calc = Calculator()
D_calc = Calculator(sequence=False)
P_calc = Calculator(sequence=False)

X1 = V('X', '1', inform='多余未知力')
δ11 = V('δ', '11', inform='单位多余未知力作用时对应的水平变位')
Δ1P = V('Δ', '1P', inform='所有荷载引起的水平变位')
Δ1G0 = V('Δ', V('1G', '0'), inform='附加集中力引起的水平变位')
Δ1M0 = V('Δ', V('1M', '0'), inform='附加弯矩引起的水平变位')
Δ1h = V('Δ', '1h', inform='槽壳自重引起的水平变位')
Δ1W = V('Δ', '1W', inform='水压力引起的水平变位')
Δ1τ = V('Δ', '1τ', inform='剪应力引起的水平变位')
E = V('E', inform='混凝土弹性模量', unit=U('kPa'))
It = V('I', 't', inform='槽壳横向惯性矩', unit=U('m') ** 4)
β = V('β', inform='计算中间参数，直段高度与半径比值')
h = V('h', inform='圆心轴至拉杆中心高度', unit=U('m'))
λ = V('λ', inform='计算中间参数，形心距与半径比值')

X1_calc.add(Formula(X1, Δ1P / δ11))
X1_calc.add(Formula(Δ1P, Δ1G0 + Δ1M0 + Δ1h + Δ1W + Δ1τ, long=True))
X1_calc.add(Formula(δ11, R ** 3 / (E * It) * Pr(0.333 * β ** 3 + 1.571 * β ** 2 + 2 * β + 0.785), long=True))

P_calc.add(Formula(λ, K / R))
P_calc.add(Formula(β, h / R))
P_calc.add(Formula(It, 1 * t ** 3 / 12))

G0 = V('G', '0', inform='附加集中力', unit=U('kN'))
M0 = V('M', '0', inform='附加弯矩', unit=U('kN') * U('m'))
γh = V('γ', 'h', inform='混凝土重度', unit=FlatDiv(U('kN'), U('m') ** 3))
γ = V('γ', inform='水的重度', unit=FlatDiv(U('kN'), U('m') ** 3))
h1 = V('h', '1', inform='圆形轴至水面高度', unit=U('m'))

D_calc.add(Formula(Δ1G0, -((G0 * R ** 3) / (E * It) * Pr(0.571 * β + 0.5)), long=True))
D_calc.add(Formula(Δ1M0, (M0 * R ** 2) / (E * It) * Pr(0.5 * β ** 2 + 1.57 * β + 1), long=True))
D_calc.add(Formula(Δ1h, -((γh * t * R ** 4) / (E * It) * Pr(0.571 * β ** 2 + 0.929 * β + 0.393)), long=True))
D_calc.add(Formula(Δ1W,
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
D_calc.add(Formula(Δ1τ,
                   1 / (E * It) * ((q * t) / I) * R ** 6 * Pr(0.214 * β - 0.294 * λ * β - 0.265 * λ + 0.197)
                   + (T * R ** 3) / (E * It) * Pr(0.571 * β + 0.5)
                   + T1 / (E * It) * ((a * R ** 2) / 2) * Pr(0.5 * β ** 2 + 1.57 * β + 1), long=True))

# 横槽向内力计算
# 弯矩
# 直段内力

My = V('M', 'y', inform='直段弯矩', unit=U('kN') * U('m'))
y = V('y', inform='直段计算内力截面位置离顶端的距离', unit=U('m'))

My_calc = Calculator()
My_calc.add(PiecewiseFormula(My,
                             [M0 + 1 / N(2) * a * T1 + X1 * y,
                              M0 + 1 / N(2) * a * T1 - 1 / N(6) * γ * Sq(y - Pr(h - h1)) ** 3 + X1 * y],
                             [y <= h - h1,
                              y > h - h1],
                             [False, True]))

Mj = V('M', 'φ', inform='圆弧段弯矩', unit=U('kN') * U('m'))
φ = V('φ', inform='圆弧段计算内力截面位置对应的圆心角', unit=U('rad'))
MM0 = V('M', V('M', '0'), inform='附加弯矩引起的截面弯矩', unit=U('kN') * U('m'))
MG0 = V('M', V('G', '0'), inform='附加力引起的截面弯矩', unit=U('kN') * U('m'))
Mh = V('M', 'h', inform='槽身自重引起的截面弯矩', unit=U('kN') * U('m'))
MW = V('M', 'W', inform='水压力引起的截面弯矩', unit=U('kN') * U('m'))
Mτ = V('M', 'τ', inform='剪应力引起的截面弯矩', unit=U('kN') * U('m'))
MX1 = V('M', V('X', '1'), inform='多余未知力引起的截面弯矩', unit=U('kN') * U('m'))

Mj_calc = Calculator()
Mj_calc.add(Formula(Mj,
                    MM0 + MG0 + Mh + MW + Mτ + MX1,
                    long=True))

Mji_calc = Calculator(sequence=False)
Mji_calc.add(Formula(MM0, M0))
Mji_calc.add(Formula(MG0, -G0 * R * Pr(1 - Cos(φ))))
Mji_calc.add(Formula(Mh,
                     -γh * t * R ** 2 * Sq(f / R * Pr(1 - Cos(φ))
                                           + Sin(φ)
                                           - φ * Cos(φ)),
                     long=True))
Mji_calc.add(Formula(MW,
                     -γ * Sq(1 / N(2) * Pr(h1 ** 2 * R + R * R0 ** 2) * Sin(φ)
                             - Pr(1 / N(2) * R * R0 ** 2 * φ + R * R0 * h1) * Cos(φ)
                             + 1 / N(6) * h1 ** 3
                             + R * R0 * h1),
                     long=True))
Mji_calc.add(Formula(Mτ,
                     q * t / (2 * I) * R ** 4 * Sq(Sin(φ) - φ * Cos(φ)
                                                   + λ * Pr(φ ** 2 - Pi * φ + 2 * Cos(φ) + Pi * Sin(φ) - 2))
                     + T * R * Pr(1 - Cos(φ))
                     + 1 / N(2) * a * T1,
                     long=True))
Mji_calc.add(Formula(MX1,
                     X1 * Pr(h + R * Sin(φ))))


def value_para(var_list, def_=None):
    p = Paragraph('相关参数取值为：')
    for var in var_list[:-1]:
        p.extend(Value(var), '、')
    if def_ is not None:
        p.extend(Value(var_list[-1]),
                 '。代入',
                 def_.reference,
                 '进行计算，计算结果及过程如下：')
    else:
        p.extend(Value(var_list[-1]), '代入相关计算式进行计算，计算结果及过程如下：')
    return p


if __name__ == '__main__':
    rep = Report()
    rep.set_cover(DefaultCover())

    rep.add_heading('横截面参数计算', level=1)
    rep.add_heading('横截面尺寸参数计算', level=2)
    rep.add(Definition(Pa_calc))
    rep.add(Note(Pa_calc))

    t.set(0.14)
    R0.set(1.85)
    f.set(0.75)
    a.set(0.26)
    b.set(0.2)
    c.set(0.1)
    t0.set(0.14)
    d0.set(1.0)
    Pa_calc.calc()

    rep.add_paragraph('计算结果及过程如下：')
    rep.add(Procedure(Pa_calc))

    rep.add_heading('横截面分块几何特性计算', level=2)

    rep.add(Definition(Split_calc))
    rep.add(Note(Split_calc))

    Split_calc.calc()

    rep.add_paragraph('计算结果及过程如下：')
    rep.add(Procedure(Split_calc))

    rep.add_heading('横截面几何特性计算', level=2)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，基本计算公式如下：')
    geo_def = Definition(Geo_calc)
    rep.add(geo_def)
    rep.add(Note(Geo_calc))

    Geo_calc.calc()

    rep.add_paragraph('计算结果及过程如下：')
    rep.add(Procedure(Geo_calc))

    rep.add_heading('纵向结构计算', level=1)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，基本计算公式如下：')
    f_def = Definition(F_calc)
    rep.add(f_def)
    rep.add(Note(F_calc))

    q.set(119.46)
    L.set(11.7)
    F_calc.calc()

    rep.add(value_para([q, L], f_def))
    rep.add(Procedure(F_calc))

    rep.add_paragraph('纵向受拉钢筋配筋按下式计算：')
    st_def = Definition(St_calc)
    rep.add(st_def)
    rep.add(Note(St_calc))

    Kg.set(1.1)
    fy.set(300)
    St_calc.calc()

    rep.add(value_para([Kg, fy], st_def))
    rep.add(Procedure(St_calc))

    rep.add_heading('横向结构计算', level=1)
    rep.add_paragraph('横向结构计算采用力法进行。')

    rep.add_heading('荷载计算', level=2)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，基本计算公式如下：')
    t_def = Definition(T_calc)
    rep.add(t_def)
    rep.add(Note(T_calc))

    T_calc.calc()

    rep.add_paragraph('带入相关参数，计算结果及过程如下：')
    rep.add(Procedure(T_calc))

    rep.add_heading('结构内力计算', level=2)
    rep.add_heading('力法体系多余未知力计算', level=3)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，基本计算公式如下：')
    rep.add(Definition(X1_calc))
    rep.add(Note(X1_calc))

    rep.add_paragraph('各荷载引起的水平位移计算公式如下：')
    rep.add(Definition(D_calc))
    rep.add(Note(D_calc))

    rep.add_paragraph('计算中间参数定义如下：')
    rep.add(Definition(P_calc))
    rep.add(Note(P_calc))

    E.set(3.0e7)
    M0.set(1.606)
    G0.set(10)
    h.set(0.65)
    h1.set(0.5)
    γ.set(10)
    γh.set(25)

    P_calc.calc()
    D_calc.calc()
    X1_calc.calc()

    rep.add(value_para([E, M0, G0, h, h1]))
    rep.add_paragraph('中间计算参数的计算结果及过程如下：')
    rep.add(Procedure(P_calc))
    rep.add_paragraph('各荷载引起的水平位移计算结果及过程如下：')
    rep.add(Procedure(D_calc))
    rep.add_paragraph('多余位置力计算结果及过程如下：')
    rep.add(Procedure(X1_calc))

    rep.add_heading('截面内力计算', level=3)

    rep.add_heading('直段内力计算', level=4)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，直段内力的计算公式如下：')
    rep.add(Definition(My_calc))
    rep.add(Note(My_calc))

    rep.add_paragraph('带入不同值进行计算。')

    for k in range(5):
        y.set(h * k / 4)
        My_calc.calc()
        rep.add_paragraph('当', Value(y), '时，直段内弯矩结算结果及过程如下：')
        rep.add(Procedure(My_calc))

    # TODO:考虑添加一系列计算功能，针对一个参数范围，计算多个结果。

    rep.add_heading('弧段内力计算', level=4)
    rep.add_paragraph('根据《取水输水建筑物丛书——渡槽》，直段内力的计算公式如下：')
    rep.add(Definition(Mj_calc))
    rep.add(Note(Mj_calc))
    rep.add_paragraph('各弯矩分量计算式如下：')
    rep.add(Definition(Mji_calc))
    rep.add(Note(Mji_calc))

    rep.add_paragraph('带入不同值进行计算。')

    value_list = list()
    for k in range(10):
        phi = math.pi * k / 18
        φ.set(phi)
        Mji_calc.calc()
        value = Mj_calc.calc()
        rep.add_heading(Content(Value(φ), '截面弯矩'), level=5)
        rep.add_paragraph('当', Value(φ), '时，弧段内弯矩各分量结算结果及过程如下：')
        rep.add(Procedure(Mji_calc))
        rep.add_paragraph('截面弯矩计算结果为：')
        rep.add(Procedure(Mj_calc))
        value_list.append([f'{phi:.2f}', f'{value:.2f}'])

    rep.add_paragraph('结果汇总如下：')
    rep.add_table([[Content('角度（', Math(Unit('rad')), '）'), Content('截面弯矩（', Math(Unit('kN')*Unit('m')), '）')]]+value_list, title='弧段弯矩计算值')

    rep.save('ushell_demo.docx')
