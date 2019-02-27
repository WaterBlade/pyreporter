from pyreporter.calculator import (Formula, PiecewiseFormula,
                                   Calculator, TrailSolver, Pr, Sq, Cos, Sin, ATan, Radical, ToRadian,
                                   Variable, SerialVariable, FractionVariable, Unit, Number, FlatDiv, Sum)
from pyreporter import Report, DefaultCover, Math, Procedure, Definition, Note, VariableValue
import math

V = Variable
SV = SerialVariable
FV = FractionVariable
U = Unit
N = Number
Value = VariableValue


Fe = V('F', 'E', inform='地震主动土压力代表值', unit=U('kN'))
q0 = V('q', '0', inform='土表面单位长度的荷重', unit=U('kPa'))
ψ1 = V('ψ', '1', inform='挡土墙面与垂直面的夹角', unit=U('rad'))
ψ2 = V('ψ', '2', inform='挡土墙面与水平面的夹角', unit=U('rad'))
H = V('H', inform='土的高度', unit=U('m'))
γ = V('γ', inform='土的重度标准值', unit=U('kN')/U('m')**3)
φ = V('φ', inform='土的内摩擦角', unit=U('rad'))
θe = V('θ', 'e', inform='地震系数角', unit=U('rad'))
δ = V('δ', inform='挡土墙面与图之间的摩擦角', unit=U('rad'))
ζ = V('ζ', inform='计算系数')
av = V('a', 'v', inform='竖向设计地震加速度代表值', unit=U('m')/U('s')**2)
ah = V('a', 'h', inform='水平向设计地震加速度代表值', unit=U('m')/U('s')**2)
g = V('g', value=9.81, inform='重力加速度', unit=U('m')/U('s')**2)
Ce = V('C', 'e', inform='中间计算参数')
Z = V('Z', inform='中间计算参数')

F1_calc = Calculator()
F1_calc.add(Formula(Fe, Sq(q0*(Cos(ψ1)/Cos(ψ1-ψ2))*H+N(1)/2*γ*H**2)*Pr(1-ζ*av/g)*Ce, long=True))

F2_calc = Calculator()
F2_calc.add(Formula(Fe, Sq(q0*(Cos(ψ1)/Cos(ψ1-ψ2))*H+N(1)/2*γ*H**2)*Pr(1+ζ*av/g)*Ce, long=True))

P_calc = Calculator()
P_calc.add(Formula(Ce, Cos(φ-θe-ψ1)**2/(Cos(θe)*Cos(ψ1)**2*Cos(δ+ψ1+θe)*Pr(1+Radical(Z, 2))**2), long=True))
P_calc.add(Formula(Z, (Sin(δ+φ)*Sin(φ-θe-ψ2))/(Cos(δ+ψ1+θe)*Cos(ψ2-ψ1)), long=True))

A1_calc = Calculator()
A1_calc.add(Formula(θe, ATan((ζ*ah)/(g-ζ*av))))

A2_calc = Calculator()
A2_calc.add(Formula(θe, ATan((ζ*ah)/(g+ζ*av))))

A = V('A', inform='挡土墙截面积', unit=U('m')**2)
L = V('L', inform='挡土墙段长度', unit=U('m'))
Ei = V('E', 'i', inform='作用在质点i的水平向地震惯性力代表值', unit=U('kN'))
ξ = V('ξ', inform='地震作用的效应折减数值', value=0.25)
GEi = V('G', 'Ei', inform='集中在质点i的重力作用标准值', unit=U('kN'))
γc = V('γ', 'c', inform='混凝土的重度标准值', unit=U('kN')/U('m')**3)
αi = V('α', 'i', inform='质点i的地震惯性力的动态分布系数')

Ei_calc = Calculator()
Ei_calc.add(Formula(Ei, ah * ξ * GEi * αi / g))

GEi_calc = Calculator()
GEi_calc.add(Formula(GEi, γc * A * L))

n = V('n', inform='挡土墙建基面内桩的数量', precision=0)
b = V('b', inform='挡土墙宽度', unit=U('m'))
Ni = V('N', inform='计算截面处竖向荷载', unit=U('kN'))
Mi = V('M', inform='计算截面处的弯矩', unit=U('kN')*U('m'))
Hi = V('H', inform='计算截面处水平荷载', unit=U('kN'))

Fn_calc = Calculator(sequence=False)
Fn_calc.add(Formula(Ni, GEi/n))
Fn_calc.add(Formula(Hi, (Ei+Fe*b)/n))
Fn_calc.add(Formula(Mi, (Ei*H+Fe*b*(H/3))/n))


if __name__ == '__main__':
    rep = Report()
    rep.set_cover(DefaultCover('康苏水库工程', '溢洪洞出口桩基水平荷载计算', '水  工', '技  施'))
    q0.set(0)
    ψ1.set(0)
    ψ2.set(0)
    H.set(13)
    γ.set(20)
    φ.set(ToRadian(40))
    δ.set(φ/3)
    ζ.set(0.35)
    ah.set(0.4*g)
    av.set(2/3*ah)
    A.set(47.03)
    L.set(5.2)
    γc.set(25)
    αi.set(2.0)
    n.set(9)
    b.set(8.26)

    F_list = list()

    # 1
    A1_calc.calc()
    P_calc.calc()
    F_list.append(F1_calc.calc())
    A1_res_1 = Procedure(A1_calc)
    P_res_1 = Procedure(P_calc)
    F1_res_1 = Procedure(F1_calc)
    # 2
    F_list.append(F2_calc.calc())
    F2_res_2 = Procedure(F2_calc)
    # 3
    A2_calc.calc()
    P_calc.calc()
    F_list.append(F1_calc.calc())
    A2_res_3 = Procedure(A1_calc)
    P_res_3 = Procedure(P_calc)
    F1_res_3 = Procedure(F1_calc)
    # 4
    F_list.append(F2_calc.calc())
    F2_res_4 = Procedure(F2_calc)

    Fe.set(max(F_list))

    GEi_calc.calc()
    Ei_calc.calc()

    Fn_calc.calc()

    rep.add_heading('计算基本参数及结果', 1)
    rep.add_heading('基本参数', 2)
    rep.add_paragraph('土体高度：', Value(H))
    rep.add_paragraph('土体重度：', Value(γ))
    rep.add_paragraph('土的内摩擦角：', Value(φ))
    rep.add_paragraph('水平向设计地震加速度代表值', Math(ah==0.4*g))
    rep.add_paragraph('挡土墙截面面积', Value(A))
    rep.add_paragraph('挡土墙长度', Value(L))
    rep.add_paragraph('挡土墙宽度', Value(b))
    rep.add_paragraph('挡土墙内桩基数量', Value(n))
    rep.add_paragraph('混凝土重度', Value(γc))
    rep.add_paragraph('挡土墙地震的惯性力分布系数', Value(αi))

    rep.add_heading('计算结果', 2)
    rep.add_paragraph('每延米的动土压力为：', Value(Fe))
    rep.add_paragraph('挡土墙地震惯性力为：', Value(Ei))
    rep.add_paragraph('挡土墙重力标准值为：', Value(GEi))
    rep.add_paragraph('计算截面处竖向荷载：', Value(Ni))
    rep.add_paragraph('计算截面处水平荷载：', Value(Hi))
    rep.add_paragraph('计算截面处弯矩：', Value(Mi))

    rep.add_heading('地震动土压力计算', 1)
    rep.add_heading('计算公式', 2)

    rep.add_paragraph('根据规范《水工建筑物抗震设计规范 NB 35047-2015》5.9.1条，动土压力按以下两式分别计算，取大值：')
    rep.add(Definition(F1_calc))
    rep.add(Definition(F2_calc))
    rep.add(Note(F1_calc))
    rep.add_paragraph('计算参数定义如下：')
    rep.add(Definition(P_calc))
    rep.add(Note(P_calc))
    rep.add_paragraph('其中的地震系数角分别按下式计算，以最终动土压力计算结果决定取用：')
    rep.add(Definition(A1_calc))
    rep.add(Definition(A2_calc))
    rep.add(Note(A1_calc))

    rep.add_heading('计算过程', 2)
    rep.add_heading('计算情况一', 3)
    rep.add_paragraph('地震角计算：')
    rep.add(A1_res_1)
    rep.add_paragraph('中间参数计算：')
    rep.add(P_res_1)
    rep.add_paragraph('动土压力计算：')
    rep.add(F1_res_1)

    rep.add_heading('计算情况二', 3)
    rep.add_paragraph('按另一公式计算动土压力：')
    rep.add(F2_res_2)

    rep.add_heading('计算情况三', 3)
    rep.add_paragraph('按另一公式计算地震角：')
    rep.add(A2_res_3)
    rep.add_paragraph('中间参数计算：')
    rep.add(P_res_3)
    rep.add_paragraph('动土压力计算：')
    rep.add(F1_res_3)

    rep.add_heading('计算情况四', 3)
    rep.add_paragraph('按另一公式计算动土压力：')
    rep.add(F2_res_4)

    rep.add_heading('汇总', 3)
    rep.add_table([['计算情况', '动土压力'],
                   ['情况一', f'{F_list[0]:.3f}'],
                   ['情况二', f'{F_list[1]:.3f}'],
                   ['情况三', f'{F_list[2]:.3f}'],
                   ['情况四', f'{F_list[3]:.3f}']],
                  title='动土压力计算成果')
    rep.add_paragraph('最终的动土压力取值为：', Value(Fe))

    rep.add_heading('挡土墙地震惯性力计算', 1)
    rep.add_heading('计算公式', 2)
    rep.add_paragraph('根据规范《水工建筑物抗震设计规范 NB 35047-2015》5.5.9条，建筑物地震惯性力代表值为：')
    rep.add(Definition(Ei_calc))
    rep.add(Note(Ei_calc))
    rep.add_paragraph('挡土墙重力作用标准值按下式计算:')
    rep.add(Definition(GEi_calc))
    rep.add(Note(GEi_calc))

    rep.add_heading('计算过程', 2)
    rep.add_paragraph('挡土墙重力标准值：')
    rep.add(Procedure(GEi_calc))
    rep.add_paragraph('根据规范《水工建筑物抗震设计规范 NB 35047-2015》9.1.3条，考虑取上下端平均值，挡土墙地震惯性力的动态分布系数为：',
                      Value(αi))
    rep.add_paragraph('挡土墙地震惯性力代表值：')
    rep.add(Procedure(Ei_calc))

    rep.add_heading('单个桩基荷载计算', 1)
    rep.add_heading('计算公式', 2)
    rep.add(Definition(Fn_calc))
    rep.add(Note(Fn_calc))
    rep.add_heading('计算过程', 2)
    rep.add(Procedure(Fn_calc))

    rep.save('康苏溢洪洞出口消能桩基荷载计算.docx')
