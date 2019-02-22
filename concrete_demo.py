import pyreporter as rpt
import pyreporter.calculator as calc

V = calc.Variable
SV = calc.SerialVariable
FV = calc.FractionVariable
Num = calc.Number
U = calc.Unit
Pr = calc.Pr

K = V('K', inform='承载力安全系数')
N = V('N', inform='轴向拉力设计值', unit=U('N'))
M = V('M', inform='弯矩设计值', unit=U('N')*U('mm'))
fy = V('f', 'y', inform='纵向钢筋抗拉强度设计值', unit=U('N') / U('mm') ** 2)
fy_ = V("f'", 'y', inform='纵向钢筋抗拉强度设计值', unit=U('N') / U('mm') ** 2)
fc = V('f', 'c', inform='混凝土轴心抗压强度设计值', unit=U('N') / U('mm') ** 2)
As = V('A', 's', inform='配置在靠近轴向拉力一侧的纵向钢筋截面面积', unit=U('mm') ** 2)
As_ = V("A'", 's', inform='配置在远离轴向拉力一侧的纵向钢筋截面面积', unit=U('mm') ** 2)
e = V('e', inform=rpt.Content('轴向拉力至钢筋', rpt.Math(V('A', 's')), '合力点之间的距离'),
      unit=U('mm'))
e_ = V('e', inform=rpt.Content('轴向拉力至钢筋', rpt.Math(V("A'", 's')), '合力点之间的距离'),
       unit=U('mm'))
a_s = V('a', 's', inform=rpt.Content('钢筋', rpt.Math(V('A', 's')), '侧保护层厚度'),
        unit=U('mm'))
a_s_ = V("a'", 's', inform=rpt.Content('钢筋', rpt.Math(V('A', 's')), '侧保护层厚度'),
         unit=U('mm'))
h0 = V('h', '0', inform=rpt.Content('钢筋', rpt.Math(V('A', 's')), '对应的截面计算高度'),
       unit=U('mm'))
h0_ = V("h'", '0', inform=rpt.Content('钢筋', rpt.Math(V('A', 's')), '对应的截面计算高度'),
        unit=U('mm'))
b = V('b', inform='截面计算高度', unit=U('mm'))
x = V('x', inform='截面受压区高度', unit=U('mm'))

S_check = calc.Equations()
S_check.add(K * N * e <= fy * As_ * Pr(h0 - a_s_))
S_check.add(K * N * e_ <= fy * As * Pr(h0_ - a_s))

L_check1 = calc.Equations()
L_check1.add(K * N <= fy * As - fy_ * As_ - fc * b * x)
L_check1.add(K * N * e <= fc * b * x * Pr(h0 - x / 2) + fy_ * As_ * Pr(h0 - a_s_))

L_check2 = calc.Equations()
L_check2.add(K * N <= fy * As - fy_ * As_ - fc * b * x)
L_check2.add(K * N * e_ <= fy * As * Pr(h0_ - a_s))

x_calc = calc.Formula(x, (N - fy * As + fy_ * As_)/(fc*b), long=True)

x_as_check = calc.Equation(x >= 2*a_s_)

x_max_check = calc.Equation(x <= 0.85)


def large_design(level):
    blk = rpt.Block()
    blk.add_heading('判别大小偏心', level=level+1)
