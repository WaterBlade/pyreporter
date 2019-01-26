import pyreporter as rpt
import pyreporter.calculator as calc


V = calc.Variable
SV = calc.SerialVariable
FV = calc.FractionVariable
Num = calc.Number
U = calc.Unit


K = V('K', inform='承载力安全系数')
N = V('N', inform='轴向拉力设计值', unit=U('N'))
fy = V('f', 'y', inform='纵向钢筋抗拉强度设计值', unit=U('N')/U('mm')**2)
As = V('A', 's', inform='配置在靠近轴向拉力一侧的纵向钢筋截面面积', unit=U('mm')**2)
As_ = V("A'", 's', inform='配置在远离轴向拉力一侧的纵向钢筋截面面积', unit=U('mm')**2)
e = V('e', inform=rpt.Content('轴向拉力至钢筋', rpt.Math(V('A', 's')), '合力点之间的距离'),
      unit=U('mm'))

