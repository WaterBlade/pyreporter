from pyreporter.docx import DocX
from pyreporter.reporter import (Report, ReportComposite, MathComposite)
from pyreporter.calculator import Variable, Radical


MC = MathComposite
RC = ReportComposite
V = Variable


if __name__ == '__main__':
    rep = Report()
    rep.set_default_cover()
    rep.add_heading('标题一', level=1)
    rep.add_paragraph('hello world!')
    rep.add_heading('小标题', level=2)
    rep.add_paragraph('hello world again')
    rep.add_heading('标题二', level=1)
    t1 = rep.add_table('这是一个表格', [[1, 2, 3], [4, 5, 6]])
    rep.add_figure('这是一个图片', '桌面.png', 50)
    rep.add_figure('这也是一个图片', '桌面.jpg', 50)
    rep.add_paragraph('根据', t1)
    rep.add_named_math(V('a') * V('b') / Radical(V('c') + 2, 2))
    rep.add_mul_named_math(MC('a', '=', V('a') * V('b') / Radical(V('c') + 2, 2)),
                           V('a') * V('b') / Radical(V('c') + 2, 2),
                           V('a') * V('b') / Radical(V('c') + 2, 2))

    rep.visit(DocX('hello_world.docx'))
