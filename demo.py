from pyreporter.docx import DocX
from pyreporter.reporter import (Report, P, MP, Run, InlineFigure, StandaloneFigure, Table, Math)
from pyreporter.calc import V, Radical

if __name__ == '__main__':
    rep = Report()
    rep.set_default_cover()
    rep.add_heading('标题一', level=1)
    rep.add_paragraph('hello world!')
    rep.add_heading('小标题', level=2)
    rep.add_paragraph('hello world again')
    rep.add_heading('标题二', level=1)
    rep.add_table('这是一个表格', [[1, 2, 3], [4, 5, 6]])
    rep.add_figure('这是一个图片', '桌面.png', 50)
    rep.add_figure('这也是一个图片', '桌面.jpg', 50)
    # rep.add(Run('hello'))
    # rep.add(Run('hello world', size=96))
    # rep.add(MP(V('a')+V('b')+Radical(V('c')+2, 2)))
    # rep.add(InlineFigure('桌面.png', 0.5))
    # rep.add(StandaloneFigure('桌面.jpg', 0.2))
    # t = Table([['1', '2', '3'], ['4', '5', Math(V('a')+V('b'))]])
    # t.set_grid_col([1000, 1000, 1000])
    # t.set_border(top=True, bottom=True, inside_h=True)
    # rep.add(t)

    rep.write_to(DocX('hello_world.docx'))
