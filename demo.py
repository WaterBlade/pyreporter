from pyreporter.docx import DocX
from pyreporter.reporter import (Reporter, P, MP, Run, InlineFigure, StandaloneFigure, Table, Math,
                                 add_default_cover)
from pyreporter.calc import V, Radical

if __name__ == '__main__':
    rep = Reporter()
    # rep.add(Run('hello'))
    add_default_cover(rep)
    # rep.add(Run('hello world', size=96))
    # rep.add(MP(V('a')+V('b')+Radical(V('c')+2, 2)))
    # rep.add(InlineFigure('桌面.png', 0.5))
    # rep.add(StandaloneFigure('桌面.jpg', 0.2))
    # t = Table([['1', '2', '3'], ['4', '5', Math(V('a')+V('b'))]])
    # t.set_grid_col([1000, 1000, 1000])
    # t.set_border(top=True, bottom=True, inside_h=True)
    # rep.add(t)

    rep.write_to(DocX('hello_world.docx'))
