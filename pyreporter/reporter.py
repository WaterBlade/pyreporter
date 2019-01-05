from PIL import Image
from io import BytesIO

# TODO: 两个模块之间耦合太多，考虑如何将其分割

class Reporter:
    def __init__(self):
        self.data_list = list()

    def add(self, item):
        if isinstance(item, Run):
            item = Paragraph(item)
        elif isinstance(item, Math):
            item = Paragraph(item)
        elif isinstance(item, InlineFigure):
            item = Paragraph(item)
        self.data_list.append(item)

    def write_to(self, writer):
        writer.write_docx(self)


class Paragraph:
    def __init__(self, *datas,
                 spacing=None,
                 snap_to_grid=False):
        self.datas = datas
        self.spacing = spacing
        self.snap_to_grid = snap_to_grid

    def write_to(self, writer):
        writer.write_paragraph(*self.datas,
                               jc=self.jc,
                               spacing=self.spacing,
                               snap_to_grid=self.snap_to_grid)

    def set_jc(self, jc):
        self.jc = jc

    def set_spacing(self):
        pass


class Run:
    def __init__(self, data, size=None,
                 ascii=None, hAnsi=None, hint=None, cs=None,
                 i=None):
        self.data = data
        self.size = size
        self.ascii = ascii
        self.hAnsi = hAnsi
        self.hint = hint
        self.cs = cs
        self.i = i

    def write_to(self, writer):
        writer.write_run(text=self.data, size=self.size,
                         ascii=self.ascii, hAnsi=self.hAnsi, hint=self.hint, cs=self.cs, i=self.i)


class MathPara:
    def __init__(self, *datas):
        self.datas = datas

    def write_to(self, writer):
        writer.write_mathpara(self.datas)


class Math:
    def __init__(self, *datas):
        self.datas = datas

    def write_to(self, writer):
        writer.write_math(self.datas)


class Table:
    def __init__(self, data_by_rows=None, jc='center', style='a7'):
        self.data_by_rows = self.process_data(data_by_rows)
        self.jc = jc
        self.style = style
        self.layout = 'auto'

        self.pos_pr = None
        self.border = None
        self.grid_col = None
        self.cell_margin = None

    def process_data(self, data_by_rows):
        ret = list()
        for row in data_by_rows:
            r = list()
            for col in row:
                if isinstance(col, Cell):
                    r.append(col)
                else:
                    r.append(Cell(col))
            ret.append(r)
        return ret

    def write_to(self, writer):
        writer.write_table(data=self.data_by_rows,
                           jc=self.jc,
                           style=self.style,
                           pos_pr=self.pos_pr,
                           border=self.border,
                           grid_col=self.grid_col,
                           cell_margin=self.cell_margin,
                           layout=self.layout)

    def set_data(self, data_by_rows):
        self.data_by_rows = self.process_data(data_by_rows)

    def set_pos_pr(self, x=None, y=None, x_spec=None, y_spec=None):
        self.pos_pr = TablePosPr(x, y, x_spec, y_spec)
        self.jc = None

    def set_border(self, top=False, bottom=False,
                   left=False, right=False,
                   inside_v=False, inside_h=False,
                   size=4):
        self.border = TableBorder(top, bottom, left, right, inside_v, inside_h, size)

    def set_grid_col(self, col_size_list: list):
        self.grid_col = TableGridCol(col_size_list)
        self.layout = 'fixed'

    def set_cell_margin(self, top=0, bottom=0, left=0, right=0):
        self.cell_margin = TableCellMargin(top, bottom, left, right)


class Cell:
    def __init__(self, item, h_align='center', v_align='center'):
        if isinstance(item, Paragraph):
            self.para = item
        elif isinstance(item, tuple) or isinstance(item, list):
            self.para = Paragraph(*item)
        else:
            self.para = Paragraph(item)
        self.jc = h_align
        self.v_align = v_align

        self.border = None

    def write_to(self, writer):
        writer.write_cell(para=self.para,
                          jc=self.jc, v_align=self.v_align,
                          border=self.border)

    def set_border(self, top=False, bottom=False, left=False, right=False, size=4):
        self.border = CellBorder(top, bottom, left, right, size)

    def set_align(self, h_align=None, v_align=None):
        if h_align:
            self.jc = h_align
        if v_align:
            self.v_align = v_align

    def set_snap_to_grid(self, snap=True):
        self.snap_to_grid = snap



class Figure:
    def __init__(self, file, scale=1):
        image = Image.open(file)

        size = image.size
        dpi = image.info['dpi']

        self.inch_size = [s / d for s, d in zip(size, dpi)]
        self.format = 'png'

        b = BytesIO()
        image.save(b, self.format)
        self.data = b.getvalue()

        self.scale = scale


class InlineFigure(Figure):
    def write_to(self, writer):
        x, y = self.inch_size
        writer.write_inline_figure(self.data, self.format,
                                   x * self.scale, y * self.scale)


class StandaloneFigure(Figure):
    def write_to(self, writer):
        x, y = self.inch_size
        writer.write_standalone_figure(self.data, self.format,
                                       x * self.scale, y * self.scale)


P = Paragraph
MP = MathPara


# Inner class
class TablePosPr:
    def __init__(self, x, y, x_spec, y_spec):
        self.x = x
        self.y = y
        self.x_spec = x_spec
        self.y_spec = y_spec

    def write_to(self, writer):
        writer.write_table_pos_pr(self.x, self.y, self.x_spec, self.y_spec)


class TableBorder:
    def __init__(self, top, bottom, left, right, inside_v, inside_h, size):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.inside_v = inside_v
        self.inside_h = inside_h
        self.size = size

    def write_to(self, writer):
        writer.write_table_border(self.top, self.bottom, self.left, self.right,
                                  self.inside_v, self.inside_h,
                                  self.size)


class TableGridCol:
    def __init__(self, col_size_list: list):
        self.col_size_list = col_size_list

    def write_to(self, writer):
        writer.write_table_grid_col(self.col_size_list)


class TableCellMargin:
    def __init__(self, top, bottom, left, right):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right

    def write_to(self, writer):
        writer.write_table_cell_margin(self.top, self.bottom, self.left, self.right)


class CellBorder:
    def __init__(self, top, bottom, left, right, size):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.size = size

    def write_to(self, writer):
        writer.write_cell_border(self.top, self.bottom, self.left, self.right, self.size)


class Spacing:
    def __init__(self, before, after):
        self.before = before
        self.after = after

    def write_to(self, writer):
        writer.write_spacing(self.before, self.after)


def add_default_cover(doc: Reporter, project='**工程', name='**计算书',
                      part='**专业', phase='**计算书',
                      number='', secret='',
                      footer_str='湖南省水利水电勘测设计研究总院'):
    lt = Table([['编号', number]])
    lt.set_grid_col([500, 2000])
    lt.set_cell_margin()
    lt.set_pos_pr(x_spec='left', y_spec='top')
    lt.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    doc.add(lt)

    rt = Table([['秘密', secret]])
    rt.set_grid_col([500, 2000])
    rt.set_cell_margin()
    rt.set_pos_pr(x_spec='right', y_spec='top')
    rt.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    doc.add(rt)

    title = Table([[Run('计  算  书', size=72)]])
    title.set_pos_pr(y=3000, x_spec='center')
    doc.add(title)

    cells = dict()
    for txt in [project, part, phase, name, '工程名称', '专业名称', '设计阶段', '计算书名称']:
        c = Cell(Run(txt, hAnsi='楷体', size=30))
        if txt in [project, part, phase, name]:
            c.set_border(bottom=True, size=6)
            c.set_align(h_align='center')
        else:
            c.set_align(h_align='right')
        c.set_align(v_align='bottom')
        cells[txt] = c

    sub = Table([[cells['工程名称'], cells[project]],
                 [cells['专业名称'], cells[part]],
                 [cells['设计阶段'], cells[phase]],
                 [cells['计算书名称'], cells[name]]])
    sub.set_grid_col([1600, 5000])
    sub.set_cell_margin(right=10)
    sub.set_pos_pr(x_spec='center', y=6000)
    doc.add(sub)

    cells = dict()
    for txt in ['审查', '校核', '计算', '年', '月', '日']:
        c = Cell(Run(txt, hAnsi='楷体'))
        c.set_align(v_align='bottom')
        cells[txt] = c

    blk = Cell('')
    blk.set_border(bottom=True, size=6)

    sig = Table([[cells['审查'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']],
                 [cells['校核'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']],
                 [cells['计算'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']]])
    sig.set_grid_col([700, 1500, 300, 900, 300, 600, 300, 600, 300])
    sig.set_cell_margin()
    sig.set_pos_pr(x_spec='center', y=10500)
    doc.add(sig)

    footer = Table([[Run(footer_str, size=24)]])
    footer.set_pos_pr(x_spec='center', y=13500)
    doc.add(footer)
