from PIL import Image
from io import BytesIO
from .docx import DocX


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
        elif isinstance(item, str):
            item = Paragraph(Run(item))
        elif isinstance(item, int) or isinstance(item, float):
            item = Paragraph(Run(f'{item}'))

        self.data_list.append(item)

    def write_to(self, writer):
        writer.write_docx(self)


class Paragraph:
    def __init__(self, *datas, style=None):
        self.datas = list()
        for data in datas:
            if isinstance(data, str):
                self.datas.append(Run(data))
            elif isinstance(data, int) or isinstance(data, float):
                self.datas.append(Run(f'{data}'))
            else:
                self.datas.append(data)
        self.prop = _ParaProp(style)

    def write_to(self, writer: DocX):
        writer.write_paragraph(*self.datas,
                               prop=self.prop)


class _ParaProp:
    def __init__(self, style):
        self.style = style
        self.jc = None
        self.spacing = None
        self.snap_to_grid = None
        self.font = None
        self.size = None
        self.keep_next = False

    def set_jc(self, jc):
        self.jc = jc

    def set_spacing(self, before=None, after=None):
        self.spacing = _ParaSpacing(before, after)

    def set_snap_to_grid(self, val):
        self.snap_to_grid = val

    def set_font(self, font):
        self.font = font

    def set_size(self, size):
        self.size = size

    def set_style(self, style):
        self.style = style

    def set_keep_next(self, val):
        self.keep_next = val

    def write_to(self, writer: DocX):
        writer.write_para_prop(style=self.style,
                               jc=self.jc,
                               spacing=self.spacing,
                               snap_to_grid=self.snap_to_grid,
                               font=self.font,
                               size=self.size,
                               keep_next=self.keep_next)


class _ParaSpacing:
    def __init__(self, before, after):
        self.before = before
        self.after = after

    def write_to(self, writer):
        writer.write_para_prop_spacing(self.before, self.after)


class Run:
    def __init__(self, text, size=None, font=None, italic=False, bold=False):
        self.text = text
        self.size = size
        self.font = font
        self.italic = italic
        self.bold = bold

    def write_to(self, writer: DocX):
        writer.write_run(text=self.text,
                         size=self.size,
                         font=self.font,
                         italic=self.italic,
                         bold=self.bold)


class Table:
    def __init__(self, data_by_rows=None, *,
                 jc='center', style='a1', title=None):
        self.data_by_rows = self._process_data(data_by_rows)
        self.prop = _TableProp(jc=jc, style=style)

        if title is not None:
            mark = Bookmark('表')
            self.reference = mark.reference
            if isinstance(title, str):
                self.title = Paragraph(mark, title)
            elif isinstance(title, list):
                self.title = Paragraph(mark, *title)
            self.title.prop.set_jc('center')
        else:
            self.title = None

    @staticmethod
    def _process_data(data_by_rows):
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

    def write_to(self, writer: DocX):
        writer.write_table(data_by_rows=self.data_by_rows,
                           prop=self.prop,
                           title=self.title)


class _TableProp:
    def __init__(self, jc, style):
        self.jc = jc
        self.style = style

        self.layout = 'auto'
        self.pos_pr = None
        self.border = None
        self.grid_col = None
        self.cell_margin = None

    def set_pos_pr(self, x=None, y=None, x_spec=None, y_spec=None):
        self.pos_pr = _TablePosPr(x, y, x_spec, y_spec)
        self.jc = None

    def set_border(self, top=False, bottom=False,
                   left=False, right=False,
                   inside_v=False, inside_h=False,
                   size=4):
        self.border = _TableBorder(top, bottom, left, right, inside_v, inside_h, size)

    def set_grid_col(self, col_size_list: list):
        self.grid_col = _TableGridCol(col_size_list)
        self.layout = 'fixed'

    def set_cell_margin(self, top=0, bottom=0, left=0, right=0):
        self.cell_margin = _TableCellMargin(top, bottom, left, right)

    def write_to(self, writer: DocX):
        writer.write_table_prop(jc=self.jc,
                                style=self.style,
                                layout=self.layout,
                                pos_pr=self.pos_pr,
                                border=self.border,
                                grid_col=self.grid_col,
                                cell_margin=self.cell_margin)


class _TablePosPr:
    def __init__(self, x, y, x_spec, y_spec):
        self.x = x
        self.y = y
        self.x_spec = x_spec
        self.y_spec = y_spec

    def write_to(self, writer: DocX):
        writer.write_table_prop_pos_pr(self.x, self.y, self.x_spec, self.y_spec)


class _TableBorder:
    def __init__(self, top, bottom, left, right, inside_v, inside_h, size):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.inside_v = inside_v
        self.inside_h = inside_h
        self.size = size

    def write_to(self, writer: DocX):
        writer.write_table_prop_border(self.top, self.bottom, self.left, self.right,
                                       self.inside_v, self.inside_h,
                                       self.size)


class _TableGridCol:
    def __init__(self, col_size_list: list):
        self.col_size_list = col_size_list

    def write_to(self, writer: DocX):
        writer.write_table_prop_grid_col(self.col_size_list)


class _TableCellMargin:
    def __init__(self, top, bottom, left, right):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right

    def write_to(self, writer: DocX):
        writer.write_table_prop_cell_margin(self.top, self.bottom, self.left, self.right)


class Cell:
    def __init__(self, item, h_align='center', v_align='center'):
        if isinstance(item, Paragraph):
            self.para = item
        elif isinstance(item, tuple) or isinstance(item, list):
            self.para = Paragraph(*item)
        else:
            self.para = Paragraph(item)

        self.prop = _CellProp(h_align=h_align, v_align=v_align)

    def write_to(self, writer: DocX):
        writer.write_cell(para=self.para,
                          prop=self.prop)


class _CellProp:
    def __init__(self, h_align, v_align):
        self.jc = h_align
        self.v_align = v_align

        self.border = None

    def set_border(self, top=False, bottom=False, left=False, right=False, size=4):
        self.border = _CellBorder(top, bottom, left, right, size)

    def set_align(self, h_align=None, v_align=None):
        if h_align:
            self.jc = h_align
        if v_align:
            self.v_align = v_align

    def write_to(self, writer: DocX):
        writer.write_cell_prop(v_align=self.v_align,
                               border=self.border)


class _CellBorder:
    def __init__(self, top, bottom, left, right, size):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.size = size

    def write_to(self, writer: DocX):
        writer.write_cell_prop_border(self.top, self.bottom, self.left, self.right, self.size)


class Footnote:
    def __init__(self, *datas):
        self.datas = list()
        for data in datas:
            if isinstance(data, str):
                self.datas.append(Run(data))
            elif isinstance(data, int) or isinstance(data, float):
                self.datas.append(Run(f'{data}'))
            else:
                self.datas.append(data)

    def write_to(self, writer):
        writer.write_footnote(self.datas)


class Bookmark:
    id_ = 1

    def __init__(self, type_):
        self.type_ = type_
        self.mark_id = self.id_
        self.id_ += 1
        self.reference = Reference(self.mark_id)

    def write_to(self, writer: DocX):
        writer.write_bookmark(type_=self.type_,
                              mark_id=self.mark_id)


class Reference:
    def __init__(self, mark_id):
        self.mark_id = mark_id

    def write_to(self, writer: DocX):
        writer.write_reference(mark_id=self.mark_id)


class Figure:
    def __init__(self, file, scale=1):
        image = Image.open(file)

        size = image.size
        dpi = image.info['dpi']

        self.inch_size = [s / d for s, d in zip(size, dpi)]
        self.format = 'jpeg'

        b = BytesIO()
        image.save(b, self.format)
        self.data = b.getvalue()

        self.scale = scale


class InlineFigure(Figure):
    def write_to(self, writer: DocX):
        x, y = self.inch_size
        writer.write_inline_figure(data=self.data, fmt=self.format,
                                   cx=x * self.scale, cy=y * self.scale)


class StandaloneFigure(Figure):
    def __init__(self, file, scale=1, title=None):
        super().__init__(file=file, scale=scale)

        if title is not None:
            mark = Bookmark('图')
            self.reference = mark.reference
            if isinstance(title, str):
                self.title = Paragraph(mark, title)
            elif isinstance(title, list):
                self.title = Paragraph(mark, *title)
            self.title.prop.set_jc('center')
        else:
            self.title = None

    def write_to(self, writer: DocX):
        x, y = self.inch_size
        writer.write_standalone_figure(title=self.title,
                                       data=self.data, fmt=self.format,
                                       cx=x * self.scale, cy=y * self.scale)


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


P = Paragraph
MP = MathPara


# Inner class


def add_default_cover(doc: Reporter, project='**工程', name='**计算书',
                      part='**专业', phase='**阶段',
                      number='', secret='',
                      footer_str='湖南省水利水电勘测设计研究总院'):
    lt = Table([['编号', number]])
    lt.prop.set_grid_col([500, 2000])
    lt.prop.set_cell_margin()
    lt.prop.set_pos_pr(x_spec='left', y_spec='top')
    lt.prop.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    doc.add(lt)

    rt = Table([['秘密', secret]])
    rt.prop.set_grid_col([500, 2000])
    rt.prop.set_cell_margin()
    rt.prop.set_pos_pr(x_spec='right', y_spec='top')
    rt.prop.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    doc.add(rt)

    title = Table([[Run('计  算  书', size=72)]])
    title.prop.set_pos_pr(y=3000, x_spec='center')
    doc.add(title)

    cells = dict()
    for txt in [project, part, phase, name, '工程名称', '专业名称', '设计阶段', '计算书名称']:
        if txt in [project, part, phase, name]:
            c = Cell(Run(txt, font='楷体', size=30))
            c.prop.set_border(bottom=True, size=6)
            c.prop.set_align(h_align='center')
        else:
            c = Cell(Run(txt, size=30))
            c.prop.set_align(h_align='right')
        c.para.prop.set_snap_to_grid('false')
        c.para.prop.set_spacing(200, 0)
        c.prop.set_align(v_align='bottom')
        cells[txt] = c

    sub = Table([[cells['工程名称'], cells[project]],
                 [cells['专业名称'], cells[part]],
                 [cells['设计阶段'], cells[phase]],
                 [cells['计算书名称'], cells[name]]])
    sub.prop.set_grid_col([1600, 5000])
    sub.prop.set_cell_margin(right=10)
    sub.prop.set_pos_pr(x_spec='center', y=6000)
    doc.add(sub)

    cells = dict()
    for txt in ['审查', '校核', '计算', '年', '月', '日']:
        c = Cell(Run(txt, size=24))
        c.prop.set_align(v_align='bottom')
        c.para.prop.set_snap_to_grid('false')
        c.para.prop.set_spacing(120, 0)
        cells[txt] = c

    blk = Cell(Run(''))
    blk.prop.set_border(bottom=True, size=6)
    blk.para.prop.set_snap_to_grid('false')
    blk.para.prop.set_spacing(120, 0)
    blk.para.prop.set_font('楷体')
    blk.para.prop.set_size(24)

    sig = Table([[cells['审查'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']],
                 [cells['校核'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']],
                 [cells['计算'], blk, '', blk, cells['年'], blk, cells['月'], blk, cells['日']]])
    sig.prop.set_grid_col([700, 1500, 300, 900, 300, 600, 300, 600, 300])
    sig.prop.set_cell_margin()
    sig.prop.set_pos_pr(x_spec='center', y=10500)
    doc.add(sig)

    footer = Table([[Run(footer_str, size=24)]])
    footer.prop.set_pos_pr(x_spec='center', y=13500)
    doc.add(footer)
