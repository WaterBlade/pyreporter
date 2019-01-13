from PIL import Image
from io import BytesIO
from .docx import DocX
from .expression import Expression, Variable
import weakref
from typing import Union


def index_gen(*, start):
    index = start
    while True:
        yield index
        index += 1


def make_title(*, type_, title):
    mark = Bookmark(type_)
    reference = mark.reference
    title = Paragraph(mark, title)
    title.prop.set_jc('center')
    return title, reference


class ReportElement:
    def set_root(self, root):
        self.root = weakref.proxy(root)

    def get_root(self):
        return self.root.get_root()

    def build_when_add_to_report(self):
        pass


class ReportComposite(ReportElement):
    def __init__(self, *datas):
        self.datas = list()
        for d in datas:
            self.add(d)

    def add(self, data):
        if isinstance(data, str):
            data = Run(data)
        data.set_root(self)
        self.datas.append(data)

    def build_when_add_to_report(self):
        for data in self.datas:
            data.build_when_add_to_report()

    def write_to(self, writer):
        writer.write_composite(self.datas)


class Report:
    def __init__(self):
        self.cover = None
        self.catalog = None

        self.heading_list = list()
        self.data_list = list()

        self.index_generator = index_gen(start=1)

        self.catalog_level = 3

    def set_default_cover(self):
        self.set_cover(make_default_cover())

    def set_cover(self, cover):
        self.cover = cover
        self.catalog = Catalog()
        self.catalog.set_root(self)

    def add(self, item: Union[ReportElement, str]):
        if isinstance(item, Run):
            item = Paragraph(item)
        elif isinstance(item, Math):
            item = Paragraph(item)
        elif isinstance(item, InlineFigure):
            item = Paragraph(item)
        elif isinstance(item, str):
            item = Paragraph(Run(item))

        self.data_list.append(item)
        item.set_root(self)
        item.build_when_add_to_report()

    def add_heading(self, head, level):
        self.add(Heading(head, level))

    def add_paragraph(self, *data):
        self.add(Paragraph(*data))

    def add_table(self, title, data_by_rows, cols=None):
        t = Table(data_by_rows, title=title)  # type: Table
        t.prop.set_border(top=True, bottom=True, left=True, right=True,
                          inside_h=True, inside_v=True)
        if cols:
            t.prop.set_grid_col(cols)
        self.add(t)
        return t.reference

    def add_figure(self, title, figure, height=None, width=None):
        fig = StandaloneFigure(figure, height=height, width=width, title=title)
        self.add(fig)
        return fig.reference

    def add_math_definition(self):
        pass

    def add_math_procedure(self):
        pass

    def add_math_symbol_note(self):
        pass

    def add_named_math(self, *math_eles):
        m = ReferencedMathPara(Math(*math_eles))
        self.add(m)
        return m.reference

    def add_mul_named_math(self, *math_eles_list):
        m = ReferencedMathPara(Math(MathMultiLine(math_eles_list)))
        self.add(m)
        return m.reference

    def add_unnamed_math(self, *math_eles):
        self.add(MathPara(Math(*math_eles)))

    def add_mul_unnamed_math(self, *math_eles_list):
        self.add(Math(MathMultiLine(math_eles_list)))

    def get_root(self):
        return self

    def write_to(self, writer):
        self.catalog.add_heading(self.heading_list)
        if self.cover is not None:
            self.data_list = [self.cover, self.catalog] + self.data_list
        writer.write_docx(self)


class PageBreak(ReportComposite):
    def write_to(self, writer):
        writer.write_page_break()


class Catalog(ReportElement):
    def __init__(self):
        self.heads = list()

    def is_empty(self):
        return len(self.heads) == 0

    def add_heading(self, heads):

        code = [0]
        level = 1

        for h in heads:
            if h.level == level:
                code[-1] += 1
            elif h.level == level + 1:
                code.append(1)
                level += 1
            elif h.level < level:
                while h.level < level:
                    code.pop()
                    level -= 1
            else:
                raise RuntimeError
            self.heads.append(_CatalogHeading(h.head,
                                              h.level,
                                              '.'.join(str(c) for c in code),
                                              h.mark.mark_id))

    def write_to(self, writer: DocX):
        catalog_level = self.get_root().catalog_level
        writer.write_catalog(heads=self.heads, catalog_level=catalog_level)


class _CatalogHeading:
    def __init__(self, head, level, serial_code, mark_id):
        self.head = head
        self.level = level
        self.serial_code = serial_code
        self.mark_id = mark_id


class Heading(ReportElement):
    def __init__(self, head: str, level=1):
        self.head = head
        self.level = level

        self.mark = None
        self.para = None
        self.need_page_break = False

    def build_when_add_to_report(self):
        root = self.get_root()
        if self.level == 1 and len(root.heading_list) > 0:
            self.need_page_break = True
        if self.level <= root.catalog_level:
            self.mark = HeadingMark(self.head)
            self.para = Paragraph(self.mark)
        else:
            self.para = Paragraph(self.head)

        self.para.set_root(self)
        self.para.build_when_add_to_report()

        root.heading_list.append(self)

    def write_to(self, writer: DocX):
        writer.write_heading(para=self.para,
                             level=self.level,
                             need_page_break=self.need_page_break)


class Paragraph(ReportElement):
    def __init__(self, *datas, style=None):
        self.datas = list()
        self.add(*datas)
        self.prop = _ParaProp(style)

        self.root = None

    def build_when_add_to_report(self):
        for data in self.datas:
            data.build_when_add_to_report()

    def add(self, *datas):
        for data in datas:
            if isinstance(data, str):
                data = Run(data)
            elif isinstance(data, int) or isinstance(data, float):
                data = Run(f'{data}')

            self.datas.append(data)
            data.set_root(self)

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


class Run(ReportElement):
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


class Table(ReportElement):
    def __init__(self, data_by_rows=None, *,
                 jc='center', style=None, title=None):
        self.data_by_rows = self._process_data(data_by_rows)
        self.prop = _TableProp(jc=jc, style=style)

        self.title = None
        self.reference = None

        self.set_title(title)

    def build_when_add_to_report(self):
        for row in self.data_by_rows:
            for col in row:
                col.build_when_add_to_report()
        if self.title is not None:
            self.title.build_when_add_to_report()

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

    def set_title(self, title):
        if title is not None:
            title, reference = make_title(type_='表', title=title)
            self.reference = reference
            title.set_root(self)
        else:
            self.reference = None
        self.title = title

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

    def set_jc(self, jc):
        self.jc = jc

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


class Cell(ReportElement):
    def __init__(self, item, h_align='center', v_align='center'):
        if isinstance(item, Paragraph):
            self.para = item
        elif isinstance(item, tuple) or isinstance(item, list):
            self.para = Paragraph(*item)
        else:
            self.para = Paragraph(item)

        self.prop = _CellProp(h_align=h_align, v_align=v_align)

    def build_when_add_to_report(self):
        self.para.build_when_add_to_report()

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


class Footnote(ReportElement):
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


class Bookmark(ReportElement):
    def __init__(self, type_):
        self.type_ = type_
        self.mark_id = 0
        self.reference = Reference(self)

    def build_when_add_to_report(self):
        self.mark_id = next(self.get_root().index_generator)

    def write_to(self, writer: DocX):
        writer.write_bookmark(type_=self.type_,
                              mark_id=self.mark_id)


class HeadingMark(ReportElement):
    def __init__(self, head: str):
        self.head = head
        self.mark_id = 0

    def build_when_add_to_report(self):
        self.mark_id = next(self.get_root().index_generator)

    def write_to(self, writer: DocX):
        writer.write_heading_mark(head=self.head,
                                  mark_id=self.mark_id)


class Reference(ReportElement):
    def __init__(self, mark):
        self.mark = mark

    def write_to(self, writer: DocX):
        writer.write_reference(mark_id=self.mark.mark_id)


class Figure(ReportElement):
    def __init__(self, file, height=None, width=None):
        # height and width assume to be mm
        image = Image.open(file)

        size = image.size
        dpi = image.info['dpi']

        self.inch_size = [s / d for s, d in zip(size, dpi)]
        self.format = 'png'

        b = BytesIO()
        image.save(b, self.format)
        self.data = b.getvalue()

        self.height = height / 25.4 if height else height
        self.width = width / 25.4 if width else width

    def get_inch_size(self):
        x, y = self.inch_size
        if self.height is not None and self.width is None:
            x = x / y * self.height
            y = self.height
        elif self.height is None and self.width is not None:
            y = y / x * self.width
            x = self.width
        elif self.height is not None and self.width is not None:
            x, y = self.width, self.height
        return x, y


class InlineFigure(Figure):
    def write_to(self, writer: DocX):
        x, y = self.get_inch_size()

        writer.write_inline_figure(data=self.data, fmt=self.format,
                                   cx=x, cy=y)


class StandaloneFigure(Figure):
    def __init__(self, file, width=None, height=None, title=None):
        super().__init__(file=file, height=height, width=width)

        self.title = None
        self.reference = None

        self.set_title(title)

    def set_title(self, title):
        if title is not None:
            title, reference = make_title(type_='图', title=title)
            self.reference = reference
            title.set_root(self)
        else:
            self.reference = None
        self.title = title

    def write_to(self, writer: DocX):
        x, y = self.get_inch_size()

        writer.write_standalone_figure(title=self.title,
                                       data=self.data, fmt=self.format,
                                       cx=x, cy=y)


class FloatFigure(Figure):
    def write_to(self, writer: DocX):
        x, y = self.get_inch_size()

        writer.write_float_figure(data=self.data, fmt=self.format,
                                  cx=x, cy=y)


class MathComposite(ReportElement):
    def __init__(self, *items):
        self.datas = list()
        for item in items:
            self.add(item)

    def add(self, item):
        if isinstance(item, str):
            item = MathRun(item)
        self.datas.append(item)

    def write_to(self, writer):
        writer.write_math_composite(self.datas)


class MathPara(ReportElement):
    def __init__(self, math_obj):
        self.math = math_obj

    def write_to(self, writer):
        writer.write_mathpara(self.math)


class ReferencedMathPara(ReportElement):
    def __init__(self, math_obj):
        mark = Bookmark('式')
        self.reference = mark.reference
        c_math = Cell(MathPara(math_obj))
        c_mark = Cell(mark)

        table = Table([[c_math, c_mark]])
        table.set_root(self)

        table.prop.set_jc('center')
        table.prop.set_grid_col([8000, 1000])
        c_math.prop.set_align(h_align='center', v_align='center')
        c_mark.prop.set_align(h_align='center', v_align='center')

        self.table = table

    def write_to(self, writer):
        self.table.write_to(writer)


class Math(ReportElement):
    def __init__(self, *eles):
        self.datas = list()
        for ele in eles:
            if isinstance(ele, MathRun) or isinstance(ele, Expression):
                self.datas.append(ele)
            elif isinstance(ele, str):
                self.datas.append(MathRun(ele))
            else:
                raise TypeError('Unknown math type')

    def write_to(self, writer):
        writer.write_math(self.datas)


class MathMultiLine(ReportElement):
    def __init__(self, eqs):
        self.elements = list()
        for eq in eqs:
            self.elements.append(eq)

    def write_to(self, writer):
        writer.write_math_multi_line(self.elements)


class MathMultiLineBrace(ReportElement):
    def __init__(self, exp):
        self.exp = exp

    def write_to(self, writer: DocX):
        writer.write_multi_line_brace(self.exp)


class MathRun(ReportElement):
    def __init__(self, data, sty=None):
        self.data = f'{data}'
        self.sty = sty

    def write_to(self, writer: DocX):
        writer.write_math_run(self.data, self.sty)


class Formula:
    def __init__(self, var, expression):
        self.variable = var  # type: Variable
        self.expression = expression  # type: Expression

    def calc(self):
        self.variable.value = self.expression.calc()
        return self.variable.value

    def get_definition(self):
        return MathComposite(self.variable, MathRun('&=', sty='p'), self.expression)

    def get_procedure(self):
        result = self.expression.copy_result()
        if self.variable.unit is not None:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result(),
                                 self.variable.unit)
        else:
            return MathComposite(self.variable,
                                 MathRun('&=', sty='p'), result,
                                 MathRun('=', sty='p'), self.variable.copy_result())


class Equation:
    def __init__(self, left_exp, right_exp):
        self.left = left_exp
        self.right = right_exp

    def calc(self):
        return self.left.calc() - self.right.calc()

    def get_definition(self):
        return MathComposite(self.left, MathRun('&=', sty='p'), self.right)


P = Paragraph
MP = MathPara


def make_default_cover(project='**工程', name='**计算书',
                       part='**专业', phase='**阶段',
                       number='', secret='',
                       footer_str='湖南省水利水电勘测设计研究总院'):
    comp = ReportComposite()
    lt = Table([['编号', number]])
    lt.prop.set_grid_col([500, 2000])
    lt.prop.set_cell_margin()
    lt.prop.set_pos_pr(x_spec='left', y_spec='top')
    lt.prop.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    comp.add(lt)

    rt = Table([['秘密', secret]])
    rt.prop.set_grid_col([500, 2000])
    rt.prop.set_cell_margin()
    rt.prop.set_pos_pr(x_spec='right', y_spec='top')
    rt.prop.set_border(top=True, bottom=True, left=True, right=True, inside_v=True)
    comp.add(rt)

    title = Table([[Run('计  算  书', size=72)]])
    title.prop.set_pos_pr(y=3000, x_spec='center')
    comp.add(title)

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
    comp.add(sub)

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
    comp.add(sig)

    footer = Table([[Run(footer_str, size=24)]])
    footer.prop.set_pos_pr(x_spec='center', y=13500)
    comp.add(footer)

    comp.add(PageBreak())

    return comp
