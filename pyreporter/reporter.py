from .docx import DocX
from PIL import Image
from io import BytesIO
from typing import Union, List
from collections import OrderedDict


__all__ = ['Report', 'DefaultCover',
           'Block', 'StandaloneFigure', 'StandaloneMath',
           'Definition', 'Procedure', 'Note', 'VariableValue',
           'Figure', 'Math']


class Report:
    def __init__(self):
        self.cover = None
        self.header = None
        self.block = Block()
        self.writer = DocX()

        self.symbol_set = set()

    def add(self, item):
        self.block.add(item)
        if isinstance(item, Note):
            item.remove_duplicate_note(self.symbol_set)
        elif isinstance(item, Block):
            item.remove_duplicate_note(self.symbol_set)

    def set_writer(self, writer):
        self.writer = writer

    def set_cover(self, cover):
        self.cover = cover

    def set_header(self, header: str):
        self.header = _Header(header)

    def save(self, path):
        if self.cover is not None:
            self.writer.set_cover(self.cover)
        if self.header is not None:
            self.header.visit(self.writer)
        self.block.visit(self.writer)
        self.writer.save(path)

    def add_heading(self, heading: str, level: int):
        self.add(Heading(heading, level))

    def add_paragraph(self, *items):
        self.add(Paragraph(*items))

    def add_math(self, math_content):
        self.add(StandaloneMath(math_content))

    def add_figure(self, file, height=None, title=None):
        fig = StandaloneFigure(file, height=height, title=title)
        self.add(fig)
        return fig.reference

    def add_table(self, content_by_rows, title=None):
        tab = Table(content_by_rows, title)
        self.add(tab)
        return tab.reference


class ReportElement:
    def visit(self, visitor):
        pass


class Block:
    def __init__(self, *items):
        self._element_list = list()
        for item in items:
            self.add(item)

    def remove_duplicate_note(self, symbol_set):
        for item in self._element_list:
            if isinstance(item, Note):
                item.remove_duplicate_note(symbol_set)

    def add(self, element):
        assert isinstance(element, Context)
        if isinstance(element, Block):
            self._element_list.extend(element._element_list)
        else:
            self._element_list.append(element)

    def add_heading(self, heading: str, level: int):
        self.add(Heading(heading, level))

    def add_paragraph(self, *items):
        self.add(Paragraph(*items))

    def add_math(self, math_content):
        self.add(StandaloneMath(math_content))

    def add_figure(self, file, height=None, title=None):
        fig = StandaloneFigure(file, height=height, title=title)
        self.add(fig)
        return fig.reference

    def add_table(self, content_by_rows, title=None):
        tab = Table(content_by_rows, title)
        self.add(tab)
        return tab.reference

    def visit(self, visitor):
        for element in self._element_list:
            element.visit(visitor)


class Content:
    def visit(self, visitor):
        pass


class Context:
    def visit(self, visitor):
        pass


class Composite:
    def __init__(self, *items):
        self.list_ = list()
        for item in items:
            self.append(item)

    def append(self, item):
        if isinstance(item, Composite):
            self.list_.extend(item.list_)
        else:
            self.list_.append(item)

    def extend(self, list_):
        self.list_.extend(list_)

    def __getitem__(self, index):
        return self.list_[index]

    def __setitem__(self, key, value):
        self.list_[key] = value

    def __len__(self):
        return len(self.list_)

    def __iter__(self):
        return iter(self.list_)

    def visit(self, visitor):
        return visitor.visit_composite(self.list_)


class Text(Content):
    def __init__(self, text, font=None, size=None,
                 italic=False, bold=False, underline=False):
        self.text = text
        self.font = font
        self.size = size
        self.italic = italic
        self.bold = bold
        self.underline = underline

    def visit(self, visitor):
        return visitor.visit_text(text=self.text,
                                  font=self.font,
                                  size=self.size,
                                  italic=self.italic,
                                  bold=self.bold,
                                  underline=self.underline)


class Math(Content):
    def __init__(self, *items):
        content = Composite(*items)
        for item in content:
            assert not isinstance(item, str)
            assert not isinstance(item, Context)
            assert not isinstance(item, Content)
        self.content = content

    def append(self, item):
        self.content.append(item)

    def visit(self, visitor):
        return visitor.visit_inline_math(content=self.content)


class VariableValue(Math):
    def __init__(self, variable):
        if variable.unit is None:
            super().__init__(variable, _MathText('='), variable.copy_result())
        else:
            super().__init__(variable, _MathText('='), variable.copy_result(), variable.unit)


class Figure(Content):
    def __init__(self, file, height=None):
        self.figure = _FigureContent(file, height)

    def visit(self, visitor):
        figure = self.figure.figure
        format_ = self.figure.format_
        width = self.figure.width
        height = self.figure.height
        return visitor.visit_inline_figure(figure=figure,
                                           format_=format_,
                                           width=width,
                                           height=height)


class Bookmark(Content):
    def __init__(self, type_, left=None, right=None):
        self.type_ = type_
        self.reference = Reference(self)
        self.left = left
        self.right = right

    def visit(self, visitor):
        return visitor.visit_bookmark(type_=self.type_, bookmark=self,
                                      left=self.left, right=self.right)


class Reference(Content):
    def __init__(self, bookmark: Bookmark):
        self.bookmark = bookmark

    def visit(self, visitor):
        return visitor.visit_reference(self.bookmark)


class Footnote(Content):
    def __init__(self, *items):
        content = Composite(*items)
        for item in content:
            assert isinstance(item, Content)
            assert not isinstance(item, Footnote)
        self.content = content

    def visit(self, visitor):
        return visitor.visit_footnote(self.content)


class Heading(Context):
    def __init__(self, heading: str, level=1):
        self.content = heading
        self.level = level

    def visit(self, visitor):
        return visitor.visit_heading(content=self.content,
                                     level=self.level,
                                     heading=self)


class Paragraph(Context):
    def __init__(self, *items):
        content = Composite(*items)
        for i in range(len(content)):
            if isinstance(content[i], str):
                content[i] = Text(content[i])
            assert isinstance(content[i], Content)
        self.content = content

    def visit(self, visitor):
        visitor.visit_paragraph(content=self.content)


class StandaloneMath(Context):
    def __init__(self, *items):
        content = Composite(*items)
        for item in content:
            assert isinstance(item, Math)
        self.content = content

    def visit(self, visitor):
        visitor.visit_standalone_math(content=self.content)


class Definition(Context):
    def __init__(self, formula_or_calculator):
        self.content = Composite()
        formula_or_calculator.visit(self)

        self.bookmark = Bookmark('式', left='(', right=')')
        self.reference = self.bookmark.reference

    def visit(self, visitor):
        visitor.visit_math_definition(content=self.content, bookmark=self.bookmark)

    def visit_formula(self, variable, expression, long):
        self.content.append(Math(variable, _MathText('=', align=True), expression))

    def visit_piecewise_formula(self, variable, expression, expression_list, condition_list, long):
        multi = _MultiLine(included='left')

        for exp, cond in zip(expression_list, condition_list):
            multi.add(Composite(_MathText('&'), exp, _MathText('&,'), cond))

        self.content.append(Math(variable, _MathText('=', align=True), multi))

    def visit_condition_formula(self, variable, expression, condition: bool, long: bool):
        self.visit_formula(variable, expression, long)

    def visit_calculator(self, formula_list):
        for formula in formula_list:
            formula.visit(self)


class Procedure(Context):
    def __init__(self, formula_or_calculator):
        self.content = Composite()
        formula_or_calculator.visit(self)

    def visit(self, visitor):
        visitor.visit_math_procedure(content=self.content)

    def visit_formula(self, variable, expression, long: bool):
        ret = list()

        if long:
            ret.append(Math(variable, _MathText('=', align=True), expression))
            ret.append(Math(_MathText('=', align=True), expression.copy_result()))
            ret.append(Math(_MathText('=', align=True), variable.copy_result()))
        else:
            ret.append(Math(variable,
                            _MathText('=', align=True), expression,
                            _MathText('='), expression.copy_result(),
                            _MathText('='), variable.copy_result()))

        if variable.unit is not None:
            ret[-1].append(variable.unit)

        self.content.extend(ret)

    def visit_piecewise_formula(self, variable, expression, expression_list, condition_list, long: bool):
        self.visit_formula(variable, expression, long)

    def visit_condition_formula(self, variable, expression, condition: bool, long: bool):
        if condition:
            self.visit_formula(variable, expression, long)

    def visit_calculator(self, formula_list):
        for formula in formula_list[::-1]:
            formula.visit(self)


class Note(Context):
    def __init__(self, formula_or_calculator):
        self._list = list()
        self.variable_dict = OrderedDict()
        formula_or_calculator.visit(self)

    def remove_duplicate_note(self, var_set):
        ret = list()
        for var in self.variable_dict.values():
            if var not in var_set:
                ret.append(var)
                var_set.add(var)
        self._list = ret

    def visit(self, visitor):
        visitor.visit_math_note(self._list)

    def visit_formula(self, variable, expression, long):
        self.variable_dict.update(variable.get_variable_dict())
        self.variable_dict.update(expression.get_variable_dict())

    def visit_piecewise_formula(self, variable, expression, expression_list, condition_list, long):
        self.visit_formula(variable, expression, long)

    def visit_condition_formula(self, variable, expression, condition: bool, long: bool):
        self.visit_formula(variable, expression, long)

    def visit_calculator(self, formula_list):
        for formula in formula_list:
            formula.visit(self)


class StandaloneFigure(Context):
    def __init__(self, file, height=None, title: Union[str, Content] = None):
        self.figure = _FigureContent(file, height)

        self.reference = None
        if title is not None:
            mark = Bookmark('图')
            self.reference = mark.reference
            title = Composite(mark, title)
            for item in title:
                assert isinstance(item, Content)
        self.title = title

    def visit(self, visitor):
        figure = self.figure.figure
        format_ = self.figure.format_
        width = self.figure.width
        height = self.figure.height
        title = self.title
        visitor.visit_inline_figure(figure=figure,
                                    format_=format_,
                                    width=width,
                                    height=height,
                                    title=title)


class Table(Context):
    def __init__(self, content_by_rows: List[List[Content]],
                 title: Union[str, Content] = None):
        self.content_by_rows = list()
        for r in content_by_rows:
            row = list()
            for c in r:
                if isinstance(c, str):
                    c = Text(c)
                elif isinstance(c, float) or isinstance(c, int):
                    c = Text(f'{c}')
                assert isinstance(c, Content)
                row.append(c)
            self.content_by_rows.append(row)

        self.reference = None
        if title is not None:
            mark = Bookmark('表')
            self.reference = mark.reference
            title = Composite(mark, title)
            for item in title:
                assert isinstance(item, Content)
        self.title = title

    def visit(self, visitor):
        visitor.visit_table(content_by_rows=self.content_by_rows,
                            title=self.title)


class DefaultCover(ReportElement):
    def __init__(self, project='**工程', name='**计算书',
                 part='**专业', phase='**阶段',
                 number='', secret='',
                 footer_str='湖南省水利水电勘测设计研究总院'):
        self.type_ = 'default cover'
        self.project = project
        self.name = name
        self.part = part
        self.phase = phase
        self.number = number
        self.secret = secret
        self.footer_str = footer_str

    def visit(self, visitor):
        visitor.visit_cover(project=self.project, name=self.name,
                            part=self.part, phase=self.phase,
                            number=self.number, secret=self.secret,
                            footer_str=self.footer_str)


class _Header(ReportElement):
    def __init__(self, header: str):
        self.header = header

    def visit(self, visitor):
        visitor.visit_header(self.header)


class _FigureContent:
    def __init__(self, file, height=None):
        image = Image.open(file)

        size = image.size
        dpi = image.info['dpi']

        self.width, self.height = [s / d for s, d in zip(size, dpi)]
        self.format_ = 'png'

        b = BytesIO()
        image.save(b, self.format_)
        self.figure = b.getvalue()

        if height is not None:
            height /= 25.4
            self.width *= (height / self.height)
            self.height = height


class _MathText:
    def __init__(self, text, align=False):
        self.text = text
        self.align = align

    def visit(self, visitor):
        return visitor.visit_math_text(self.text, align=self.align)


class _MultiLine:
    def __init__(self, *exps, included: str = None):
        self._list = list()
        for exp in exps:
            self.add(exp)
        self.included = included

    def add(self, item):
        if isinstance(item, Composite):
            self._list.append(item)
        else:
            raise TypeError('Unknown math type % s' % type(item))

    def visit(self, visitor):
        return visitor.visit_multi_line(self._list, self.included)




