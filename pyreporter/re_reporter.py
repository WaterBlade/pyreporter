from .re_docx import DocX
from PIL import Image
from io import BytesIO
from typing import Union, List
import weakref


class Report:
    def __init__(self):
        self.cover = None
        self.header = _Header()
        self.block = Block()
        self.writer = DocX()

    def add(self, item):
        self.block.add(item)

    def set_writer(self, writer):
        self.writer = writer

    def set_cover(self, cover):
        self.cover = cover

    def set_header(self, header):
        self.header = _Header(header)

    def save(self, path):
        if self.cover is not None:
            self.cover.visit(self.writer)
        self.header.visit(self.writer)
        self.block.visit(self.writer)
        self.writer.save(path)

    def add_paragraph(self, *items):
        self.add(Paragraph(*items))

    def add_math(self, math_content):
        self.add(Math(math_content))

    def add_figure(self, file, height=None, title=None):
        self.add(Figure(file, height=height, title=title))

    def add_table(self, content_by_rows, title=None):
        self.add(Table(content_by_rows, title))


class ReportElement:
    def visit(self, visitor):
        pass


class Block:
    def __init__(self, *items):
        self._element_list = list()
        for item in items:
            self.add(item)

    def add(self, element):
        assert isinstance(element, Context)
        if isinstance(element, Block):
            self._element_list.extend(element._element_list)
        else:
            self._element_list.append(element)

    def add_paragraph(self, *items):
        self.add(Paragraph(*items))

    def add_math(self, math_content):
        self.add(Math(math_content))

    def add_figure(self, file, height=None, title=None):
        self.add(Figure(file, height=height, title=title))

    def add_table(self, content_by_rows, title=None):
        self.add(Table(content_by_rows, title))

    def visit(self, visitor):
        for element in self._element_list:
            element.visit(visitor)


class ContentElement:
    def visit(self, visitor):
        pass


class Context:
    def visit(self, visitor):
        pass


class Content(ContentElement):
    def __init__(self, *items):
        self._element_list = list()
        for item in items:
            self.add(item)

    def add(self, element):
        if isinstance(element, str):
            element = Text(element)

        assert isinstance(element, ContentElement)

        if isinstance(element, Content):
            self._element_list.extend(element._element_list)
        else:
            self._element_list.append(element)

    def visit(self, visitor):
        for element in self._element_list:
            element.visit(visitor)


class Text(ContentElement):
    def __init__(self, text, font=None, size=None,
                 italic=False, bold=False, underline=False):
        self.text = text
        self.font = font
        self.size = size
        self.italic = italic
        self.bold = bold
        self.underline = underline

    def visit(self, visitor):
        visitor.visit_text(text=self.text,
                           font=self.font,
                           size=self.size,
                           italic=self.italic,
                           bold=self.bold,
                           underline=self.underline)


class InlineMath(ContentElement):
    def __init__(self, content):
        self.content = content

    def visit(self, visitor):
        visitor.visit_inline_math(content=self.content)


class InlineFigure(ContentElement):
    def __init__(self, file, height=None):
        self.figure = _FigureContent(file, height)

    def visit(self, visitor):
        figure = self.figure.figure
        format_ = self.figure.format_
        width = self.figure.width
        height = self.figure.height
        visitor.visit_inline_figure(figure=figure,
                                    format_=format_,
                                    width=width,
                                    height=height)


class Bookmark(ContentElement):
    def __init__(self, type_):
        self.type_ = type_
        self.reference = Reference(self)

    def visit(self, visitor):
        visitor.visit_bookmark(self.type_, self)


class Reference(ContentElement):
    def __init__(self, bookmark: Bookmark):
        self.bookmark = bookmark

    def visit(self, visitor):
        visitor.visit_reference(self.bookmark)


class Footnote(ContentElement):
    def __init__(self, *items):
        self.content = Content(*items)

    def visit(self, visitor):
        visitor.visit_footnote(self.content)


class Heading(Context):
    def __init__(self, heading: str, level=1):
        self.content = Content(heading)
        self.level = level

    def visit(self, visitor):
        visitor.visit_heading(content=self.content,
                              level=self.level,
                              heading=self)


class Paragraph(Context):
    def __init__(self, *content: Union[str, ContentElement]):
        self.content = Content(*content)

    def visit(self, visitor):
        visitor.visit_paragraph(content=self.content)


class Math(Context):
    def __init__(self, content):
        self.content = content

    def visit(self, visitor):
        visitor.visit_math(content=self.content)


class Figure(Context):
    def __init__(self, file, height=None, title: Union[str, ContentElement]=None):
        self.figure = _FigureContent(file, height)
        self.title = Content(title)

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
    def __init__(self, content_by_rows: List[List[ContentElement]],
                 title: Union[str, ContentElement]=None):
        self.content_by_rows = list()
        for r in content_by_rows:
            row = list()
            for c in r:
                row.append(Content(c))
            self.content_by_rows.append(row)
        self.title = Content(title)

    def visit(self, visitor):
        visitor.visit_table(content_by_rows=self.content_by_rows,
                            title=self.title)


class _Cover(ReportElement):
    def visit(self, visitor):
        visitor.visit_cover(self)


class _Header(ReportElement):
    def __init__(self, header: str='计算书'):
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
