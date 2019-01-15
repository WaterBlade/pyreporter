from .docx import DocX
from PIL import Image
from io import BytesIO
from typing import Union, List
import weakref


class Reporter:
    def __init__(self):
        self.content = StandaloneComposite()
        self.writer = DocX()

    def set_writer(self, writer):
        self.writer = writer

    def save(self, path):
        self.content.visit(self.writer)
        self.writer.save(path)


class ElementBase:
    def visit(self, visitor):
        pass


class InlineElement(ElementBase):
    pass


class StandaloneElement(ElementBase):
    pass


def preprocess(item: Union[str, InlineElement]):
    if isinstance(item, str):
        return Text(str)
    else:
        return item


class InlineComposite(InlineElement):
    def __init__(self, *items):
        self._element_list = list()
        for item in items:
            self.add(item)

    def add(self, element):
        element = preprocess(element)
        assert isinstance(element, InlineElement)
        self._element_list.append(element)

    def visit(self, visitor):
        for element in self._element_list:
            element.visit(visitor)


class StandaloneComposite(StandaloneElement):
    def __init__(self, *items):
        self._element_list = list()
        for item in items:
            self.add(item)

    def add(self, element):
        assert isinstance(element, StandaloneElement)
        self._element_list.append(element)

    def visit(self, visitor):
        for element in self._element_list:
            element.visit(visitor)


class Text(InlineElement):
    def __init__(self, text, size='normal',
                 italic=False, bold=False, underline=False):
        self.text = text
        self.size = size
        self.italic = italic
        self.bold = bold
        self.underline = underline

    def visit(self, visitor):
        visitor.visit_text(text=self.text,
                           size=self.size,
                           italic=self.italic,
                           bold=self.bold,
                           underline=self.underline)


class InlineMath(InlineElement):
    def __init__(self, content):
        self.content = content

    def visit(self, visitor):
        visitor.visit_inline_math(content=self.content)


class InlineFigure(InlineElement):
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


class Bookmark(InlineElement):
    def __init__(self, type_):
        self.type_ = type_
        self.reference = Reference(self)

    def visit(self, visitor):
        visitor.visit_bookmark(self.type_, self.reference)


class Reference(InlineElement):
    def __init__(self, bookmark: Bookmark):
        self.bookmark = weakref.proxy(bookmark)

    def visit(self, visitor):
        visitor.visit_reference(self.bookmark)


class Footnote(InlineElement):
    def __init__(self, *items):
        self.content = InlineComposite(*items)

    def visit(self, visitor):
        visitor.visit_footnote(self.content)


class Math(StandaloneElement):
    def __init__(self, content):
        self.content = content

    def visit(self, visitor):
        visitor.visit_math(content=self.content)


class Figure(StandaloneElement):
    def __init__(self, file, height=None, title: Union[str, InlineElement]=None):
        self.figure = _FigureContent(file, height)
        self.title = preprocess(title)

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


class Heading(StandaloneElement):
    def __init__(self, content: Union[str, InlineElement], level=1):
        self.content = preprocess(content)
        self.level = level

    def visit(self, visitor):
        visitor.visit_heading(content=self.content,
                              level=self.level)


class Paragraph(ElementBase):
    def __init__(self, content: Union[str, InlineElement]):
        self.content = preprocess(content)

    def visit(self, visitor):
        visitor.visit_paragraph(content=self.content)


class Table(ElementBase):
    def __init__(self, content_by_rows: List[List[InlineElement]]):
        self.content_by_rows = content_by_rows

    def visit(self, visitor):
        visitor.visit_table(content_by_rows=self.content_by_rows)


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
