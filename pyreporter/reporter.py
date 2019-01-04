import os


class Reporter:
    def __init__(self):
        self.data_list = list()
        self.figure_list = list()

    def add_paragraph(self, *item):
        self.data_list.append(Paragraph(*item))

    def add_mathpara(self, *item):
        self.data_list.append(MathPara(*item))

    def add_table(self, data_by_rows):
        self.data_list.append(Table(data_by_rows))

    def write_to(self, writer):
        writer.write_docx(self)


class Paragraph:
    def __init__(self, *datas):
        self.datas = datas

    def write_to(self, writer):
        return writer.write_paragraph(*self.datas)


class Run:
    def __init__(self, data):
        self.data = data

    def write_to(self, writer):
        return writer.write_run(self.data)


class MathPara:
    def __init__(self, *datas):
        self.datas = datas

    def write_to(self, writer):
        return writer.write_mathpara(self.datas)


class Math:
    def __init__(self, *datas):
        self.datas = datas

    def write_to(self, writer):
        return writer.write_math(self.datas)


class Table:
    def __init__(self, data_by_rows):
        self.data_by_rows = data_by_rows

    def write_to(self, writer):
        return writer.write_table(self.data_by_rows)


class Figure:
    def __init__(self):
        self.data = None
        self.ext = '.png'

    def read(self, path_name):
        self.ext = os.path.splitext(path_name)[1]
        with open(path_name, 'rb') as f:
            self.data = f.read()


class InlineFigure(Figure):
    def write_to(self, writer):
        return writer.write_inline_figure(self.data, self.ext)

