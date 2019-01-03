class Reporter:
    def __init__(self):
        self.data_list = list()

    def add_paragraph(self, *item):
        self.data_list.append(Paragraph(*item))

    def add_mathpara(self, *item):
        self.data_list.append(MathPara(*item))

    def write_to(self, writer):
        writer.write(self.data_list)


class Paragraph:
    def __init__(self, *data):
        self.data = data

    def write_to(self, writer):
        writer.write_paragraph(*self.data)


class Run:
    def __init__(self, data):
        self.data = data

    def write_to(self, writer):
        writer.write_run(self.data)


class MathPara:
    def __init__(self, *data):
        self.data = data

    def write_to(self, writer):
        writer.write_mathpara(*self.data)


class Math:
    def __init__(self, *data):
        self.data = data

    def write_to(self, writer):
        writer.write_math(*self.data)


class Table:
    pass


class Figure:
    pass
