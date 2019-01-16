from pyreporter.re_reporter import Report, DefaultCover, Footnote, InlineFigure
from pyreporter import expression as exp

V = exp.Variable


if __name__ == '__main__':
    doc = Report()
    doc.set_cover(DefaultCover())
    doc.set_header('湖南省水利水电勘测设计研究总院')
    doc.add_paragraph('hello world')
    doc.add_heading('Hello World', 1)
    doc.add_math(V('a')+V('b'))
    ref1 = doc.add_definition(V('a')+V('b'))
    ref = doc.add_table([['1', '2', '3'],
                   ['4', '5', '6']], title='123')
    doc.add_paragraph('abcdefg', Footnote('hijklmn'), '见', ref, InlineFigure('桌面.PNG', height=1))
    doc.save('re_hello.docx')