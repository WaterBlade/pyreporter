from pyreporter.docx import DocX
from pyreporter.reporter import Reporter
from pyreporter.calc import V, Radical

if __name__ == '__main__':
    rep = Reporter()
    rep.add_paragraph('hello world')
    rep.add_mathpara(V('a')+V('b')+Radical(V('c')+2, 2))
    rep.write_to(DocX('hello_world.docx'))
