from pyreporter.docx import DocX
from pyreporter.reporter import Reporter
from pyreporter.calc import V, Radical

if __name__ == '__main__':
    rep = Reporter()
    rep.add_paragraph('hello world')
    rep.add_mathpara(V('a')+V('b')/V('c')+5+Radical(12, 3)+V('d')**2)
    rep.write_to(DocX('hello_world.docx'))
