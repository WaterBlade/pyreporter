from pyreporter.docbuilder import ReportBuilder


if __name__ == '__main__':
    rep = ReportBuilder()
    rep.save('builder_demo.docx')