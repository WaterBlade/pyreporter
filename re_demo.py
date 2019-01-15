from pyreporter.re_reporter import Report


if __name__ == '__main__':
    doc = Report()
    doc.add_paragraph('hello world')
    doc.save('re_hello.docx')