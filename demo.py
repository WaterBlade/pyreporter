from pyreporter.docx import DocX

if __name__ == '__main__':
    d = DocX()
    d.save('hello_world.zip')