from pyreporter.docbuilder import ReportBuilder


if __name__ == '__main__':
    import cProfile, pstats, io
    pr = cProfile.Profile()
    pr.enable()
    rep = ReportBuilder()
    rep.heading(1, 'hello')
    rep.paragraph('hello')
    rep.heading(2, 'hello again')
    rep.paragraph('hello again')
    rep.heading(3, 'hello again again')
    rep.paragraph('hello again again')
    rep.save('builder_demo.docx')
    pr.disable()
    s = io.StringIO()
    sortby = 'cumulative'
    ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
    ps.print_stats()
    print(s.getvalue())

