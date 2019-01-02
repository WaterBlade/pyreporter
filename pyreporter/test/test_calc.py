from ..calc import Variable, Constant, Unit

V = Variable
C = Constant


def test_add():
    v1 = V('v1', value=2)
    v2 = V('v2', value=3)
    assert 5 == (v1+v2).calc()


def test_complex_expression():
    v1 = V('v1', value=2)
    v2 = V('v2', value=3)
    v3 = V('v3', value=4)
    exp = (1+v1*v2+v3)/1
    assert exp.calc() == (1+2*3+4)/1


def test_unit():
    u1 = Unit.m
    u2 = Unit.mm
    assert u1.convert_factor(u2) == 1000
    assert u1.base == u2.base


def test_unit_m2():
    u = Unit.m**2
    assert u.symbol.__class__.__name__ == 'Pow'
    assert u.symbol.left.symbol == 'm'


def test_unit_pa():
    u = Unit.Pa
    u2 = Unit.kPa
    assert u.base == (1, -1, -2)
    assert u2.base == (1, -1, -2)
    assert u2.convert_factor(u) == 1000

