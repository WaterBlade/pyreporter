from ..expression import Variable, Number

V = Variable
C = Number


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


def test_variable_list():
    v1 = V('v1')
    v2 = V('v2')
    v3 = V('v3')
    exp = v1 + v2 + v3
    assert exp.get_variable_dict() == [v1, v2, v3]