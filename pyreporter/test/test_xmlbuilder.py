from ..xmlbuilder import NameSpace, NameSpaceWrapper, Element, XML


def test_namespace():
    ns = NameSpace('http://schemas.openxmlformats.org/markup-compatibility/2006',
                   'mc')
    attr = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    assert str(ns) == attr


def test_specificname():
    sn = NameSpaceWrapper(NameSpace('http://schemas.openxmlformats.org/markup-compatibility/2006',
                                     'mc'),
                           '100')
    assert sn.attribute_name('telephone') == 'mc:telephone'


def test_element():
    w = NameSpace('http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                  'w')
    ele = Element(w('document'))
    assert ele.to_string() == '<w:document/>'
    ele.set(width=w('100'))
    assert ele.to_string() == '<w:document w:width="100"/>'
    ele.add(Element(w('paragraph')))
    assert ele.to_string() == '<w:document w:width="100"><w:paragraph/></w:document>'
    ele.put('Hello World')
    assert ele.to_string() == '<w:document w:width="100">Hello World<w:paragraph/></w:document>'


def test_xml():
    ele = Element('document')
    xml = XML(ele)
    assert xml.to_string() == '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<document/>'
