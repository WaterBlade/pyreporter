from xml.etree.ElementTree import tostring, Element
from zipfile import ZipFile, ZIP_DEFLATED
from collections import namedtuple

DataFile = namedtuple('DataFile', 'path xml')
FigureFile = namedtuple('FigureFile', 'path figure')


# copied from the standard library
class TreeBuilder:
    """Generic element structure builder.

    This builder converts a sequence of start, data, and end method
    calls to a well-formed element structure.

    You can use this class to build an element structure using a custom XML
    parser, or a parser for some other XML-like format.

    *element_factory* is an optional element factory which is called
    to create new Element instances, as necessary.

    """

    def __init__(self, element_factory=None):
        self._data = []  # data collector
        self._elem = []  # element stack
        self._last = None  # last element
        self._tail = None  # true if we're after an end tag
        if element_factory is None:
            element_factory = Element
        self._factory = element_factory

    def close(self):
        """Flush builder buffers and return toplevel document Element."""
        assert len(self._elem) == 0, "missing end tags"
        assert self._last is not None, "missing toplevel element"
        return self._last

    def _flush(self):
        if self._data:
            if self._last is not None:
                text = "".join(self._data)
                if self._tail:
                    assert self._last.tail is None, "internal error (tail)"
                    self._last.tail = text
                else:
                    assert self._last.text is None, "internal error (text)"
                    self._last.text = text
            self._data = []

    def data(self, data):
        """Add text to current element."""
        self._data.append(data)

    def start(self, tag, attrs=None):
        """Open new element and return it.

        *tag* is the element name, *attrs* is a dict containing element
        attributes.

        """
        if attrs is None:
            attrs = {}
        self._flush()
        self._last = elem = self._factory(tag, attrs)
        if self._elem:
            self._elem[-1].append(elem)
        self._elem.append(elem)
        self._tail = 0
        return elem

    def end(self, tag):
        """Close and return current Element.

        *tag* is the element name.

        """
        self._flush()
        self._last = self._elem.pop()
        assert self._last.tag == tag, "end tag mismatch (expected %s, got %s)" % (self._last.tag, tag)
        self._tail = 1
        return self._last

    def add(self, element):
        """Add Element to last Element"""
        assert len(self._elem) > 0
        self._elem[-1].append(element)


def make_element(tag, *items) -> Element:
    ele = Element(tag)
    last_ele = None
    for item in items:
        if isinstance(item, Element):
            ele.append(item)
            last_ele = item
        elif isinstance(item, dict):
            ele.attrib.update(item)
        elif isinstance(item, str):
            if last_ele is None:
                ele.text = item
            else:
                last_ele.tail = item
    return ele


E = make_element


class SequenceGenerator:
    def __init__(self, prefix: str = '', start: int = 0):
        self.prefix = prefix
        self.index = start

    def get_id(self):
        ret = f'{self.prefix}{self.index}'
        self.index += 1
        return ret


class DocInterface:
    def add_heading(self):
        pass

    def add_paragraph(self):
        pass

    def add_figure(self):
        pass

    def add_table(self):
        pass

    def add_math(self):
        pass

    def add_definition(self):
        pass

    def add_procedure(self):
        pass

    def add_symbol_notes(self):
        pass

    def set_header(self):
        pass

    def set_cover(self):
        pass

    def get_bookmark(self):
        pass


class DocX:
    def __init__(self, path=''):
        self.path = path  # type: str

        self.document_xml_rels_tree = None  # type: TreeBuilder
        self.document_tree = None  # type: TreeBuilder
        self.footnotes_tree = None  # type: TreeBuilder

        self.file_list = list()
        self.figure_list = list()

        self.rel_id_gen = SequenceGenerator('rId', 1)
        self.foot_id_gen = SequenceGenerator('', 2)
        self.mark_id_gen = SequenceGenerator('', 1)

    def set_path(self, path):
        self.path = path

    def _add_xml(self, path: str, xml: str, with_head=True):
        if with_head:
            xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
        self.file_list.append(DataFile(path=path, xml=xml))

    def build_tree(self):
        self.build_document_xml_rels_tree()
        self.build_document_tree()
        self.build_footnotes_tree()

    def build_docx(self):
        self.build_content_type()
        self.build_rels()
        self.build_custom_xml()
        self.build_docProps()
        self.build_word()

    def build_content_type(self):
        xml = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"""
        self._add_xml(path='[Content_Types].xml', xml=xml)

    def build_rels(self):
        xml = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
        self._add_xml(path='_rels/.rels', xml=xml)

    def build_custom_xml(self):
        rels_xml = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>"""
        self._add_xml(path='customXml/_rels/item1.xml.rels', xml=rels_xml)

        item1_xml = """<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="\APA.XSL" StyleName="APA"/>"""
        self._add_xml(path='customXml/item1.xml', xml=item1_xml, with_head=False)

        props1_xml = """<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{16CE1163-3A16-4C5A-834F-C1E16F9680D6}"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/></ds:schemaRefs></ds:datastoreItem>"""
        self._add_xml(path='customXml/itemProps1.xml', xml=props1_xml)

    def build_docProps(self):
        xml = """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>0</TotalTime><Pages>1</Pages><Words>1</Words><Characters>12</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>12</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>"""
        self._add_xml(path='docProps/app.xml', xml=xml)

        xml = """<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Tao</dc:creator><cp:lastModifiedBy>HHPDI</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2019-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2019-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>"""
        self._add_xml(path='docProps/core.xml', xml=xml)

    def build_word(self):
        self.build_document_xml_rels()
        self.build_theme1()
        self.build_document()
        self.build_endnotes()
        self.build_fontTable()
        self.build_footer()
        self.build_footnotes()
        self.build_numbering()
        self.build_settings()
        self.build_styles()
        self.build_webSettings()

    def build_document_xml_rels_tree(self):
        ns = {'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        tree = TreeBuilder()
        self.document_xml_rels_tree = tree
        tree.start('Relationships', ns)

        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                                "theme/theme1.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                                "webSettings.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                                "fontTable.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                                "settings.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                                "styles.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
                                "endnotes.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
                                "footnotes.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                                "../customXml/item1.xml")
        self.write_relationship(self.rel_id_gen.get_id(),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                                "numbering.xml")
        footer_1 = self.rel_id_gen.get_id()
        footer_2 = self.rel_id_gen.get_id()
        self.write_relationship(footer_1,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                "footer1.xml")
        self.write_relationship(footer_2,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                "footer2.xml")
        self.footer_1_rel_id = footer_1
        self.footer_2_rel_id = footer_2

    def build_document_xml_rels(self):
        self.document_xml_rels_tree.end('Relationships')
        xml = tostring(self.document_xml_rels_tree.close(), encoding='unicode')
        self._add_xml(path='word/_rels/document.xml.rels', xml=xml)

    def build_theme1(self):
        xml = """<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ 明朝"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"""
        self._add_xml(path='word/theme/theme1.xml', xml=xml)

    def build_document_tree(self):
        ns = {'xmlns:ve': 'http://schemas.openxmlformats.org/markup_compatibility/2006',
              'xmlns:o': 'urn:schemas-microsoft-com:office:office',
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'xmlns:v': 'urn:schemas-microsoft-com:vml',
              'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
              'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml'}
        tree = TreeBuilder()
        self.document_tree = tree
        tree.start('w:document', ns)
        tree.start('w:body')

    def build_document(self):
        tree = self.document_tree
        self._write_w_sectPr(self.footer_2_rel_id, page_start=1)
        tree.end('w:body')
        tree.end('w:document')

        xml = tostring(self.document_tree.close(), encoding='unicode')
        self._add_xml(path='word/document.xml', xml=xml)

    def build_endnotes(self):
        xml = """<w:endnotes xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:endnote w:type="separator" w:id="0"><w:p w:rsidR="00E93427" w:rsidRDefault="00E93427" w:rsidP="00C825F3"><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p w:rsidR="00E93427" w:rsidRDefault="00E93427" w:rsidP="00C825F3"><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"""
        self._add_xml(path='word/endnotes.xml', xml=xml)

    def build_fontTable(self):
        xml = """<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font><w:font w:name="宋体"><w:altName w:val="SimSun"/><w:panose1 w:val="02010600030101010101"/><w:charset w:val="86"/><w:family w:val="auto"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="288F0000" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/></w:font><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002AFF" w:usb1="C0007841" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Cambria"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E00002FF" w:usb1="400004FF" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font><w:font w:name="楷体"><w:panose1 w:val="02010609060101010101"/><w:charset w:val="86"/><w:family w:val="modern"/><w:pitch w:val="fixed"/><w:sig w:usb0="800002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/></w:font><w:font w:name="Cambria Math"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E00002FF" w:usb1="420024FF" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font></w:fonts>"""
        self._add_xml(path='word/fontTable.xml', xml=xml)

    def build_footer(self):
        xml = """<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="003221F4" w:rsidRDefault="003221F4"><w:pPr><w:pStyle w:val="a5"/><w:jc w:val="center"/></w:pPr></w:p><w:p w:rsidR="003221F4" w:rsidRDefault="003221F4"><w:pPr><w:pStyle w:val="a5"/></w:pPr></w:p></w:ftr>"""
        self._add_xml(path='word/footer1.xml', xml=xml)

        xml = """<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:sdt><w:sdtPr><w:id w:val="23625823"/><w:docPartObj><w:docPartGallery w:val="Page Numbers (Bottom of Page)"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p w:rsidR="003221F4" w:rsidRDefault="00A805B6"><w:pPr><w:pStyle w:val="a5"/><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE   \* MERGEFORMAT "><w:r w:rsidR="006E593F" w:rsidRPr="006E593F"><w:rPr><w:noProof/><w:lang w:val="zh-CN"/></w:rPr><w:t>2</w:t></w:r></w:fldSimple></w:p></w:sdtContent></w:sdt><w:p w:rsidR="003221F4" w:rsidRDefault="003221F4"><w:pPr><w:pStyle w:val="a5"/></w:pPr></w:p></w:ftr>"""
        self._add_xml(path='word/footer2.xml', xml=xml)

    def build_footnotes_tree(self):
        ns = {'xmlns:ve': 'http://schemas.openxmlformats.org/markup_compatibility/2006',
              'xmlns:o': 'urn:schemas-microsoft-com:office:office',
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'xmlns:v': 'urn:schemas-microsoft-com:vml',
              'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
              'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml'}
        tree = TreeBuilder()
        self.footnotes_tree = tree
        tree.start('w:footnotes', ns)
        tree.add(E('w:footnote', {'w:type': 'separator', 'w:id': '0'},
                   E('w:p', E('w:r', E('w:separator')))))
        tree.add(E('w:footnote', {'w:type': 'continuationSeparator', 'w:id': '1'},
                   E('w:p', E('w:r', E('w:continuationSeparator')))))

    def build_footnotes(self):
        self.footnotes_tree.end('w:footnotes')
        xml = tostring(self.footnotes_tree.close(), encoding='unicode')
        self._add_xml(path='word/footnotes.xml', xml=xml)

    def build_numbering(self):
        xml = """<w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="5D861EAE"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="04090025"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="3"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="4"/><w:lvlText w:val="%1.%2.%3.%4"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="864" w:hanging="864"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="5"/><w:lvlText w:val="%1.%2.%3.%4.%5"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1008" w:hanging="1008"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="6"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1152" w:hanging="1152"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="7"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1296" w:hanging="1296"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="8"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="1440"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="9"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8.%9"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1584" w:hanging="1584"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""
        self._add_xml(path='word/numbering.xml', xml=xml)

    def build_settings(self):
        xml = """<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="160"/><w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/><w:defaultTabStop w:val="420"/><w:drawingGridVerticalSpacing w:val="156"/><w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="2"/><w:characterSpacingControl w:val="compressPunctuation"/><w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr><w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="1"/></w:endnotePr><w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/></w:compat><w:rsids><w:rsidRoot w:val="003221F4"/><w:rsid w:val="0023729F"/><w:rsid w:val="003221F4"/><w:rsid w:val="00581F8C"/><w:rsid w:val="005B06A3"/><w:rsid w:val="00626EC2"/><w:rsid w:val="006A65DB"/><w:rsid w:val="006E593F"/><w:rsid w:val="00926F4A"/><w:rsid w:val="00971FF4"/><w:rsid w:val="00A805B6"/><w:rsid w:val="00C501FC"/><w:rsid w:val="00D637A7"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="off"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>"""
        self._add_xml(path='word/settings.xml', xml=xml)

    def build_styles(self):
        xml = """<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="00971FF4"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="2Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="3Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="4"><w:name w:val="heading 4"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="4Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="3"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="5"><w:name w:val="heading 5"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="5Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="4"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="6"><w:name w:val="heading 6"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="6Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="5"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="7"><w:name w:val="heading 7"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="7Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="6"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="6"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="8"><w:name w:val="heading 8"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="8Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="7"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="7"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="9"><w:name w:val="heading 9"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="9Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="8"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="8"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="a2"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="character" w:customStyle="1" w:styleId="1Char"><w:name w:val="标题 1 Char"/><w:basedOn w:val="a0"/><w:link w:val="1"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="2Char"><w:name w:val="标题 2 Char"/><w:basedOn w:val="a0"/><w:link w:val="2"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="3Char"><w:name w:val="标题 3 Char"/><w:basedOn w:val="a0"/><w:link w:val="3"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="4Char"><w:name w:val="标题 4 Char"/><w:basedOn w:val="a0"/><w:link w:val="4"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="5Char"><w:name w:val="标题 5 Char"/><w:basedOn w:val="a0"/><w:link w:val="5"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="6Char"><w:name w:val="标题 6 Char"/><w:basedOn w:val="a0"/><w:link w:val="6"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="7Char"><w:name w:val="标题 7 Char"/><w:basedOn w:val="a0"/><w:link w:val="7"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="8Char"><w:name w:val="标题 8 Char"/><w:basedOn w:val="a0"/><w:link w:val="8"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="9Char"><w:name w:val="标题 9 Char"/><w:basedOn w:val="a0"/><w:link w:val="9"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="10"><w:name w:val="toc 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/></w:style><w:style w:type="paragraph" w:styleId="20"><w:name w:val="toc 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="200" w:left="420"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="30"><w:name w:val="toc 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style><w:style w:type="character" w:styleId="a3"><w:name w:val="Hyperlink"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:rPr><w:color w:val="0000FF" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a4"><w:name w:val="header"/><w:basedOn w:val="a"/><w:link w:val="Char"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char"><w:name w:val="页眉 Char"/><w:basedOn w:val="a0"/><w:link w:val="a4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a5"><w:name w:val="footer"/><w:basedOn w:val="a"/><w:link w:val="Char0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char0"><w:name w:val="页脚 Char"/><w:basedOn w:val="a0"/><w:link w:val="a5"/><w:uiPriority w:val="99"/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a6"><w:name w:val="Title"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="Char1"/><w:uiPriority w:val="10"/><w:qFormat/><w:rsid w:val="00C501FC"/><w:pPr><w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char1"><w:name w:val="标题 Char"/><w:basedOn w:val="a0"/><w:link w:val="a6"/><w:uiPriority w:val="10"/><w:rsid w:val="00C501FC"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a7"><w:name w:val="Document Map"/><w:basedOn w:val="a"/><w:link w:val="Char2"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char2"><w:name w:val="文档结构图 Char"/><w:basedOn w:val="a0"/><w:link w:val="a7"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a8"><w:name w:val="Balloon Text"/><w:basedOn w:val="a"/><w:link w:val="Char3"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char3"><w:name w:val="批注框文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a8"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a9"><w:name w:val="footnote text"/><w:basedOn w:val="a"/><w:link w:val="Char4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:pPr><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char4"><w:name w:val="脚注文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a9"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="006E593F"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="aa"><w:name w:val="footnote reference"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"""
        self._add_xml(path='word/styles.xml', xml=xml)

    def build_webSettings(self):
        xml = """<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>"""
        self._add_xml(path='word/webSettings.xml', xml=xml)

    def save(self, path_name=None):
        if path_name is not None:
            self.path = path_name

        with ZipFile(self.path, 'w', ZIP_DEFLATED) as z:
            for f in self.file_list:
                z.writestr(f.path, data=f.xml)
            for fig in self.figure_list:
                z.writestr(fig.path, data=fig.figure)

    def write_reporter(self, doc):
        self.build_tree()
        for data in doc.data_list:
            data.write_to(self)
        self.build_docx()
        self.save()

    def write_composite(self, datas):
        for data in datas:
            data.write_to(self)

    # ===============================================================
    # document.xml
    # ===============================================================

    def _write_w_sectPr(self, footer_rel_id, page_start=None):
        tree = self.document_tree
        tree.start('w:sectPr')

        tree.add(E('w:footerReference', {'w:type': 'default', 'r:id': footer_rel_id}))
        tree.add(E('w:pgSz', {'w:w': '11906', 'w:h': '16838'}))
        tree.add(E('w:pgMar', {'w:top': '1440', 'w:right': '1800',
                              'w:bottom': '1440', 'w:left': '1800',
                              'w:footer': '992', 'w:gutter': '0'}))
        if page_start is not None:
            tree.add(E('w:pgNumType', {'w:start': f'{page_start}'}))
        tree.add(E('w:cols', {'w:space': '425'}))
        tree.add(E('w:docGrid', {'w:type': 'lines', 'w:linePitch': '312'}))

        tree.end('w:sectPr')

    def write_page_break(self):
        self.document_tree.add(E('w:p', E('w:r', E('w:br', {'w:type': 'page'}))))

    def write_catalog(self, *, heads, catalog_level):
        tree = self.document_tree

        tree.add(E('w:p',
                   E('w:pPr', E('w:jc', {'w:val': 'center'})),
                   E('w:r',
                     E('w:rPr', E('w:b'), E('w:sz', {'w:val': '32'})),
                     E('w:t', '目录'))))

        first_head = True
        for head in heads:
            tree.start('w:p')
            tree.add(E('w:pPr',
                       E('w:pStyle', {'w:val': f'{head.level*10}'}),
                       E('w:tabs',
                         E('w:tab', {'w:val': 'left', 'w:pos': f'{420+630*(head.level-1)}'}),
                         E('w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': '8296'})),
                       E('w:rPr', E('w:noProof'), E('w:rStyle', {'w:val': 'a3'}))))
            if first_head:
                first_head = False
                tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
                tree.add(E('w:r', E('w:instrText', {'xml:space': 'preserve'}, ' ')))
                tree.add(E('w:r', E('w:instrText', f'TOC \\o "1-{catalog_level}" \\h \\z \\u')))
                tree.add(E('w:r', E('w:instrText', ' ', {'xml:space': 'preserve'})))
                tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
            self._write_heading_hyperlink(serial_code=head.serial_code,
                                          head=head.head,
                                          mark_name=f'_Mark{head.mark_id}')
            tree.end('w:p')

        tree.add(E('w:p', E('w:r', E('w:fldChar', {'w:fldCharType': 'end'}))))
        tree.start('w:p')
        tree.start('w:pPr')
        self._write_w_sectPr(self.footer_1_rel_id)
        tree.end('w:pPr')
        tree.end('w:p')

    def _write_heading_hyperlink(self, serial_code, head, mark_name):
        tree = self.document_tree
        tree.add(E('w:hyperlink',
                   {'w:anchor': mark_name, 'w:history': '1'},
                   E('w:r', E('w:t', serial_code)),
                   E('w:r', E('w:tab')),
                   E('w:r', E('w:t', head)),
                   E('w:r', E('w:tab')),
                   E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})),
                   E('w:r', E('w:instrText', {'xml:space': 'preserve'},
                              f' PAGEREF {mark_name} \\h ')),
                   E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})),
                   E('w:r', E('w:t', '0')),
                   E('w:r', E('w:fldChar', {'w:fldCharType': 'end'}))))

    def write_heading(self, *, para, level, need_page_break):
        if need_page_break:
            self.write_page_break()
        para.prop.set_style(f'{level}')
        para.write_to(self)

    def write_paragraph(self, *data, prop):
        tree = self.document_tree
        tree.start('w:p')

        prop.write_to(self)

        for item in data:
            item.write_to(self)

        tree.end('w:p')

    def write_para_prop(self, *, style, jc, spacing, snap_to_grid, font, size, keep_next):
        tree = self.document_tree
        tree.start('w:pPr')
        if style is not None:
            tree.add(E('w:pStyle', {'w:val': style}))
        if keep_next:
            tree.add(E('w:keepNext'))
        if jc:
            tree.add(E('w:jc', {'w:val': jc}))
        if snap_to_grid:
            tree.add(E('w:snapToGrid', {'w:val': snap_to_grid}))
        if spacing:
            spacing.write_to(self)
        if font:
            tree.add(E('w:rPr',
                       E('w:rFonts', {'w:eastAsia': font}),
                       E('w:sz', {'w:val': f'{size}'})))
        tree.end('w:pPr')

    def write_para_prop_spacing(self, before, after):
        attr = dict()
        if before is not None:
            attr['w:before'] = f'{before}'
        if after is not None:
            attr['w:after'] = f'{after}'

        self.document_tree.add(E('w:spacing', attr))

    def write_run(self, text: str, *,
                  size=None, font=None, italic=False, bold=False):
        tree = self.document_tree
        tree.start('w:r')

        tree.start('w:rPr')
        if font:
            tree.add(E('w:rFonts', {'w:eastAsia': font}))
        if bold:
            tree.add(E('w:b'))
        if italic:
            tree.add(E('w:i'))
        if size is not None:
            tree.add(E('w:sz', {'w:val': f'{size}'}))
        tree.end('w:rPr')

        tree.add(E('w:t', text))
        tree.end('w:r')

    def write_table(self, *, data_by_rows, prop, title):
        tree = self.document_tree
        if title is not None:
            title.prop.set_style('a9')
            title.prop.set_keep_next(True)
            title.write_to(self)

        tree.start('w:tbl')

        prop.write_to(self)

        for row in data_by_rows:
            tree.start('w:tr')
            for col in row:
                col.write_to(self)
            tree.end('w:tr')

        tree.end('w:tbl')

    def write_table_prop(self, *,
                         jc, style, layout, pos_pr, border, grid_col, cell_margin):
        tree = self.document_tree
        tree.start('w:tblPr')
        if jc:
            tree.add(E('w:jc', {'w:val': jc}))
        tree.add(E('w:tblLayout', {'w:type': layout}))
        if style:
            tree.add(E('w:tblStyle', {'w:val': style}))

        if pos_pr is not None:
            pos_pr.write_to(self)
        if border is not None:
            border.write_to(self)
        if cell_margin is not None:
            cell_margin.write_to(self)
        tree.end('w:tblPr')

        if grid_col is not None:
            grid_col.write_to(self)

    def write_table_prop_pos_pr(self, tblpX, tblpY, tblpXSpec, tblpYSpec):
        if tblpX or tblpXSpec or tblpY or tblpYSpec:
            attr = dict()
            attr['w:horizonAnchor'] = 'margin'
            if tblpX:
                attr['w:tblpX'] = f'{tblpX}'
            if tblpXSpec:
                attr['w:tblpXSpec'] = f'{tblpXSpec}'
            attr['w:vertAnchor'] = 'margin'
            if tblpY:
                attr['w:tblpY'] = f'{tblpY}'
            if tblpYSpec:
                attr['w:tblpYSpec'] = f'{tblpYSpec}'
            self.document_tree.add((E('w:tblpPr', attr)))

    def write_table_prop_border(self, top, bottom, left, right,
                                inside_v, inside_h, size=4):
        tree = self.document_tree
        tree.start('w:tblBorders')
        for exist, pos in zip([top, bottom, left, right, inside_v, inside_h],
                              ['top', 'bottom', 'left', 'right', 'insideV', 'insideH']):
            if exist:
                tree.add(E(f'w:{pos}',
                           {'w:val': 'single',
                            'w:sz': f'{size}',
                            'w:space': '0',
                            'w:color': 'auto'}))
        tree.end('w:tblBorders')

    def write_table_prop_grid_col(self, width_list: list = None):
        tree = self.document_tree
        tree.start('w:tblGrid')
        for width in width_list:
            tree.add(E('w:gridCol', {'w:w': f'{width}'}))
        tree.end('w:tblGrid')

    def write_table_prop_cell_margin(self, top, bottom, left, right):
        self.document_tree.add(E('w:tblCellMar',
                                 E('w:left', {'w:w': f'{left}', 'w:type': 'dxa'}),
                                 E('w:right', {'w:w': f'{right}', 'w:type': 'dxa'}),
                                 E('w:top', {'w:w': f'{top}', 'w:type': 'dxa'}),
                                 E('w:bottom', {'w:w': f'{bottom}', 'w:type': 'dxa'})))

    def write_cell(self, *, para, prop):
        tree = self.document_tree
        tree.start('w:tc')

        prop.write_to(self)

        if prop.jc:
            para.prop.set_jc(prop.jc)
        para.write_to(self)

        tree.end('w:tc')

    def write_cell_prop(self, v_align, border):
        tree = self.document_tree
        tree.start('w:tcPr')
        tree.add(E('w:tcW', {'w:w': '0', 'w:type': 'auto'}))
        tree.add(E('w:vAlign', {'w:val': v_align}))
        if border:
            border.write_to(self)
        tree.end('w:tcPr')

    def write_cell_prop_border(self, top, bottom, left, right, size):
        tree = self.document_tree
        tree.start('w:tcBorders')
        for exist, pos in zip([top, bottom, left, right],
                              ['top', 'bottom', 'left', 'right']):
            if exist:
                tree.add(E(f'w:{pos}',
                           {'w:val': 'single',
                            'w:sz': f'{size}',
                            'w:space': '0',
                            'w:color': 'auto'}))
        tree.end('w:tcBorders')

    def write_footnote(self, datas):
        foot_id = self.foot_id_gen.get_id()

        self.document_tree.add(E('w:r',
                                 E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                                 E('w:footnoteReference', {'w:id': foot_id})))

        tree = self.footnotes_tree

        tree.start('w:footnote', {'w:id': foot_id})

        tree.start('w:p')
        tree.add(E('w:pPr', E('w:pStyle', {'w:val': 'a9'})))
        tree.add(E('w:r',
                   E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                   E('w:footnoteRef')))
        tree.add(E('w:r', {'xml:space': 'preserve'}, ' '))

        for data in datas:
            data.write_to(self)

        tree.end('w:p')

        tree.end('w:footnote')

    def write_bookmark(self, type_, mark_id):
        mark_id = f'{mark_id}'
        mark_name = f'_Mark{mark_id}'
        tree = self.document_tree

        tree.add(E('w:bookmarkStart', {'w:id': mark_id, 'w:name': mark_name}))
        tree.add(E('w:r', E('w:t', type_)))
        tree.add(E('w:r', E('w:t', '', {'xml:space': 'preserve'})))
        tree.add(E('w:fldSimple',
                   E('w:r', E('w:t', '0')), {'w:instr': r' STYLEREF 1 \s'}))
        tree.add(E('w:r', E('w:noBreakHyphen')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
        tree.add(E('w:r',
                   E('w:instrText', ' ', {'xml:space': 'preserve'})))
        tree.add(E('w:r',
                   E('w:instrText', 'SEQ ', {'xml:space': 'preserve'})))
        tree.add(E('w:r', E('w:instrText', type_)))
        tree.add(E('w:r',
                   E('w:instrText', r' \* ARABIC \s 1', {'xml:space': 'preserve'})))
        tree.add(E('w:r', E('w:instrText', '', {'xml:space': 'preserve'})))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
        tree.add(E('w:r', E('w:t', '1')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        tree.add(E('w:bookmarkEnd', {'w:id': mark_id}))

    def write_heading_mark(self, *, head, mark_id):
        mark_id = f'{mark_id}'
        mark_name = f'_Mark{mark_id}'

        tree = self.document_tree
        tree.add(E('w:bookmarkStart', {'w:id': mark_id, 'w:name': mark_name}))
        tree.add(E('w:r', E('w:t', head)))
        tree.add(E('w:bookmarkEnd', {'w:id': mark_id}))

    def write_reference(self, mark_id):
        mark_name = f'_Mark{mark_id}'
        tree = self.document_tree
        tree.add(E('w:r', E('w:t', '(')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
        tree.add(E('w:r', E('w:instrText', f'REF {mark_name} \\h')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
        tree.add(E('w:r', E('w:t', '0')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        tree.add(E('w:r', E('w:t', ')')))

    def write_inline_figure(self, *, data, fmt, cx, cy):
        cx = int(cx * 914400)
        cy = int(cy * 914400)

        fig_id = len(self.figure_list) + 1
        rel_id = self.rel_id_gen.get_id()

        self.figure_list.append(FigureFile(path=f'word/media/image{fig_id}.{fmt}', figure=data))
        self.write_relationship(id_=rel_id,
                                type_='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                target=f'media/image{fig_id}.{fmt}')

        tree = self.document_tree

        tree.start('w:r')

        tree.start('w:drawing')

        tree.start('wp:inline')

        tree.add(E('wp:extent', {'cx': f'{cx}', 'cy': f'{cy}'}))
        tree.add(E('wp:effectExtent', {'l': '0', 't': '0', 'r': '0', 'b': '0'}))
        tree.add(E('wp:docPr', {'id': f'{fig_id}', 'name': f'image{fig_id}'}))

        tree.add(E('wp:cNvGraphicFramePr',
                   E('a:graphicFrameLocks',
                     {'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main", 'noChangeAspect': '1'})))

        self._write_a_graphic(fig_id, rel_id, cx, cy)

        tree.end('wp:inline')

        tree.end('w:drawing')

        tree.end('w:r')

    def write_standalone_figure(self, *, title, data, fmt, cx, cy):
        tree = self.document_tree
        tree.start('w:p')
        tree.start('w:pPr')
        tree.add(E('w:jc', {'w:val': 'center'}))
        if title is not None:
            tree.add(E('w:keepNext'))
        tree.end('w:pPr')
        self.write_inline_figure(data=data, fmt=fmt, cx=cx, cy=cy)
        tree.end('w:p')

        if title is not None:
            title.prop.set_style('a9')
            title.write_to(self)

    def write_float_figure(self, *, data, fmt, cx, cy,
                           xoff=0, yoff=0, xspec=None, yspec=None):
        cx = int(cx * 914400)
        cy = int(cy * 914400)
        xoff = int(xoff * 914400)
        yoff = int(yoff * 914400)

        fig_id = len(self.figure_list) + 1
        rel_id = self.rel_id_gen.get_id()

        self.figure_list.append(FigureFile(path=f'word/media/image{fig_id}.{fmt}', figure=data))
        self.write_relationship(id_=rel_id,
                                type_='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                target=f'media/image{fig_id}.{fmt}')

        if xspec is not None:
            xpos = E('wp:align', xspec)
        else:
            xpos = E('wp:posOffset', f'{xoff}')

        if yspec is not None:
            ypos = E('wp:align', yspec)
        else:
            ypos = E('wp:posOffset', f'{yoff}')

        tree = self.document_tree

        tree.start('w:p')
        tree.start('w:drawing')
        tree.start('wp:anchor',
                   {'simplePos': '0',
                    'relativeHeight': '1',
                    'behindDoc': '0',
                    'locked': '0',
                    'layoutInCell': '1',
                    'allowOverlap': '1'})

        tree.add(E('wp:simplePos', {'x': '0', 'y': '0'}))
        tree.add(E('wp:positionH', {'relativeFrom': 'margin'}, xpos))
        tree.add(E('wp:positionV', {'relativeFrom': 'margin'}, ypos))

        tree.add(E('wp:extent', {'cx': f'{cx}', 'cy': f'{cy}'}))
        tree.add(E('wp:effectExtent', {'l': '0', 't': '0', 'r': '0', 'b': '0'}))
        tree.add(E('wp:wrapTopAndBottom'))
        tree.add(E('wp:docPr', {'id': f'{fig_id}', 'name': f'image{fig_id}'}))

        tree.add(E('wp:cNvGraphicFramePr',
                   E('a:graphicFrameLocks',
                     {'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main", 'noChangeAspect': 'true'})))

        self._write_a_graphic(fig_id, rel_id, cx, cy)

        tree.end('wp:anchor')
        tree.end('w:drawing')
        tree.end('w:p')

    def _write_a_graphic(self, fig_id, rel_id, cx, cy):
        ele = E('a:graphic', {'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main"},
                E('a:graphicData', {'uri': "http://schemas.openxmlformats.org/drawingml/2006/picture"},
                  E('pic:pic', {'xmlns:pic': "http://schemas.openxmlformats.org/drawingml/2006/picture"},
                    E('pic:nvPicPr',
                      E('pic:cNvPr', {'id': f'{fig_id}', 'name': f'image{fig_id}'}),
                      E('pic:cNvPicPr')),
                    E('pic:blipFill',
                      E('a:blip', {'r:embed': rel_id}),
                      E('a:stretch',
                        E('a:fillRect'))),
                    E('pic:spPr',
                      E('a:xfrm',
                        E('a:off', {'x': '0', 'y': '0'}),
                        E('a:ext', {'cx': f'{cx}', 'cy': f'{cy}'})),
                      E('a:prstGeom', {'prst': 'rect'},
                        E('a:avLst'))))))
        self.document_tree.add(ele)

    # math related
    def write_math_composite(self, datas):
        for data in datas:
            data.write_to(self)

    def write_mathpara(self, math_obj):
        tree = self.document_tree
        tree.start('m:oMathPara')
        math_obj.write_to(self)
        tree.end('m:oMathPara')

    def write_math(self, datas):
        tree = self.document_tree
        tree.start('m:oMath')
        for item in datas:
            item.write_to(self)
        tree.end('m:oMath')

    def write_math_multi_line(self, exps):
        tree = self.document_tree
        tree.start('m:eqArr')

        tree.start('m:eqArrPr')
        self._write_m_ctrlPr()
        tree.end('m:eqArrPr')

        for exp in exps:
            tree.start('m:e')
            exp.write_to(self)
            tree.end('m:e')

        tree.end('m:eqArr')

    def write_math_run(self, run: str, sty=None):
        tree = self.document_tree
        tree.start('m:r')

        if sty:
            tree.add(E('m:rPr', E('m:sty', {'m:val': sty})))
        tree.add(E('w:rPr',
                   E('w:rFonts', {'w:ascii': 'Cambria Math', 'w:hAnsi': 'Cambria Math'})))

        tree.add(E('m:t', run))
        tree.end('m:r')

    def write_symbol_notes(self, notes):
        for note in notes:
            self.write_symbol_note(note[0], note[1], note[2])

    def write_symbol_note(self, var, inform, unit=None):
        tree = self.document_tree
        tree.start('w:p')

        tree.add(E('w:pPr',
                   E('w:tabs',
                     E('w:tab', {'w:val': 'right', 'w:pos': '500'}),
                     E('w:tab', {'w:val': 'center', 'w:pos':'600'}),
                     E('w:tab', {'w:val': 'left', 'w:pos': '700'}))))
        tree.add(E('w:r', E('w:tab')))
        self.write_math([var])
        tree.add(E('w:r', E('w:tab')))
        self.write_run('―')
        tree.add(E('w:r', E('w:tab')))
        self.write_run(inform)
        if unit is not None:
            self.write_run('，单位：')
            self.write_math([unit])

        tree.end('w:p')

    def _write_m_ctrlPr(self, italic=False):
        tree = self.document_tree
        tree.start('m:ctrlPr')

        tree.start('w:rPr')
        tree.add(E('w:rFonts', {'w:ascii': 'Cambria Math', 'w:hAnsi':'Cambria Math'}))
        if italic:
            tree.add(E('w:i'))
        tree.end('w:rPr')

        tree.end('m:ctrlPr')

    def write_variable(self, var):
        self.write_math_run(var)

    def write_unit(self, unit):
        self.write_math_run(unit, sty='p')

    def write_subscript_variable(self, base, sub):
        self._write_m_sSup(base, sub)

    def write_add(self, left, right):
        left.write_to(self)
        self.write_math_run('+', sty='p')
        right.write_to(self)

    def write_sub(self, left, right):
        left.write_to(self)
        self.write_math_run('-', sty='p')
        right.write_to(self)

    def write_mul(self, left, right):
        left.write_to(self)
        self.write_math_run('⋅', sty='p')
        right.write_to(self)

    def _write_named_element(self, name, exp):
        tree = self.document_tree
        tree.start(name)
        if exp is not None:
            exp.write_to(self)
        tree.end(name)

    def write_div(self, left, right):
        tree = self.document_tree

        tree.start('m:f')
        tree.start('m:fPr')
        tree.add(E('m:type', {'m:val': 'bar'}))
        self._write_m_ctrlPr()
        tree.end('m:fPr')
        self._write_named_element('m:num', left)
        self._write_named_element('m:den', right)
        tree.end('m:f')

    def write_flat_div(self, left, right):
        tree = self.document_tree

        tree.start('m:f')
        tree.start('m:fPr')
        tree.add(E('m:type', {'m:val': 'lin'}))
        self._write_m_ctrlPr()
        tree.end('m:fPr')
        self._write_named_element('m:num', left)
        self._write_named_element('m:den', right)
        tree.end('m:f')

    def write_radical(self, exp, index):
        tree = self.document_tree
        tree.start('m:rad')

        tree.start('m:radPr')
        if index.value == 2:
            tree.add(E('m:degHide', {'m:val': 'on'}))
        self._write_m_ctrlPr()
        tree.end('m:radPr')

        self._write_named_element('m:deg', index if index.value != 2 else None)
        self._write_named_element('m:e', exp)

        tree.end('m:rad')

    def write_lesser_than(self, left, right):
        left.write_to(self)
        self.write_math_run('<', sty='p')
        right.write_to(self)

    def write_lesser_or_equal(self, left, right):
        left.write_to(self)
        self.write_math_run('≤', sty='p')
        right.write_to(self)

    def write_equal(self, left, right):
        left.write_to(self)
        self.write_math_run('=', sty='p')
        right.write_to(self)

    def write_not_equal(self, left, right):
        left.write_to(self)
        self.write_math_run('≠', sty='p')
        right.write_to(self)

    def write_greater_than(self, left, right):
        left.write_to(self)
        self.write_math_run('>', sty='p')
        right.write_to(self)

    def write_greater_or_equal(self, left, right):
        left.write_to(self)
        self.write_math_run('≥', sty='p')
        right.write_to(self)

    def _write_m_sSup(self, base, sup):
        tree = self.document_tree
        tree.start('m:sSup')

        tree.start('m:sSupPr')
        self._write_m_ctrlPr()
        tree.end('m:sSupPr')

        self._write_named_element('m:e', base)
        self._write_named_element('m:sup', sup)

        tree.end('m:sSup')

    def _write_m_sSubSup(self, base, sub, sup):
        tree = self.document_tree
        tree.start('m:sSubSup')

        tree.start('m:sSubSupPr')
        self._write_m_ctrlPr()
        tree.end('m:sSubSupPr')

        self._write_named_element('m:e', base)
        self._write_named_element('m:sub', sub)
        self._write_named_element('m:sup', sup)

        tree.end('m:sSubSup')

    def _write_m_nary(self, exp, sub, sup, name=None, limLoc='subSup'):
        tree = self.document_tree
        tree.start('m:nary')

        tree.start('m:naryPr')
        if name:
            tree.add(E('m:chr', {'m:val': name}))
        tree.add(E('m:limLoc', {'m:val': limLoc}))
        self._write_m_ctrlPr(italic=True)
        tree.end('m:naryPr')

        self._write_named_element('m:sub', sub)
        self._write_named_element('m:sup', sup)
        self._write_named_element('m:e', exp)

        tree.end('m:nary')

    def write_pow(self, exp, index):
        self._write_m_sSup(exp, index)

    def write_pow_with_sub(self, exp, sub, index):
        self._write_m_sSubSup(exp, sub, index)

    def write_sum(self, exp, sub, sup):
        self._write_m_nary(exp, sub, sup, name='∑', limLoc='undOvr')

    def write_integrate(self, exp, var, sub, sup):
        # TODO: 需要完善积分与微分运算部分，微分运算似乎有点复杂。
        pass

    def _write_m_func(self, name, exp):
        tree = self.document_tree
        tree.start('m:func')

        tree.start('m:funcPr')
        self._write_m_ctrlPr()
        tree.end('m:funcPr')

        tree.start('m:fName')
        self.write_math_run(name, sty='p')
        tree.end('m:fName')

        self._write_named_element('m:e', exp)

        tree.end('m:func')

    def write_sin(self, exp):
        self._write_m_func('sin', exp)

    def write_cos(self, exp):
        self._write_m_func('cos', exp)

    def write_tan(self, exp):
        self._write_m_func('tan', exp)

    def write_cot(self, exp):
        self._write_m_func('cot', exp)

    def write_asin(self, exp):
        self._write_m_func('arcsin', exp)

    def write_acos(self, exp):
        self._write_m_func('arccos', exp)

    def write_atan(self, exp):
        self._write_m_func('arctan', exp)

    def write_acot(self, exp):
        self._write_m_func('arccot', exp)

    def _write_m_delimeter(self, *exps, left='(', right=')', sep='|'):
        tree = self.document_tree
        tree.start('m:d')

        tree.start('m:dPr')
        if left is not None:
            tree.add(E('m:begChr', {'m:val': left}))
        if right is not None:
            tree.add(E('m:endChr', {'m:val': right}))
        self._write_m_ctrlPr()
        tree.end('m:dPr')

        for exp in exps:
            self._write_named_element('m:e', exp)

        tree.end('m:d')

    def write_parenthesis(self, exp):
        self._write_m_delimeter(exp, left='(', right=')')

    def write_square_bracket(self, exp):
        self._write_m_delimeter(exp, left='[', right=']')

    def write_brace(self, exp):
        self._write_m_delimeter(exp, left='{', right='}')

    def write_multi_line_brace(self, exp):
        self._write_m_delimeter(exp, left='{', right=None)



    # ===============================================================
    # document.xml
    # ===============================================================

    # ===============================================================
    # document.xml.rels
    # ===============================================================

    def write_relationship(self, id_, type_, target):
        tree = self.document_xml_rels_tree
        tree.add(E('Relationship',
                   {'Id': id_, 'Type': type_, 'Target': target}))
