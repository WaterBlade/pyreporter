from xml.etree.ElementTree import tostring, Element
from zipfile import ZipFile, ZIP_DEFLATED
from collections import namedtuple

DataFile = namedtuple('DataFile', 'path xml')
FigureFile = namedtuple('FigureFile', 'path figure')
Id = namedtuple('Id', 'number text')


TEXT_SIZE = {'normal': 12, 'large': 18, 'small': 8}


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

    def is_closed(self):
        return len(self._elem) == 0

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


class ElementsBuilder:
    def __init__(self):
        self._curr = None  # type: TreeBuilder
        self._list = []

    def is_empty(self):
        return len(self._list) == 0

    def start(self, tag, attrs=None):
        if self._curr is not None:
            self._curr.start(tag, attrs)
        else:
            tree = TreeBuilder()
            tree.start(tag, attrs)
            self._curr = tree

    def end(self, tag):
        if self._curr is None:
            raise RuntimeError('mismatching')
        else:
            self._curr.end(tag)
            if self._curr.is_closed():
                self._list.append(self._curr.close())
                self._curr = None

    def add(self, element):
        if self._curr is None:
            self._list.append(element)
        else:
            self._curr.add(element)

    def get_elements(self):
        return self._list


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
        elif isinstance(item, list):
            ele.extend(item)
    return ele


E = make_element


class DocX:
    def __init__(self):
        self.file_list = list()
        self.figure_list = list()

        self.rel_id = 1
        self.footnote_id = 2
        self.mark_id = 1
        self.mark_id_dict = dict()

        self.doc_rels_elements = ElementsBuilder()

        self.cover_elements = ElementsBuilder()
        self.catalog_elements = ElementsBuilder()
        self.body_elements = ElementsBuilder()

        self.footnotes_elements = ElementsBuilder()
        self.header_elements = ElementsBuilder()

    def save(self, path):
        self._build_docx()
        with ZipFile(path, 'w', ZIP_DEFLATED) as z:
            for f in self.file_list:
                z.writestr(f.path, data=f.xml)
            for fig in self.figure_list:
                z.writestr(fig.path, data=fig.figure)

    def _build_docx(self):
        # Dir: .
        self._build_content_type()
        self._build_rels()
        self._build_custom_xml()
        self._build_docProps()
        # Dir: ./word
        self._build_document_xml_rels()
        self._build_theme1()
        self._build_document()
        self._build_endnotes()
        self._build_fontTable()
        self._build_footer()
        self._build_footnotes()
        self._build_header()
        self._build_numbering()
        self._build_settings()
        self._build_styles()
        self._build_webSettings()

    def _get_rel_id(self):
        ret = Id(f'{self.rel_id}', f'rId{self.rel_id}')
        self.rel_id += 1
        return ret

    def _get_footnote_id(self):
        ret = Id(f'{self.footnote_id}', f'{self.footnote_id}')
        self.footnote_id += 1
        return ret

    def _get_mark_id(self):
        ret = Id(f'{self.mark_id}', f'_Mark{self.mark_id}')
        self.mark_id += 1
        return ret

    def _add_xml(self, path: str, xml: str, with_head=True):
        if with_head:
            xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
        self.file_list.append(DataFile(path=path, xml=xml))

    def _build_content_type(self):
        xml = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"""
        self._add_xml(path='[Content_Types].xml', xml=xml)

    def _build_rels(self):
        xml = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
        self._add_xml(path='_rels/.rels', xml=xml)

    def _build_custom_xml(self):
        rels_xml = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>"""
        self._add_xml(path='customXml/_rels/item1.xml.rels', xml=rels_xml)

        item1_xml = """<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="\APA.XSL" StyleName="APA"/>"""
        self._add_xml(path='customXml/item1.xml', xml=item1_xml, with_head=False)

        props1_xml = """<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{16CE1163-3A16-4C5A-834F-C1E16F9680D6}"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/></ds:schemaRefs></ds:datastoreItem>"""
        self._add_xml(path='customXml/itemProps1.xml', xml=props1_xml)

    def _build_docProps(self):
        xml = """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>0</TotalTime><Pages>1</Pages><Words>1</Words><Characters>12</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>12</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>"""
        self._add_xml(path='docProps/app.xml', xml=xml)

        xml = """<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Tao</dc:creator><cp:lastModifiedBy>HHPDI</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2019-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2019-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>"""
        self._add_xml(path='docProps/core.xml', xml=xml)

    def _build_document_xml_rels(self):
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                                "theme/theme1.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                                "webSettings.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                                "fontTable.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                                "settings.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                                "styles.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
                                "endnotes.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
                                "footnotes.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                                "../customXml/item1.xml")
        self._write_relationship(self._get_rel_id().text,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                                "numbering.xml")

        self.footer_rel_id = self._get_rel_id().text
        self.header_rel_id = self._get_rel_id().text

        self._write_relationship(self.footer_rel_id,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                "footer1.xml")
        self._write_relationship(self.header_rel_id,
                                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                                 "header1.xml")

        ns = {'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        root = E('Relationships', ns)
        root.extend(self.doc_rels_elements.get_elements())
        xml = tostring(root, encoding='unicode')
        self._add_xml(path='word/_rels/document.xml.rels', xml=xml)

    def _build_theme1(self):
        xml = """<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ 明朝"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"""
        self._add_xml(path='word/theme/theme1.xml', xml=xml)

    def _build_document(self):

        ns = {'xmlns:ve': 'http://schemas.openxmlformats.org/markup_compatibility/2006',
              'xmlns:o': 'urn:schemas-microsoft-com:office:office',
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'xmlns:v': 'urn:schemas-microsoft-com:vml',
              'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
              'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml'}
        doc = E('w:document', ns)

        body = E('w:body')

        if not self.cover_elements.is_empty():
            body.extend(self.cover_elements.get_elements())
            body.append(self._make_page_break())

        if not self.catalog_elements.is_empty():
            body.extend(self.catalog_elements.get_elements())
            body.append(E('w:p', E('w:pPr', self._make_w_sectPr())))

        body.extend(self.body_elements.get_elements())
        body.append(self._make_w_sectPr(None,
                                        self.footer_rel_id,
                                        page_start=1))

        doc.append(body)

        xml = tostring(doc, encoding='unicode')
        self._add_xml(path='word/document.xml', xml=xml)

    def _build_endnotes(self):
        xml = """<w:endnotes xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:endnote w:type="separator" w:id="0"><w:p w:rsidR="00E93427" w:rsidRDefault="00E93427" w:rsidP="00C825F3"><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p w:rsidR="00E93427" w:rsidRDefault="00E93427" w:rsidP="00C825F3"><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"""
        self._add_xml(path='word/endnotes.xml', xml=xml)

    def _build_fontTable(self):
        xml = """<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font><w:font w:name="宋体"><w:altName w:val="SimSun"/><w:panose1 w:val="02010600030101010101"/><w:charset w:val="86"/><w:family w:val="auto"/><w:pitch w:val="variable"/><w:sig w:usb0="00000003" w:usb1="288F0000" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/></w:font><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002AFF" w:usb1="C0007841" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Cambria"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E00002FF" w:usb1="400004FF" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font><w:font w:name="楷体"><w:panose1 w:val="02010609060101010101"/><w:charset w:val="86"/><w:family w:val="modern"/><w:pitch w:val="fixed"/><w:sig w:usb0="800002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/></w:font><w:font w:name="Cambria Math"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E00002FF" w:usb1="420024FF" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font></w:fonts>"""
        self._add_xml(path='word/fontTable.xml', xml=xml)

    def _build_footer(self):
        xml = """<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:sdt><w:sdtPr><w:id w:val="23625823"/><w:docPartObj><w:docPartGallery w:val="Page Numbers (Bottom of Page)"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p w:rsidR="003221F4" w:rsidRDefault="00A805B6"><w:pPr><w:pStyle w:val="a5"/><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE   \* MERGEFORMAT "><w:r w:rsidR="006E593F" w:rsidRPr="006E593F"><w:rPr><w:noProof/><w:lang w:val="zh-CN"/></w:rPr><w:t>2</w:t></w:r></w:fldSimple></w:p></w:sdtContent></w:sdt><w:p w:rsidR="003221F4" w:rsidRDefault="003221F4"><w:pPr><w:pStyle w:val="a5"/></w:pPr></w:p></w:ftr>"""
        self._add_xml(path='word/footer1.xml', xml=xml)

    def _build_footnotes(self):
        ns = {'xmlns:ve': 'http://schemas.openxmlformats.org/markup_compatibility/2006',
              'xmlns:o': 'urn:schemas-microsoft-com:office:office',
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'xmlns:v': 'urn:schemas-microsoft-com:vml',
              'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
              'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml'}
        footnotes = E('w:footnotes', ns)

        footnotes.append(E('w:footnote', {'w:type': 'separator', 'w:id': '0'},
                           E('w:p', E('w:r', E('w:separator')))))
        footnotes.append(E('w:footnote', {'w:type': 'continuationSeparator', 'w:id': '1'},
                           E('w:p', E('w:r', E('w:continuationSeparator')))))

        footnotes.extend(self.footnotes_elements.get_elements())

        xml = tostring(footnotes, encoding='unicode')
        self._add_xml(path='word/footnotes.xml', xml=xml)

    def _build_header(self):
        ns = {'xmlns:ve': 'http://schemas.openxmlformats.org/markup_compatibility/2006',
              'xmlns:o': 'urn:schemas-microsoft-com:office:office',
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'xmlns:v': 'urn:schemas-microsoft-com:vml',
              'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
              'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml'}
        hdr = E('w:hdr', ns)
        hdr.extend(self.header_elements.get_elements())
        xml = tostring(hdr, encoding='unicode')
        self._add_xml(path='word/header1.xml', xml=xml)

    def _build_numbering(self):
        xml = """<w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="5D861EAE"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="04090025"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="3"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="4"/><w:lvlText w:val="%1.%2.%3.%4"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="864" w:hanging="864"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="5"/><w:lvlText w:val="%1.%2.%3.%4.%5"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1008" w:hanging="1008"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="6"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1152" w:hanging="1152"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="7"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1296" w:hanging="1296"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="8"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="1440"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="9"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8.%9"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1584" w:hanging="1584"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""
        self._add_xml(path='word/numbering.xml', xml=xml)

    def _build_settings(self):
        xml = """<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="160"/><w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/><w:defaultTabStop w:val="420"/><w:drawingGridVerticalSpacing w:val="156"/><w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="2"/><w:characterSpacingControl w:val="compressPunctuation"/><w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr><w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="1"/></w:endnotePr><w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/></w:compat><w:rsids><w:rsidRoot w:val="003221F4"/><w:rsid w:val="0023729F"/><w:rsid w:val="003221F4"/><w:rsid w:val="00581F8C"/><w:rsid w:val="005B06A3"/><w:rsid w:val="00626EC2"/><w:rsid w:val="006A65DB"/><w:rsid w:val="006E593F"/><w:rsid w:val="00926F4A"/><w:rsid w:val="00971FF4"/><w:rsid w:val="00A805B6"/><w:rsid w:val="00C501FC"/><w:rsid w:val="00D637A7"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="off"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>"""
        self._add_xml(path='word/settings.xml', xml=xml)

    def _build_styles(self):
        xml = """<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="00971FF4"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="2Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="3Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="4"><w:name w:val="heading 4"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="4Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="3"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="5"><w:name w:val="heading 5"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="5Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="4"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="6"><w:name w:val="heading 6"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="6Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="5"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="7"><w:name w:val="heading 7"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="7Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="6"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="6"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="8"><w:name w:val="heading 8"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="8Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="7"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="7"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="9"><w:name w:val="heading 9"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="9Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="8"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="8"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="a2"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="character" w:customStyle="1" w:styleId="1Char"><w:name w:val="标题 1 Char"/><w:basedOn w:val="a0"/><w:link w:val="1"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="2Char"><w:name w:val="标题 2 Char"/><w:basedOn w:val="a0"/><w:link w:val="2"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="3Char"><w:name w:val="标题 3 Char"/><w:basedOn w:val="a0"/><w:link w:val="3"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="4Char"><w:name w:val="标题 4 Char"/><w:basedOn w:val="a0"/><w:link w:val="4"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="5Char"><w:name w:val="标题 5 Char"/><w:basedOn w:val="a0"/><w:link w:val="5"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="6Char"><w:name w:val="标题 6 Char"/><w:basedOn w:val="a0"/><w:link w:val="6"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="7Char"><w:name w:val="标题 7 Char"/><w:basedOn w:val="a0"/><w:link w:val="7"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="8Char"><w:name w:val="标题 8 Char"/><w:basedOn w:val="a0"/><w:link w:val="8"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="9Char"><w:name w:val="标题 9 Char"/><w:basedOn w:val="a0"/><w:link w:val="9"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="10"><w:name w:val="toc 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/></w:style><w:style w:type="paragraph" w:styleId="20"><w:name w:val="toc 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="200" w:left="420"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="30"><w:name w:val="toc 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style><w:style w:type="character" w:styleId="a3"><w:name w:val="Hyperlink"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:rPr><w:color w:val="0000FF" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a4"><w:name w:val="header"/><w:basedOn w:val="a"/><w:link w:val="Char"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char"><w:name w:val="页眉 Char"/><w:basedOn w:val="a0"/><w:link w:val="a4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a5"><w:name w:val="footer"/><w:basedOn w:val="a"/><w:link w:val="Char0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char0"><w:name w:val="页脚 Char"/><w:basedOn w:val="a0"/><w:link w:val="a5"/><w:uiPriority w:val="99"/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a6"><w:name w:val="Title"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="Char1"/><w:uiPriority w:val="10"/><w:qFormat/><w:rsid w:val="00C501FC"/><w:pPr><w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char1"><w:name w:val="标题 Char"/><w:basedOn w:val="a0"/><w:link w:val="a6"/><w:uiPriority w:val="10"/><w:rsid w:val="00C501FC"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a7"><w:name w:val="Document Map"/><w:basedOn w:val="a"/><w:link w:val="Char2"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char2"><w:name w:val="文档结构图 Char"/><w:basedOn w:val="a0"/><w:link w:val="a7"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a8"><w:name w:val="Balloon Text"/><w:basedOn w:val="a"/><w:link w:val="Char3"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char3"><w:name w:val="批注框文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a8"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a9"><w:name w:val="footnote text"/><w:basedOn w:val="a"/><w:link w:val="Char4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:pPr><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char4"><w:name w:val="脚注文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a9"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="006E593F"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="aa"><w:name w:val="footnote reference"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"""
        self._add_xml(path='word/styles.xml', xml=xml)

    def _build_webSettings(self):
        xml = """<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>"""
        self._add_xml(path='word/webSettings.xml', xml=xml)


    # ===============================================================
    # document.xml.rels START
    # ---------------------------------------------------------------

    def _write_relationship(self, id_, type_, target):
        self.doc_rels_elements.add(E('Relationship', {'Id': id_, 'Type': type_, 'Target': target}))

    # ---------------------------------------------------------------
    # document.xml.rels END
    # ===============================================================

    # ===============================================================
    # header1.xml START
    # ---------------------------------------------------------------

    def visit_header(self, header):
        tree = self.header_elements
        tree.add(E('w:p',
                   E('w:pPr',
                     E('w:pStyle', {'w:val': 'a4'}),
                     E('w:rPr', E('w:rFonts', {'w:eastAsia': '楷体'}))),
                   E('w:r',
                     E('w:t', header))))

    # ---------------------------------------------------------------
    # header1.xml END
    # ===============================================================

    # ===============================================================
    # document.xml START
    # ---------------------------------------------------------------

    def visit_text(self, text, font, size, italic, bold, underline):
        tree = self.body_elements
        tree.start('w:r')

        tree.start('w:rPr')
        if font is not None:
            tree.add(E('w:rFonts', {'w:eastAsia': font}))
        if size is not None:
            tree.add(E('w:sz', {'w:val': f'{TEXT_SIZE[size]}'}))
        if italic:
            tree.add(E('w:i'))
        if bold:
            tree.add(E('w:b'))
        if underline:
            pass

        tree.end('w:rPr')

        tree.add(E('w:t', text))
        tree.end('w:r')

    def visit_inline_math(self, content):
        pass

    def visit_inline_figure(self, content):
        pass

    def visit_bookmark(self, type_, bookmark):
        number, text = self._retrieve_mark_id(bookmark)

        tree = self.body_elements

        tree.add(E('w:bookmarkStart', {'w:id': number, 'w:name': text}))
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
        tree.add(E('w:bookmarkEnd', {'w:id': number}))

    def visit_reference(self, bookmark):
        mark_id = self._retrieve_mark_id(bookmark)

        tree = self.body_elements
        tree.add(E('w:r', E('w:t', '(')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
        tree.add(E('w:r', E('w:instrText', f'REF {mark_id.text} \\h')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
        tree.add(E('w:r', E('w:t', '0')))
        tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        tree.add(E('w:r', E('w:t', ')')))

    def visit_footnote(self, content):
        number, _ = self._get_footnote_id()

        self.body_elements.add(E('w:r',
                                 E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                                 E('w:footnoteReference', {'w:id': number})))

        tree = self.footnotes_elements

        tree.start('w:footnote', {'w:id': number})

        tree.start('w:p')
        tree.add(E('w:pPr', E('w:pStyle', {'w:val': 'a9'})))
        tree.add(E('w:r',
                   E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                   E('w:footnoteRef')))
        tree.add(E('w:r', {'xml:space': 'preserve'}, ' '))

        content.visit(self)

        tree.end('w:p')

        tree.end('w:footnote')

    def visit_heading(self, content, level, heading):
        number, text = self._retrieve_mark_id(heading)

        tree = self.body_elements
        tree.add(E('w:bookmarkStart', {'w:id': number, 'w:name': text}))
        tree.add(E('w:r', E('w:t', content)))
        tree.add(E('w:bookmarkEnd', {'w:id': number}))

        tree = self.catalog_elements
        prop = E('w:pPr',
                 E('w:pStyle', {'w:val': f'{level*10}'}),
                 E('w:tabs',
                   E('w:tab', {'w:val': 'left', 'w:pos': f'{420+630*(level-1)}'}),
                   E('w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': '8296'})),
                 E('w:rPr', E('w:noProof'), E('w:rStyle', {'w:val': 'a3'})))
        link = E('w:hyperlink', {'w:anchor': text, 'w:history': '1'},
                 E('w:r', E('w:t', '0')),
                 E('w:r', E('w:tab')),
                 E('w:r', E('w:t', content)),
                 E('w:r', E('w:tab')),
                 E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})),
                 E('w:r', E('w:instrText', {'xml:space': 'preserve'},
                            f' PAGEREF {text} \\h ')),
                 E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})),
                 E('w:r', E('w:t', '0')),
                 E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        if tree.is_empty():
            tree.add(E('w:p',
                       E('w:pPr', E('w:jc', {'w:val': 'center'})),
                       E('w:r',
                         E('w:rPr', E('w:b'), E('w:sz', {'w:val': '32'})),
                         E('w:t', '目录'))))
            tree.start('w:p')
            tree.add(prop)
            tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
            tree.add(E('w:r', E('w:instrText', {'xml:space': 'preserve'}, ' ')))
            tree.add(E('w:r', E('w:instrText', 'TOC \\o "1-3" \\h \\z \\u')))
            tree.add(E('w:r', E('w:instrText', ' ', {'xml:space': 'preserve'})))
            tree.add(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
            tree.add(link)
            tree.end('w:p')
        else:
            tree.start('w:p')
            tree.add(prop)
            tree.add(link)
            tree.end('w:p')

    def visit_paragraph(self, content):
        tree = self.body_elements
        tree.start('w:p')
        content.visit(self)
        tree.end('w:p')

    def visit_math(self, content):
        pass

    def visit_figure(self, figure, format_, width, height, title):
        pass

    def visit_table(self, content_by_rows, title):
        pass

    def _retrieve_mark_id(self, obj):
        if id(obj) in self.mark_id_dict.keys():
            return self.mark_id_dict[id(obj)]
        else:
            mark_id = self._get_mark_id()
            self.mark_id_dict[id(obj)] = mark_id
            return mark_id

    def _make_page_break(self):
        return E('w:p', E('w:r', E('w:br', {'w:type': 'page'})))

    def _make_w_sectPr(self, header_rel_id=None, footer_rel_id=None, page_start=None):
        sect_pr = E('w:sectPr')

        if header_rel_id is not None:
            sect_pr.append(E('w:headerReference', {'w:type': 'default', 'r:id': header_rel_id}))

        if footer_rel_id is not None:
            sect_pr.append(E('w:footerReference', {'w:type': 'default', 'r:id': footer_rel_id}))
        sect_pr.append(E('w:pgSz', {'w:w': '11906', 'w:h': '16838'}))
        sect_pr.append(E('w:pgMar', {'w:top': '1440', 'w:right': '1800',
                              'w:bottom': '1440', 'w:left': '1800',
                              'w:footer': '992', 'w:gutter': '0'}))
        if page_start is not None:
            sect_pr.append(E('w:pgNumType', {'w:start': f'{page_start}'}))
        sect_pr.append(E('w:cols', {'w:space': '425'}))
        sect_pr.append(E('w:docGrid', {'w:type': 'lines', 'w:linePitch': '312'}))
        return sect_pr


    # ---------------------------------------------------------------
    # document.xml END
    # ===============================================================


class _Catalog:
    pass










