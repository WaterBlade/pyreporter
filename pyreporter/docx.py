from xml.etree.ElementTree import tostring, Element
from zipfile import ZipFile, ZIP_DEFLATED
from collections import namedtuple
import math

__all__ = ['DocX']

DataFile = namedtuple('DataFile', 'path xml')
FigureFile = namedtuple('FigureFile', 'path figure')
Id = namedtuple('Id', 'number text')

TEXT_SIZE = {'normal': 12, 'large': 18, 'small': 8}
PAGE_WIDTH = 8200
MIN_TABLE_CELL_WIDTH = 1600


# visitor will return list of element.
# not all visitor will return.
# only the Content class visitor will return.


def make_element(tag, *items) -> Element:
    ele = Element(tag)
    last_ele = None
    for item in items:
        assert not isinstance(item, bool)
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
        self.catalog_list = list()

        self.rel_id = 1
        self.footnote_id = 2
        self.mark_id = 1
        self.mark_id_dict = dict()

        self.header_rel_id = None
        self.footer_rel_id = None

        self.doc_rels_elements = list()

        self.cover_elements = list()
        self.catalog_elements = list()
        self.body_elements = list()

        self.footnotes_elements = list()
        self.header_elements = list()

    def set_cover(self, cover):
        builder = Cover.get_cover_builder(cover.type_)
        cover.visit(builder)
        self.cover_elements = builder.elements

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
        xml = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"""
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
        self._write_relationship(self.footer_rel_id,
                                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                 "footer1.xml")

        if self.header_rel_id is not None:
            self._write_relationship(self.header_rel_id,
                                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                                     "header1.xml")

        ns = {'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        root = E('Relationships', ns)
        root.extend(self.doc_rels_elements)
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

        if len(self.cover_elements) > 0:
            body.extend(self.cover_elements)
            body.append(self._make_page_break())

        if len(self.catalog_elements) > 0:
            self._write_catalog_end()
            body.extend(self.catalog_elements)

        body.extend(self.body_elements)
        body.append(self._make_w_sectPr(self.header_rel_id,
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

        footnotes.extend(self.footnotes_elements)

        xml = tostring(footnotes, encoding='unicode')
        self._add_xml(path='word/footnotes.xml', xml=xml)

    def _build_header(self):
        if self.header_rel_id is not None:
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
            hdr.extend(self.header_elements)
            xml = tostring(hdr, encoding='unicode')
            self._add_xml(path='word/header1.xml', xml=xml)

    def _build_numbering(self):
        xml = """<w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="5D861EAE"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="04090025"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="3"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="4"/><w:lvlText w:val="%1.%2.%3.%4"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="864" w:hanging="864"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="5"/><w:lvlText w:val="%1.%2.%3.%4.%5"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1008" w:hanging="1008"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="6"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1152" w:hanging="1152"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="7"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1296" w:hanging="1296"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="8"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="1440"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="9"/><w:lvlText w:val="%1.%2.%3.%4.%5.%6.%7.%8.%9"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1584" w:hanging="1584"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""
        self._add_xml(path='word/numbering.xml', xml=xml)

    def _build_settings(self):
        xml = """<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="160"/><w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/><w:defaultTabStop w:val="420"/><w:drawingGridVerticalSpacing w:val="156"/><w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="2"/><w:characterSpacingControl w:val="compressPunctuation"/><w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr><w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="1"/></w:endnotePr><w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/></w:compat><w:rsids><w:rsidRoot w:val="003221F4"/><w:rsid w:val="0023729F"/><w:rsid w:val="003221F4"/><w:rsid w:val="00581F8C"/><w:rsid w:val="005B06A3"/><w:rsid w:val="00626EC2"/><w:rsid w:val="006A65DB"/><w:rsid w:val="006E593F"/><w:rsid w:val="00926F4A"/><w:rsid w:val="00971FF4"/><w:rsid w:val="00A805B6"/><w:rsid w:val="00C501FC"/><w:rsid w:val="00D637A7"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="off"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="600"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="10242"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>"""
        self._add_xml(path='word/settings.xml', xml=xml)

    def _build_styles(self):
        xml = """<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="00971FF4"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="2Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="3Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="4"><w:name w:val="heading 4"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="4Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="3"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="5"><w:name w:val="heading 5"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="5Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="4"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="6"><w:name w:val="heading 6"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="6Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="5"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="7"><w:name w:val="heading 7"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="7Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="6"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="6"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="8"><w:name w:val="heading 8"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="8Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="7"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="7"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="9"><w:name w:val="heading 9"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="9Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val="003221F4"/><w:pPr><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="8"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/><w:outlineLvl w:val="8"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="a2"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="character" w:customStyle="1" w:styleId="1Char"><w:name w:val="标题 1 Char"/><w:basedOn w:val="a0"/><w:link w:val="1"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="2Char"><w:name w:val="标题 2 Char"/><w:basedOn w:val="a0"/><w:link w:val="2"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="3Char"><w:name w:val="标题 3 Char"/><w:basedOn w:val="a0"/><w:link w:val="3"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="4Char"><w:name w:val="标题 4 Char"/><w:basedOn w:val="a0"/><w:link w:val="4"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="5Char"><w:name w:val="标题 5 Char"/><w:basedOn w:val="a0"/><w:link w:val="5"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="6Char"><w:name w:val="标题 6 Char"/><w:basedOn w:val="a0"/><w:link w:val="6"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="7Char"><w:name w:val="标题 7 Char"/><w:basedOn w:val="a0"/><w:link w:val="7"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="8Char"><w:name w:val="标题 8 Char"/><w:basedOn w:val="a0"/><w:link w:val="8"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="9Char"><w:name w:val="标题 9 Char"/><w:basedOn w:val="a0"/><w:link w:val="9"/><w:uiPriority w:val="9"/><w:rsid w:val="003221F4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:szCs w:val="21"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="10"><w:name w:val="toc 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/></w:style><w:style w:type="paragraph" w:styleId="20"><w:name w:val="toc 2"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="200" w:left="420"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="30"><w:name w:val="toc 3"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style><w:style w:type="character" w:styleId="a3"><w:name w:val="Hyperlink"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:rPr><w:color w:val="0000FF" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a4"><w:name w:val="header"/><w:basedOn w:val="a"/><w:link w:val="Char"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char"><w:name w:val="页眉 Char"/><w:basedOn w:val="a0"/><w:link w:val="a4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a5"><w:name w:val="footer"/><w:basedOn w:val="a"/><w:link w:val="Char0"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="003221F4"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char0"><w:name w:val="页脚 Char"/><w:basedOn w:val="a0"/><w:link w:val="a5"/><w:uiPriority w:val="99"/><w:rsid w:val="003221F4"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a6"><w:name w:val="Title"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="Char1"/><w:uiPriority w:val="10"/><w:qFormat/><w:rsid w:val="00C501FC"/><w:pPr><w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char1"><w:name w:val="标题 Char"/><w:basedOn w:val="a0"/><w:link w:val="a6"/><w:uiPriority w:val="10"/><w:rsid w:val="00C501FC"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="宋体" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a7"><w:name w:val="Document Map"/><w:basedOn w:val="a"/><w:link w:val="Char2"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char2"><w:name w:val="文档结构图 Char"/><w:basedOn w:val="a0"/><w:link w:val="a7"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a8"><w:name w:val="Balloon Text"/><w:basedOn w:val="a"/><w:link w:val="Char3"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char3"><w:name w:val="批注框文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a8"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00581F8C"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a9"><w:name w:val="footnote text"/><w:basedOn w:val="a"/><w:link w:val="Char4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:pPr><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char4"><w:name w:val="脚注文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a9"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="006E593F"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="aa"><w:name w:val="footnote reference"/><w:basedOn w:val="a0"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="006E593F"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="table" w:styleId="ab"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/><w:uiPriority w:val="59"/><w:rsid w:val="00D06FBB"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000" w:themeColor="text1"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"""
        self._add_xml(path='word/styles.xml', xml=xml)

    def _build_webSettings(self):
        xml = """<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>"""
        self._add_xml(path='word/webSettings.xml', xml=xml)

    # ===============================================================
    # document.xml.rels START
    # ---------------------------------------------------------------

    def _write_relationship(self, id_, type_, target):
        self.doc_rels_elements.append(E('Relationship', {'Id': id_, 'Type': type_, 'Target': target}))

    # ---------------------------------------------------------------
    # document.xml.rels END
    # ===============================================================

    # ===============================================================
    # header1.xml START
    # ---------------------------------------------------------------

    def visit_header(self, header):
        tree = self.header_elements
        tree.append(E('w:p',
                      E('w:pPr',
                        E('w:pStyle', {'w:val': 'a4'})),
                      E('w:r',
                        E('w:rPr', E('w:rFonts', {'w:eastAsia': '楷体'})),
                        E('w:t', header))))
        self.header_rel_id = self._get_rel_id().text

    # ---------------------------------------------------------------
    # header1.xml END
    # ===============================================================

    # ===============================================================
    # document.xml START
    # ---------------------------------------------------------------

    def visit_composite(self, list_):
        ret = list()
        for ele in list_:
            ret.extend(ele.visit(self))
        return ret

    def visit_text(self, text, font, size, italic, bold, underline):
        run = E('w:r')

        if font or size or italic or bold or underline:
            prop = E('w:rPr')
            run.append(prop)
            if font is not None:
                prop.append(E('w:rFonts', {'w:eastAsia': font}))
            if size is not None:
                prop.append(E('w:sz', {'w:val': f'{TEXT_SIZE[size]}'}))
            if italic:
                prop.append(E('w:i'))
            if bold:
                prop.append(E('w:b'))
            if underline:
                pass

        run.append(E('w:t', text))
        return [run]

    def visit_inline_math(self, content):
        m = E('m:oMath')
        m.extend(content.visit(self))
        return [m]

    def visit_inline_figure(self, *, figure, format_, width, height):
        width = int(width * 914400)
        height = int(height * 914400)

        fig_id = len(self.figure_list) + 1
        rel_id = self._get_rel_id().text

        self.figure_list.append(FigureFile(path=f'word/media/image{fig_id}.{format_}', figure=figure))
        self._write_relationship(id_=rel_id,
                                 type_='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                 target=f'media/image{fig_id}.{format_}')

        run = E('w:r')

        draw = E('w:drawing')
        run.append(draw)

        inline = E('wp:inline')
        draw.append(inline)

        inline.append(E('wp:extent', {'cx': f'{width}', 'cy': f'{height}'}))
        inline.append(E('wp:effectExtent', {'l': '0', 't': '0', 'r': '0', 'b': '0'}))
        inline.append(E('wp:docPr', {'id': f'{fig_id}', 'name': f'image{fig_id}'}))

        inline.append(E('wp:cNvGraphicFramePr',
                        E('a:graphicFrameLocks',
                          {'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main", 'noChangeAspect': '1'})))

        inline.append(self._make_a_graphic(fig_id, rel_id, width, height))

        return [run]

    def visit_bookmark(self, *, type_, bookmark, left, right):
        number, text = self._retrieve_mark_id(bookmark)

        ret = list()

        if left is not None:
            ret.append(E('w:r', E('w:t', left)))
        ret.append(E('w:bookmarkStart', {'w:id': number, 'w:name': text}))
        ret.append(E('w:r', E('w:t', type_)))
        ret.append(E('w:r', E('w:t', '', {'xml:space': 'preserve'})))
        ret.append(E('w:fldSimple',
                     E('w:r', E('w:t', '0')), {'w:instr': r' STYLEREF 1 \s'}))
        ret.append(E('w:r', E('w:noBreakHyphen')))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
        ret.append(E('w:r',
                     E('w:instrText', ' ', {'xml:space': 'preserve'})))
        ret.append(E('w:r',
                     E('w:instrText', 'SEQ ', {'xml:space': 'preserve'})))
        ret.append(E('w:r', E('w:instrText', type_)))
        ret.append(E('w:r',
                     E('w:instrText', r' \* ARABIC \s 1', {'xml:space': 'preserve'})))
        ret.append(E('w:r', E('w:instrText', '', {'xml:space': 'preserve'})))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
        ret.append(E('w:r', E('w:t', '1')))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        ret.append(E('w:bookmarkEnd', {'w:id': number}))
        if right is not None:
            ret.append(E('w:r', E('w:t', right)))

        return ret

    def visit_reference(self, bookmark):
        mark_id = self._retrieve_mark_id(bookmark)

        ret = list()
        ret.append(E('w:r', E('w:t', '(')))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
        ret.append(E('w:r', E('w:instrText', f'REF {mark_id.text} \\h')))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))
        ret.append(E('w:r', E('w:t', '0')))
        ret.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
        ret.append(E('w:r', E('w:t', ')')))

        return ret

    def visit_footnote(self, content):
        number, _ = self._get_footnote_id()

        run = E('w:r',
                E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                E('w:footnoteReference', {'w:id': number}))

        foot = E('w:footnote', {'w:id': number})
        self.footnotes_elements.append(foot)

        para = E('w:p')
        foot.append(para)

        para.append(E('w:pPr', E('w:pStyle', {'w:val': 'a9'})))
        para.append(E('w:r',
                      E('w:rPr', E('w:rStyle', {'w:val': 'aa'})),
                      E('w:footnoteRef')))
        para.append(E('w:r', {'xml:space': 'preserve'}, ' '))

        para.extend(content.visit(self))

        return [run]

    def visit_heading(self, *, content, level, heading):
        number, text = self._retrieve_mark_id(heading)

        body = self.body_elements
        if len(self.catalog_elements) > 0 and level == 1:
            body.append(self._make_page_break())

        para = E('w:p')
        body.append(para)

        para.append(E('w:pPr', E('w:pStyle', {'w:val': f'{level}'})))
        para.append(E('w:bookmarkStart', {'w:id': number, 'w:name': text}))
        para.extend(content.visit(self))
        para.append(E('w:bookmarkEnd', {'w:id': number}))

        log = self.catalog_elements

        if level <= 3:

            para = E('w:p')

            prop = E('w:pPr',
                     E('w:pStyle', {'w:val': f'{level*10}'}),
                     E('w:tabs',
                       E('w:tab', {'w:val': 'left', 'w:pos': f'{420+630*(level-1)}'}),
                       E('w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': '8296'})),
                     E('w:rPr', E('w:noProof'), E('w:rStyle', {'w:val': 'a3'})))
            para.append(prop)

            if len(log) == 0:
                log.append(E('w:p',
                             E('w:pPr', E('w:jc', {'w:val': 'center'})),
                             E('w:r',
                               E('w:rPr', E('w:b'), E('w:sz', {'w:val': '32'})),
                               E('w:t', '目录'))))
                log.append(para)

                para.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})))
                para.append(E('w:r', E('w:instrText', {'xml:space': 'preserve'}, ' ')))
                para.append(E('w:r', E('w:instrText', 'TOC \\o "1-3" \\h \\z \\u')))
                para.append(E('w:r', E('w:instrText', ' ', {'xml:space': 'preserve'})))
                para.append(E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})))

            else:
                log.append(para)

            link = E('w:hyperlink', {'w:anchor': text, 'w:history': '1'},
                     E('w:r', E('w:t', self._get_catalog_number(level))),
                     E('w:r', E('w:tab')),
                     content.visit(self),
                     E('w:r', E('w:tab')),
                     E('w:r', E('w:fldChar', {'w:fldCharType': 'begin'})),
                     E('w:r', E('w:instrText', {'xml:space': 'preserve'},
                                f' PAGEREF {text} \\h ')),
                     E('w:r', E('w:fldChar', {'w:fldCharType': 'separate'})),
                     E('w:r', E('w:t', '0')),
                     E('w:r', E('w:fldChar', {'w:fldCharType': 'end'})))
            para.append(link)

    def visit_paragraph(self, content):
        para = E('w:p')
        self.body_elements.append(para)

        para.append(E('w:pPr',
                      E('w:spacing', {'w:before': '156', 'w:after': '156'}),
                      E('w:ind', {'w:firstLine': '420'})))

        para.extend(content.visit(self))

    def visit_standalone_math(self, content):
        para = E('w:p')
        self.body_elements.append(para)

        m_p = E('m:oMathPara')
        para.append(m_p)

        for i in range(len(content)):
            m = content[i].visit(self)[0]
            m_p.append(m)
            if i != len(content) - 1:
                m_p.append(E('w:r', E('w:br')))

    def visit_math_definition(self, content, bookmark):
        m_p = E('m:oMathPara')
        for i in range(len(content)):
            m = content[i].visit(self)[0]
            m_p.append(m)
            if i != len(content) - 1:
                m_p.append(E('w:r', E('w:br')))
        tb = E('w:tbl',
               E('w:tblPr',
                 E('w:jc', {'w:val': 'center'}),
                 E('w:tblLayout', {'w:type': 'fixed'})),
               E('w:tblGrid',
                 E('w:gridCol', {'w:w': '8000'}),
                 E('w:gridCol', {'w:w': '1000'})),
               E('w:tr',
                 E('w:trPr', E('w:cantSplit'), E('w:jc', {'w:val': 'center'})),
                 E('w:tc',
                   E('w:tcPr', E('w:vAlign', {'w:val': 'center'})),
                   E('w:p',
                     E('w:pPr', E('w:jc', {'w:val': 'center'})),
                     m_p)),
                 E('w:tc',
                   E('w:tcPr', E('w:vAlign', {'w:val': 'center'})),
                   E('w:p',
                     E('w:pPr', E('w:jc', {'w:val': 'center'})),
                     bookmark.visit(self)))
                 ))
        self.body_elements.append(tb)

    def visit_math_procedure(self, content):
        p = E('w:p')
        m_p = E('m:oMathPara')
        p.append(m_p)
        for i in range(len(content)):
            m = content[i].visit(self)[0]
            m_p.append(m)
            if i != len(content) - 1:
                m_p.append(E('w:r', E('w:br')))
        self.body_elements.append(p)

    def visit_math_note(self, var_list):
        if len(var_list) > 0:
            self.body_elements.append(E('w:p', E('w:r', E('w:t', '式中：'))))
        for var in var_list:
            p = E('w:p')
            self.body_elements.append(p)

            p.append(E('w:pPr',
                       E('w:tabs',
                         E('w:tab', {'w:val': 'right', 'w:pos': '500'}),
                         E('w:tab', {'w:val': 'center', 'w:pos': '600'}),
                         E('w:tab', {'w:val': 'left', 'w:pos': '700'}))))
            p.append(E('w:r', E('w:tab')))
            p.append(E('m:oMath',
                       var.visit(self)))
            p.append(E('w:r', E('w:tab')))
            p.append(E('w:r', E('w:t', '―')))
            p.append(E('w:r', E('w:tab')))
            p.append(E('w:r', E('w:t', var.inform)))
            if var.unit is not None:
                p.append(E('w:r', E('w:t', '，')))
                p.append(E('m:oMath',
                           var.unit.visit(self)))

    def visit_figure(self, figure, format_, width, height, title):
        para = E('w:p')
        self.body_elements.append(para)

        prop = E('w:pPr')
        para.append(prop)

        prop.append(E('w:jc', {'w:val': 'center'}))
        if title is not None:
            prop.append(E('w:keepNext'))

        para.extend(self.visit_inline_figure(figure=figure, format_=format_, width=width, height=height))

        if title is not None:
            p = E('w:p')
            self.body_elements.append(p)
            p.append(E('w:pPr', E('w:pStyle', {'w:val': 'a9'})))
            p.extend(title.visit(self))

    def visit_table(self, content_by_rows, title):
        width = min(PAGE_WIDTH // len(content_by_rows[0]),
                    MIN_TABLE_CELL_WIDTH)

        if title is not None:
            p = E('w:p')
            self.body_elements.append(p)
            p.append(E('w:pPr',
                       E('w:pStyle', {'w:val': 'a9'}),
                       E('w:jc', {'w:val': 'center'}),
                       E('w:keepNext')))
            p.extend(title.visit(self))

        tb = E('w:tbl')
        self.body_elements.append(tb)

        tb_pr = E('w:tblPr')
        tb.append(tb_pr)
        tb_pr.append(E('w:tblStyle', {'w:val': 'ab'}))
        tb_pr.append(E('w:jc', {'w:val': 'center'}))
        tb_pr.append(E('w:tblLayout', {'w:type': 'fixed'}))

        grid = E('w:tblGrid')
        tb.append(grid)
        for i in range(len(content_by_rows[0])):
            grid.append(E('w:gridCol', {'w:w': f'{width}'}))

        for row in content_by_rows:
            tr = E('w:tr')
            tb.append(tr)

            for cell in row:
                tc = E('w:tc')
                tr.append(tc)
                tc.append(E('w:tcPr',
                            E('w:vAlign', {'w:val': 'center'})))

                p = E('w:p')
                tc.append(p)
                p.append(E('w:pPr',
                           E('w:jc', {'w:val': 'center'})))

                p.extend(cell.visit(self))

    def _retrieve_mark_id(self, obj):
        if id(obj) in self.mark_id_dict.keys():
            return self.mark_id_dict[id(obj)]
        else:
            mark_id = self._get_mark_id()
            self.mark_id_dict[id(obj)] = mark_id
            return mark_id

    def _write_catalog_end(self):
        log = self.catalog_elements
        log.append(E('w:p', E('w:r', E('w:fldChar', {'w:fldCharType': 'end'}))))
        log.append(E('w:p', E('w:pPr', self._make_w_sectPr())))

    def _get_catalog_number(self, level):
        if level > len(self.catalog_list):
            self.catalog_list.append(1)
        else:
            self.catalog_list = self.catalog_list[:level]
            self.catalog_list[-1] += 1
        return '.'.join(str(item) for item in self.catalog_list)

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

    def _make_a_graphic(self, fig_id, rel_id, cx, cy):
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
        return ele

    # ---------------------------------------------------------------
    # document.xml END
    # ===============================================================

    # ===============================================================
    # math START
    # ---------------------------------------------------------------

    def visit_negative(self, exp):
        return [self._make_m_r('-', sty='p'),
                *exp.visit(self)]

    def visit_add(self, left, right):
        return [*left.visit(self),
                self._make_m_r('+', sty='p'),
                *right.visit(self)]

    def visit_sub(self, left, right):
        return [*left.visit(self),
                self._make_m_r('-', sty='p'),
                *right.visit(self)]

    def visit_mul(self, left, right):
        return [*left.visit(self),
                self._make_m_r('⋅', sty='p'),
                *right.visit(self)]

    def visit_div(self, left, right, type_='bar'):
        f = E('m:f')
        f.append(E('m:fPr',
                   E('m:type', {'m:val': type_})))
        f.append(E('m:num', left.visit(self)))
        f.append(E('m:den', right.visit(self)))
        return [f]

    def visit_flat_div(self, left, right):
        return self.visit_div(left, right, type_='lin')

    def visit_pow(self, exp, index):
        ret_exp = exp.visit(self)
        if ret_exp[0].tag == 'm:sSub':
            ret_exp[0].tag = 'm:sSubSup'
            ret_exp[0].append(E('m:sup', index.visit(self)))
            return ret_exp
        return [self._make_m_sSup(exp, index)]

    def visit_radical(self, exp, index):
        rad = E('m:rad')

        pr = E('m:radPr')
        rad.append(pr)
        if index.value == 2:
            pr.append(E('m:degHide', {'m:val': 'on'}))

        if index.value != 2:
            rad.append(E('m:deg', index.visit(self)))
        rad.append(E('m:e', exp.visit(self)))

        return [rad]

    def visit_lesser_than(self, left, right):
        return [*left.visit(self),
                self._make_m_r('<', sty='p'),
                *right.visit(self)]

    def visit_lesser_or_equal(self, left, right):
        return [*left.visit(self),
                self._make_m_r('≤', sty='p'),
                *right.visit(self)]

    def visit_equal(self, left, right):
        return [*left.visit(self),
                self._make_m_r('=', sty='p'),
                *right.visit(self)]

    def visit_not_equal(self, left, right):
        return [*left.visit(self),
                self._make_m_r('≠', sty='p'),
                *right.visit(self)]

    def visit_greater_than(self, left, right):
        return [*left.visit(self),
                self._make_m_r('>', sty='p'),
                *right.visit(self)]

    def visit_greater_or_equal(self, left, right):
        return [*left.visit(self),
                self._make_m_r('≥', sty='p'),
                *right.visit(self)]

    def visit_sin(self, exp):
        return [self._make_m_func('sin', exp)]

    def visit_cos(self, exp):
        return [self._make_m_func('cos', exp)]

    def visit_tan(self, exp):
        return [self._make_m_func('tan', exp)]

    def visit_cot(self, exp):
        return [self._make_m_func('cot', exp)]

    def visit_arcsin(self, exp):
        return [self._make_m_func('arcsin', exp)]

    def visit_arccos(self, exp):
        return [self._make_m_func('arccos', exp)]

    def visit_arctan(self, exp):
        return [self._make_m_func('arctan', exp)]

    def visit_arccot(self, exp):
        return [self._make_m_func('arccot', exp)]

    def visit_parenthesis(self, exp):
        return [self._make_m_d(exp, left='(', right=')')]

    def visit_square_bracket(self, exp):
        return [self._make_m_d(exp, left='[', right=']')]

    def visit_brace(self, exp):
        return [self._make_m_d(exp, left='{', right='}')]

    def visit_variable(self, var, sub):
        if sub is None:
            return [self._make_m_r(var)]
        else:
            return [self._make_m_sSub(var, sub)]

    def visit_number(self, value, precision):
        if abs(value) < 1e-10:
            ret = [self._make_m_r('0')]
        elif abs(value) > 10000 or abs(value) < 0.001 and value != 0:
            sup = math.floor(math.log10(abs(value)))
            base = value / math.pow(10, sup)
            ret = [self._make_m_r(f'{base:.2f}'),
                   self._make_m_r('⋅', sty='p'),
                   self._make_m_sSup('10', f'{sup}')]
        else:
            if precision is None:
                value = f'{value}'
            elif precision == 'auto':
                if isinstance(value, int):
                    value = f'{value}'
                elif abs(value) > 1:
                    value = f'{value:.3f}'
                else:
                    precision = abs(math.floor(math.log10(abs(value)))) + 2
                    value = f'{value:.{precision}f}'
            else:
                value = f'{value:.{precision}f}'
            ret = [self._make_m_r(value)]

        return ret

    def visit_unit(self, symbol):
        return [self._make_m_r(symbol, sty='p')]

    def visit_sum(self, exp):
        return [self._make_m_nary(exp, name='∑')]

    def visit_serial_variable(self, var, sub, index):
        if sub is None:
            return [self._make_m_sSub(var, f'{index}')]
        elif isinstance(sub, str):
            return [self._make_m_sSub(var, sub + f'-{index}')]
        else:
            E('m:sSub',
              E('m:e', self._make_m_r(var)),
              E('m:sub', sub.visit(self), self._make_m_r(f'-{index}')))

    def visit_math_text(self, text, align, sty, color):
        return [self._make_m_r(text, align=align, sty=sty, color=color)]

    def visit_multi_line(self, list_, included):
        arr = E('m:eqArr')
        for item in list_:
            arr.append(E('m:e', item.visit(self)))
        if included is None:
            return [arr]
        elif included == 'left':
            return [E('m:d',
                      E('m:dPr',
                        E('m:begChr', {'m:val': '{'}),
                        E('m:endChr', {'m:val': ''})),
                      E('m:e',
                        arr))]
        elif included == 'right':
            return [E('m:d',
                      E('m:dPr',
                        E('m:begChr', {'m:val': ''}),
                        E('m:endChr', {'m:val': '}'})),
                      E('m:e',
                        arr))]
        else:
            raise TypeError('Unknown included type: %s in math multi line' % included)

    def _make_m_r(self, text, *, sty=None, align=False, color=None):
        run = E('m:r')
        if sty or align:
            pr = E('m:rPr')
            run.append(pr)
            if sty:
                pr.append(E('m:sty', {'m:val': sty}))
            if align:
                pr.append(E('m:aln'))
        wpr = E('w:rPr')
        run.append(wpr)

        wpr.append(E('w:rFonts', {'w:ascii': 'Cambria Math',
                                  'w:hAnsi': 'Cambria Math'}))
        if color is not None:
            wpr.append(E('w:color', {'w:val': color}))

        run.append(E('m:t', text))
        return run

    def _make_m_sSub(self, base, sub):
        base = self._make_m_r(base) if isinstance(base, str) else base.visit(self)
        sub = self._make_m_r(sub) if isinstance(sub, str) else sub.visit(self)
        return E('m:sSub',
                 E('m:e', base),
                 E('m:sub', sub))

    def _make_m_sSup(self, base, sup):
        if isinstance(base, str):
            base = self._make_m_r(base)
        else:
            base = base.visit(self)
            if base[-1].tag == 'm:sSup':
                base = E('m:d',
                         E('m:dPr',
                           E('m:begChr', {'m:val': '('}),
                           E('m:endChr', {'m:val': ')'})),
                         E('m:e',
                           base))
        sup = self._make_m_r(sup) if isinstance(sup, str) else sup.visit(self)
        return E('m:sSup',
                 E('m:e', base),
                 E('m:sup', sup))

    def _make_m_func(self, name, exp):
        return E('m:func',
                 E('m:fName', self._make_m_r(name, sty='p')),
                 E('m:e', exp.visit(self)))

    def _make_m_d(self, *exps, left=None, right=None):
        d = E('m:d')

        pr = E('m:dPr')
        d.append(pr)

        if left is not None:
            pr.append(E('m:begChr', {'m:val': left}))
        if right is not None:
            pr.append(E('m:endChr', {'m:val': right}))

        for exp in exps:
            d.append(E('m:e', exp.visit(self)))

        return d

    def _make_m_nary(self, exp, sub=None, sup=None, name=None, limLoc='subSup'):
        n = E('m:nary')

        pr = E('m:naryPr')
        n.append(pr)

        if name:
            pr.append(E('m:chr', {'m:val': name}))
        if sub is None:
            pr.append(E('m:subHide', {'m:val': 'on'}))
        if sup is None:
            pr.append(E('m:supHide', {'m:val': 'on'}))
        pr.append(E('m:ctrlPr', E('w:rPr', E('w:i'))))

        if sub is not None:
            n.append(E('m:sub', sub.visit(self)))
        if sup is not None:
            n.append(E('m:sup', sup.visit(self)))
        n.append(E('m:e', exp.visit(self)))

        return n

    # ---------------------------------------------------------------
    # math END
    # ===============================================================


class Cover:
    @classmethod
    def get_cover_builder(cls, type_):
        if type_ == 'default cover':
            return DefaultCover()


class DefaultCover(Cover):
    def __init__(self):
        self.elements = list()

    def visit_cover(self, *, project, name, part, phase, number, secret, footer_str):
        tree = self.elements

        c1 = self._make_enclosed_cell('秘密')
        c2 = self._make_enclosed_cell(secret)
        tree.append(self._make_table([[c1, c2]],
                                     grid_col=[600, 2000],
                                     x_spec='left', y_spec='top',
                                     top_margin=0, bottom_margin=0,
                                     left_margin=0, right_margin=0))

        c1 = self._make_enclosed_cell('编号')
        c2 = self._make_enclosed_cell(number)
        tree.append(self._make_table([[c1, c2]],
                                     grid_col=[600, 2000],
                                     x_spec='right', y_spec='top',
                                     top_margin=0, bottom_margin=0,
                                     left_margin=0, right_margin=0))

        c1 = self._make_cell(text='计 算 书', size=72)
        tree.append(self._make_table([[c1]],
                                     grid_col=[5000],
                                     x_spec='center', y=3000))

        rows = list()
        for left, right in zip(['工程名称', '专业名称', '设计阶段', '计算书名称'],
                               [project, part, phase, name]):
            c1 = self._make_cell(text=left, size=30,
                                 h_align='right', v_align='bottom',
                                 space=(200, 0))
            c2 = self._make_cell(text=right, size=30, font='楷体',
                                 v_align='bottom',
                                 borders=['bottom'],
                                 b_size=6,
                                 space=(200, 0))
            rows.append([c1, c2])
        tree.append(self._make_table(rows, grid_col=[1600, 5000],
                                     x_spec='center', y=6000,
                                     left_margin=0, right_margin=10))

        rows = list()
        year = self._make_cell('年', size=24, v_align='bottom', space=(120, 0))
        month = self._make_cell('月', size=24, v_align='bottom', space=(120, 0))
        day = self._make_cell('日', size=24, v_align='bottom', space=(120, 0))
        line = self._make_cell(v_align='bottom', size=24,
                               space=(120, 0),
                               borders=['bottom'], b_size=6)
        blk = self._make_cell()
        for txt in ['审查', '校核', '计算']:
            c1 = self._make_cell(txt, size=24, v_align='bottom', space=(120, 0))
            rows.append([c1, line, blk, line, year, line, month, line, day])
        tree.append(self._make_table(rows,
                                     grid_col=[700, 1500, 300, 900, 300, 600, 300, 600, 300],
                                     x_spec='center', y=10500,
                                     left_margin=0, right_margin=0))

        footer = self._make_cell(footer_str, size=24)
        tree.append(self._make_table([[footer]],
                                     grid_col=[5000],
                                     x_spec='center', y=13500))

    def _make_table(self, rows, grid_col=None, x=None, y=None,
                    x_spec=None, y_spec=None,
                    top_margin=None, bottom_margin=None,
                    left_margin=None, right_margin=None):
        tb = E('w:tbl')

        pr = E('w:tblPr')
        tb.append(pr)

        pr.append(E('w:tblLayout', {'w:type': 'fixed'}))

        if x or y or x_spec or y_spec:
            attr = dict()
            attr['w:horizonAnchor'] = 'margin'
            if x:
                attr['w:tblpX'] = f'{x}'
            if x_spec:
                attr['w:tblpXSpec'] = f'{x_spec}'
            attr['w:vertAnchor'] = 'margin'
            if y:
                attr['w:tblpY'] = f'{y}'
            if y_spec:
                attr['w:tblpYSpec'] = f'{y_spec}'
            pr.append(E('w:tblpPr', attr))

        if (top_margin is not None
                or bottom_margin is not None
                or left_margin is not None
                or right_margin is not None):
            mar = E('w:tblCellMar')
            pr.append(mar)
            if top_margin is not None:
                mar.append(E('w:top', {'w:w': f'{top_margin}', 'w:type': 'dxa'}))
            if bottom_margin is not None:
                mar.append(E('w:bottom', {'w:w': f'{bottom_margin}', 'w:type': 'dxa'}))
            if left_margin is not None:
                mar.append(E('w:left', {'w:w': f'{left_margin}', 'w:type': 'dxa'}))
            if right_margin is not None:
                mar.append(E('w:right', {'w:w': f'{right_margin}', 'w:type': 'dxa'}))

        if grid_col is not None:
            grid = E('w:tblGrid')
            tb.append(grid)
            for w in grid_col:
                grid.append(E('w:gridCol', {'w:w': f'{w}'}))

        for r in rows:
            row = E('w:tr')
            tb.append(row)
            for c in r:
                row.append(c)

        return tb

    def _make_enclosed_cell(self, text=None):
        return self._make_cell(text=text, borders=['left', 'right', 'top', 'bottom'])

    def _make_cell(self, text=None, size=None, font=None,
                   h_align='center', v_align='center',
                   borders: list = None, space=(0, 0),
                   b_size=4):
        tc = E('w:tc')

        prop = E('w:tcPr')
        tc.append(prop)

        prop.append(E('w:vAlign', {'w:val': v_align}))

        if borders is not None:
            border = E('w:tcBorders')
            for side in borders:
                border.append(E(f'w:{side}',
                                {'w:val': 'single',
                                 'w:sz': f'{b_size}',
                                 'w:space': '0',
                                 'w:color': 'auto'}))
            prop.append(border)

        para = E('w:p')
        tc.append(para)

        p_pr = E('w:pPr')
        para.append(p_pr)

        p_pr.append(E('w:jc', {'w:val': h_align}))
        if v_align != 'center':
            p_pr.append(E('w:snapToGrid', {'w:val': 'false'}))
            p_pr.append(E('w:spacing',
                          {'w:before': f'{space[0]}', 'w:after': f'{space[1]}'}))

        r = E('w:r')
        para.append(r)

        if size or font:
            r_pr_in_p = E('w:rPr')
            r_pr_in_r = E('w:rPr')
            p_pr.append(r_pr_in_p)
            r.append(r_pr_in_r)

            if size:
                r_pr_in_p.append(E('w:sz', {'w:val': f'{size}'}))
                r_pr_in_r.append(E('w:sz', {'w:val': f'{size}'}))
            if font:
                r_pr_in_p.append(E('w:rFonts', {'w:eastAsia': font}))
                r_pr_in_r.append(E('w:rFonts', {'w:eastAsia': font}))

        if text:
            r.append(E('w:t', text))

        return tc
