# Writes the Word samples for the self-test (tests/samples), as Word itself would store them:
#   complex.docx — title, table of contents field, headings with bookmarks, formatted text, link, footnote, comment,
#                  tracked change, bullet and numbered lists, a table with merged and shaded cells, a picture, a text box,
#                  header and footer with a page-number field, a page break and a landscape last section
#   messy.docx   — the kind of formatting Fix Formatting and Auto Format clean up
# Run: python tests/make_docx_samples.py   (check them with: powershell -File tests/word_check.ps1 tests/samples/complex.docx)
import os, zipfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'tests', 'samples')
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS = (W + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
      ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"')
H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

def esc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def r(text, rpr=''): return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
def p(inner, ppr='', attrs=''): return f'<w:p{attrs}>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{inner}</w:p>'
def style(sid): return f'<w:pStyle w:val="{sid}"/>'

def package(name, body, parts, rels, types, sect):
    doc = f'{H}<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
    files = {
        '[Content_Types].xml': f'{H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
                               f'<Override PartName="/word/document.xml" ContentType="{CT}document.main+xml"/>' + ''.join(f'<Override PartName="/{n}" ContentType="{t}"/>' for n, t in types) +
                               '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
        '_rels/.rels': f'{H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',
        'word/_rels/document.xml.rels': f'{H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(f'<Relationship Id="{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{t}" Target="{g}"{" TargetMode=\"External\"" if ext else ""}/>' for i, t, g, ext in rels) + '</Relationships>',
        'word/document.xml': doc,
        'docProps/core.xml': f'{H}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{name}</dc:title><dc:creator>XD3 tests</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-01T10:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-01T10:00:00Z</dcterms:modified></cp:coreProperties>',
        'docProps/app.xml': f'{H}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>',
        **parts,
    }
    with zipfile.ZipFile(os.path.join(OUT, name), 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items(): z.writestr(n, data if isinstance(data, bytes) else data.encode('utf-8'))

STYLES = f'''{H}<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>'''.replace('\n', '')

def lvl(i, fmt, text, font=None):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ''
    return f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr>{rpr}</w:lvl>'
NUMBERING = (f'{H}<w:numbering {W}>'
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' + ''.join(lvl(i, 'bullet', ['', 'o', ''][i % 3], ['Symbol', 'Courier New', 'Wingdings'][i % 3]) for i in range(9)) + '</w:abstractNum>'
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' + ''.join(lvl(i, ['decimal', 'lowerLetter', 'lowerRoman'][i % 3], f'%{i + 1}.') for i in range(9)) + '</w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>')
SETTINGS = f'{H}<w:settings {W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'

def complex_docx():
    fld = lambda t: f'<w:r><w:fldChar w:fldCharType="{t}"/></w:r>'
    toc_line = lambda anchor, text, style_id: p(f'<w:hyperlink w:anchor="{anchor}" w:history="1">{r(text, "<w:noProof/>")}<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>{r("1", "<w:noProof/><w:webHidden/>")}</w:hyperlink>', style(style_id) + '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>')
    toc = (p(f'{fld("begin")}<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>{fld("separate")}<w:hyperlink w:anchor="_Toc1" w:history="1">{r("Introduction", "<w:noProof/>")}<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>{r("1", "<w:noProof/><w:webHidden/>")}</w:hyperlink>', style('TOC1') + '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>')
           + toc_line('_Toc2', 'Data', 'TOC2') + p(fld('end')))
    intro = (p(r('Project Report'), style('Title'), ' w14:paraId="1A2B3C01" w14:textId="1A2B3C01" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"') + toc
        + p('<w:bookmarkStart w:id="1" w:name="_Toc1"/>' + r('Introduction') + '<w:bookmarkEnd w:id="1"/>', style('Heading1'))
        + p(r('This report has ') + r('bold', '<w:b/><w:bCs/>') + r(', ') + r('italic', '<w:i/><w:iCs/>') + r(', ') + r('red', '<w:color w:val="FF0000"/>') + r(' and ')
            + r('underlined', '<w:u w:val="single"/>') + r(' words. ') + r('Highlighted', '<w:highlight w:val="yellow"/>') + r(' text, a link to ')
            + '<w:hyperlink r:id="rId9" w:history="1">' + r('the website', '<w:rStyle w:val="Hyperlink"/>') + '</w:hyperlink>' + r(' and a footnote')
            + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' + r('. The quick brown fox jumps over the lazy dog.'))
        + p(r('A sentence with a ') + '<w:commentRangeStart w:id="0"/>' + r('comment') + '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>'
            + r(' and a ') + '<w:bookmarkStart w:id="3" w:name="Mark1"/>' + r('bookmark') + '<w:bookmarkEnd w:id="3"/>' + '<w:proofErr w:type="spellStart"/>' + r('.') + '<w:proofErr w:type="spellEnd"/>', '', ' w:rsidR="00B2C3D4"')
        + p('<w:ins w:id="5" w:author="Tester" w:date="2026-10-01T10:00:00Z">' + r('This sentence was added with track changes on.') + '</w:ins>'))
    bullets = ''.join(p(r(t), style('ListParagraph') + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') for t in ['First bullet', 'Second bullet', 'Third bullet'])
    steps = ''.join(p(r(t), style('ListParagraph') + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>') for t in ['Step one', 'Step two', 'Step three'])
    cell = lambda inner, w=3116, extra='': f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{extra}</w:tcPr>{inner}</w:tc>'
    head = ''.join(cell(p(r(t, '<w:b/><w:bCs/>')), extra='<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') for t in ['Item', 'Count', 'Note'])
    table = (f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3116"/><w:gridCol w:w="3117"/><w:gridCol w:w="3117"/></w:tblGrid>'
        f'<w:tr>{head}</w:tr><w:tr>{cell(p(r("Apples")))}{cell(p(r("12")))}{cell(p(r("Fresh")), extra="<w:vMerge w:val=\"restart\"/>")}</w:tr><w:tr>{cell(p(r("Pears")))}{cell(p(r("7")))}{cell(p(""), extra="<w:vMerge/>")}</w:tr></w:tbl>')
    emu = 914400
    picture = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{emu}" cy="{emu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="icon.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu}" cy="{emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    textbox = ('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3600000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2032000" cy="635000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Text Box 2"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2032000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Text in a box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
        '<mc:Fallback><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe" xmlns:o="urn:schemas-microsoft-com:office:office"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
        '<v:shape id="Text Box 2" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:283.45pt;margin-top:0;width:160pt;height:50pt;z-index:251659264" xmlns:o="urn:schemas-microsoft-com:office:office"><v:textbox><w:txbxContent><w:p><w:r><w:t>Text in a box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>')
    hf_refs = '<w:headerReference w:type="default" r:id="rId6"/><w:footerReference w:type="default" r:id="rId7"/>'
    margins = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    body = (intro + bullets + p(r('Between the lists.')) + steps
        + p('<w:bookmarkStart w:id="2" w:name="_Toc2"/>' + r('Data') + '<w:bookmarkEnd w:id="2"/>', style('Heading2')) + table
        + p(r('A picture: ') + picture) + p(r('A paragraph with a text box beside it.') + textbox)
        + p('<w:r><w:br w:type="page"/></w:r>') + p(r('Second page'), style('Heading1'))
        + p(r('Text on the second page. It has enough words to be an ordinary paragraph of body text.'), f'<w:sectPr>{hf_refs}<w:pgSz w:w="12240" w:h="15840"/>{margins}<w:cols w:space="720"/></w:sectPr>')
        + p(r('A landscape section at the end.')))
    sect = f'<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>{margins}<w:cols w:space="720"/></w:sectPr>'
    footnotes = (f'{H}<w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> This is the footnote text.</w:t></w:r></w:p></w:footnote></w:footnotes>')
    comments = f'{H}<w:comments {W}><w:comment w:id="0" w:author="Tester" w:date="2026-10-01T10:00:00Z" w:initials="T"><w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Please check this.</w:t></w:r></w:p></w:comment></w:comments>'
    header = f'{H}<w:hdr {W}>' + p(r('XD3 Test Header'), style('Header')) + '</w:hdr>'
    footer = (f'{H}<w:ftr {W}>' + p(r('Page ') + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        + r('1', '<w:noProof/>') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>', style('Footer') + '<w:jc w:val="center"/>') + '</w:ftr>')
    png = open(os.path.join(ROOT, 'icons', 'icon-192.png'), 'rb').read()
    package('complex.docx', body, {'word/styles.xml': STYLES, 'word/numbering.xml': NUMBERING, 'word/settings.xml': SETTINGS, 'word/footnotes.xml': footnotes, 'word/comments.xml': comments, 'word/header1.xml': header, 'word/footer1.xml': footer, 'word/media/image1.png': png},
        [('rId1', 'styles', 'styles.xml', 0), ('rId2', 'numbering', 'numbering.xml', 0), ('rId3', 'footnotes', 'footnotes.xml', 0), ('rId4', 'comments', 'comments.xml', 0), ('rId5', 'settings', 'settings.xml', 0),
         ('rId6', 'header', 'header1.xml', 0), ('rId7', 'footer', 'footer1.xml', 0), ('rId8', 'image', 'media/image1.png', 0), ('rId9', 'hyperlink', 'https://example.com/', 1)],
        [('word/styles.xml', CT + 'styles+xml'), ('word/numbering.xml', CT + 'numbering+xml'), ('word/settings.xml', CT + 'settings+xml'), ('word/footnotes.xml', CT + 'footnotes+xml'), ('word/comments.xml', CT + 'comments+xml'), ('word/header1.xml', CT + 'header+xml'), ('word/footer1.xml', CT + 'footer+xml')],
        sect)

def messy_docx():
    J = '<w:jc w:val="both"/>'
    body = (p(r('ANNUAL SUMMARY', '<w:b/><w:sz w:val="36"/><w:szCs w:val="36"/>')) + p('') + p('') + p('')
        + p(r('Background', '<w:b/><w:sz w:val="28"/><w:szCs w:val="28"/>'))
        + p(r('This year  the team  delivered more projects than ever before , and the results were "good" overall. ') + r('Some of this text was pasted in a different font. ', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>')
            + r('A smaller part. ', '<w:sz w:val="18"/><w:szCs w:val="18"/>') + r('A red part.', '<w:color w:val="FF0000"/>') + r(' It is the end of the first paragraph which is long enough to count as body text for the checks.'), J)
        + p(r('The second paragraph has "straight" quotes and ‘curly’ ones, and it is also long enough to be counted as body text in the checks.'), J)
        + p(r('This paragraph is aligned left while the others are justified, and it is long enough to count for the alignment checks of the document.'))
        + p(r('Another justified paragraph keeps the majority justified, so that the left aligned one above stands out as the odd one in this document.'), J)
        + p('', J) + p('', J)
        + p(r('Key Points', '<w:b/><w:sz w:val="28"/><w:szCs w:val="28"/>'))
        + p(r('• First point about the work')) + p(r('• Second point about the work')) + p(r('• Third point about the work'))
        + p(r('1. Step one of the plan')) + p(r('2. Step two of the plan'))
        + p('<w:r><w:tab/><w:t xml:space="preserve">This paragraph is indented with a tab instead of a first-line indent, which Fix Formatting should notice and fix.</w:t></w:r>', J)
        + p(r('These lines were pasted from an e-mail') + '<w:r><w:br/></w:r>' + r('so every line ends with a manual break') + '<w:r><w:br/></w:r>' + r('even in the middle of a sentence like this one.'), J)
        + p(r('Final words -- with dashes - and an ellipsis... to tidy.'), J))
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    package('messy.docx', body, {'word/styles.xml': STYLES, 'word/settings.xml': SETTINGS.replace('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>', '')},
        [('rId1', 'styles', 'styles.xml', 0), ('rId2', 'settings', 'settings.xml', 0)], [('word/styles.xml', CT + 'styles+xml'), ('word/settings.xml', CT + 'settings+xml')], sect)

complex_docx(); messy_docx()
print('complex.docx and messy.docx written to', OUT)
