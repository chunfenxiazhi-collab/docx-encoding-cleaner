# -*- coding: utf-8 -*-
"""测试用的最小 docx 生成器（不依赖第三方库）。"""

from __future__ import annotations

import zipfile
from pathlib import Path

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# 复刻真实故障文档的形状：WPS 写出的 <w:cr/> 软回车 + 行尾 U+FEFF，
# 以及 <w:u w:color="auto"/> 这类 WPS 特有的 run 属性。
DOCUMENT_WITH_JUNK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p>"
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="\u5b8b\u4f53"/>'
    '<w:u w:color="auto"/></w:rPr><w:t>\u4eba\u7269\u4ecb\u7ecd</w:t></w:r>'
    # run 1：行尾 U+FEFF + 软回车
    '<w:r><w:rPr><w:u w:color="auto"/></w:rPr><w:t>\u4eba\u7269 1\uff1a\u7532' + "\ufeff"
    "</w:t></w:r>"
    "<w:r><w:rPr><w:u w:color=\"auto\"/></w:rPr><w:cr/></w:r>"
    # run 2：夹在文本中的零宽字符（U+200B / U+200D / U+2060）
    '<w:r><w:rPr><w:u w:color="auto"/></w:rPr>'
    "<w:t>\u5f85\u4e1a\u9752\u5e74\u200b\uff0c\u751f\u6d3b\u200d\u677e\u2060\u5f1b\u3002"
    + "\ufeff</w:t></w:r>"
    "<w:r><w:rPr><w:u w:color=\"auto\"/></w:rPr><w:cr/></w:r>"
    # run 3：干净的普通文本，不该被动
    '<w:r><w:rPr><w:u w:color="auto"/></w:rPr><w:t>\u4eba\u7269 2\uff1aEXAMPLE-01</w:t></w:r>'
    "<w:r><w:rPr><w:u w:color=\"auto\"/></w:rPr><w:cr/></w:r>"
    '<w:r><w:rPr><w:u w:color="auto"/></w:rPr><w:t>\u6545\u4e8b\u80cc\u666f</w:t></w:r>'
    "</w:p></w:body></w:document>"
)

PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082"
)


def build_docx(path: Path, document_xml: str = DOCUMENT_WITH_JUNK, with_media: bool = True) -> Path:
    """写一个结构完整的最小 docx。"""
    path = Path(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        if with_media:
            z.writestr("word/media/image1.png", PNG_BYTES)
    return path


def read_part(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def read_bytes(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as z:
        return z.read(name)


def namelist(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as z:
        return z.namelist()
