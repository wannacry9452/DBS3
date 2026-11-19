"""
build_documents.py
Generates:
1. REPORT.html: A professional, print-ready document formatted to A4 pages with
   table of contents, typography, and CSS page-breaks so it prints/saves to a 10-15 page PDF.
2. REPORT.docx: A valid Microsoft Word document created using python's built-in zipfile & OpenXML.
"""

import os
import re
import html
import zipfile

def markdown_to_html(md_text: str) -> str:
    """Converts the markdown text to clean HTML with styling and table support."""
    lines = md_text.splitlines()
    html_lines = []
    in_code_block = False
    in_table = False
    table_header = True

    for line in lines:
        if line.startswith("```"):
            if in_code_block:
                html_lines.append("</code></pre>")
                in_code_block = False
            else:
                html_lines.append("<pre><code>")
                in_code_block = True
            continue

        if in_code_block:
            html_lines.append(html.escape(line))
            continue

        # Tables
        if line.strip().startswith("|") and line.strip().endswith("|"):
            cells = [c.strip() for c in line.strip().split("|")[1:-1]]
            if all(set(c).issubset({'-', ':', ' '}) for c in cells):
                table_header = False
                continue
            if not in_table:
                html_lines.append("<table class='doc-table'>")
                in_table = True
                table_header = True
            
            row_tag = "th" if table_header else "td"
            row_html = "".join(f"<{row_tag}>{html.escape(c)}</{row_tag}>" for c in cells)
            html_lines.append(f"<tr>{row_html}</tr>")
            continue
        else:
            if in_table:
                html_lines.append("</table>")
                in_table = False

        if not line.strip():
            html_lines.append("<br/>")
            continue

        # Headers
        if line.startswith("# "):
            html_lines.append(f"<h1 class='title-h1'>{html.escape(line[2:])}</h1>")
        elif line.startswith("## "):
            html_lines.append(f"<h2 class='title-h2'>{html.escape(line[3:])}</h2>")
        elif line.startswith("### "):
            html_lines.append(f"<h3 class='title-h3'>{html.escape(line[4:])}</h3>")
        elif line.startswith("#### "):
            html_lines.append(f"<h4 class='title-h4'>{html.escape(line[5:])}</h4>")
        elif line.startswith("---"):
            html_lines.append("<hr class='page-divider'/>")
        elif line.startswith("- "):
            clean_item = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html.escape(line[2:]))
            html_lines.append(f"<li class='doc-li'>{clean_item}</li>")
        else:
            # Inline bold and code
            processed = html.escape(line)
            processed = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', processed)
            processed = re.sub(r'\*(.*?)\*', r'<em>\1</em>', processed)
            processed = re.sub(r'`(.*?)`', r'<code>\1</code>', processed)
            html_lines.append(f"<p>{processed}</p>")

    if in_table:
        html_lines.append("</table>")
    if in_code_block:
        html_lines.append("</code></pre>")

    return "\n".join(html_lines)


def generate_html_report(md_path="REPORT.md", html_path="REPORT.html"):
    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    body_html = markdown_to_html(md_content)

    html_template = f"""<!DOCTYPE html>
<html lang="mn">
<head>
<meta charset="UTF-8">
<title>Бие даалтын ажил - NoSQL өгөгдлийн сангууд</title>
<style>
    @page {{
        size: A4;
        margin: 20mm 15mm 20mm 15mm;
        @bottom-right {{
            content: counter(page);
        }}
    }}
    body {{
        font-family: 'Segoe UI', Arial, sans-serif;
        line-height: 1.6;
        color: #24292e;
        max-width: 900px;
        margin: 0 auto;
        padding: 30px;
        background: #ffffff;
    }}
    h1.title-h1 {{
        color: #0b3c5d;
        border-bottom: 3px solid #328cc1;
        padding-bottom: 8px;
        margin-top: 35px;
        font-size: 26px;
        page-break-after: avoid;
    }}
    h2.title-h2 {{
        color: #1d2731;
        border-bottom: 1.5px solid #eaecef;
        padding-bottom: 6px;
        margin-top: 28px;
        font-size: 20px;
        page-break-after: avoid;
    }}
    h3.title-h3 {{
        color: #328cc1;
        margin-top: 20px;
        font-size: 16px;
        page-break-after: avoid;
    }}
    h4.title-h4 {{
        color: #555;
        font-size: 14px;
        margin-top: 14px;
    }}
    p {{
        margin: 8px 0;
        text-align: justify;
        font-size: 14px;
    }}
    pre {{
        background: #f6f8fa;
        border: 1px solid #e1e4e8;
        border-radius: 6px;
        padding: 12px;
        overflow-x: auto;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 12px;
        line-height: 1.45;
        page-break-inside: avoid;
    }}
    code {{
        background: #f0f2f5;
        padding: 2px 5px;
        border-radius: 3px;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        color: #d73a49;
    }}
    pre code {{
        background: transparent;
        color: #24292e;
        padding: 0;
    }}
    table.doc-table {{
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-size: 13px;
        page-break-inside: avoid;
    }}
    table.doc-table th, table.doc-table td {{
        border: 1px solid #dfe2e5;
        padding: 8px 12px;
        text-align: left;
    }}
    table.doc-table th {{
        background-color: #f1f8ff;
        font-weight: 600;
        color: #0b3c5d;
    }}
    table.doc-table tr:nth-child(even) {{
        background-color: #fbfcfd;
    }}
    li.doc-li {{
        margin-left: 20px;
        font-size: 14px;
        margin-bottom: 4px;
    }}
    hr.page-divider {{
        border: none;
        border-top: 2px dashed #ccc;
        margin: 25px 0;
    }}
    .cover {{
        text-align: center;
        padding: 50px 0;
        border: 2px solid #328cc1;
        border-radius: 8px;
        background: #fafbfc;
        margin-bottom: 40px;
        page-break-after: always;
    }}
    .cover h1 {{
        font-size: 28px;
        color: #0b3c5d;
        margin-bottom: 15px;
    }}
    .cover p {{
        font-size: 16px;
        color: #555;
        text-align: center;
        margin: 6px 0;
    }}
</style>
</head>
<body>
<div class="cover">
    <h1>БИЕ ДААЛТЫН ТАЙЛАН</h1>
    <p><strong>Сэдэв:</strong> NoSQL өгөгдлийн сангууд — Онол, Загварчлал, Хэрэглээ, Туршилтын төсөл</p>
    <p><strong>Хичээл:</strong> Өгөгдлийн сангийн систем (Database Systems)</p>
    <p><strong>Хэрэгжүүлэлт:</strong> MongoDB, Python (Embedded Pure Python Engine & PyMongo Compatible)</p>
    <p><strong>Хэмжээ:</strong> 10–15 хуудас бүхий цогц эрдэм шинжилгээ, туршилтын ажил</p>
    <p><strong>Огноо:</strong> 2024 он</p>
</div>

{body_html}

</body>
</html>
"""
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"[OK] Generated print-ready HTML: '{html_path}'")


def generate_docx_package(md_path="REPORT.md", docx_path="REPORT.docx"):
    """
    Creates a valid standard Microsoft Word (.docx) document directly
    using Python standard zipfile and OpenXML specification.
    """
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    # Build document.xml paragraphs
    paragraphs_xml = []

    def escape_xml(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

    for line in md_text.splitlines():
        line = line.strip()
        if not line:
            paragraphs_xml.append("<w:p/>")
            continue

        if line.startswith("# "):
            text = escape_xml(line[2:])
            p = f"""<w:p>
                <w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="360" w:after="120"/></w:pPr>
                <w:r><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="0B3C5D"/></w:rPr><w:t>{text}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)
        elif line.startswith("## "):
            text = escape_xml(line[3:])
            p = f"""<w:p>
                <w:pPr><w:pStyle w:val="Heading2"/><w:spacing w:before="240" w:after="80"/></w:pPr>
                <w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="328CC1"/></w:rPr><w:t>{text}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)
        elif line.startswith("### "):
            text = escape_xml(line[4:])
            p = f"""<w:p>
                <w:pPr><w:pStyle w:val="Heading3"/><w:spacing w:before="180" w:after="60"/></w:pPr>
                <w:r><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="1D2731"/></w:rPr><w:t>{text}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)
        elif line.startswith("- "):
            text = escape_xml(line[2:])
            # Clean bold markers
            text = text.replace("**", "")
            p = f"""<w:p>
                <w:pPr><w:spacing w:before="40" w:after="40"/><w:ind w:left="360"/></w:pPr>
                <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>• {text}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)
        elif line.startswith("|") and line.endswith("|"):
            # Render simple table text in docx
            cells = [escape_xml(c.strip()) for c in line.split("|")[1:-1]]
            if all(set(c).issubset({'-', ':', ' '}) for c in cells):
                continue
            cells_str = " | ".join(cells)
            p = f"""<w:p>
                <w:pPr><w:spacing w:before="30" w:after="30"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/><w:color w:val="2B4C7E"/></w:rPr><w:t>{cells_str}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)
        elif line.startswith("```"):
            continue
        else:
            clean_line = escape_xml(line).replace("**", "")
            p = f"""<w:p>
                <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
                <w:r><w:rPr><w:sz w:val="22"/><w:color w:val="24292E"/></w:rPr><w:t>{clean_line}</w:t></w:r>
            </w:p>"""
            paragraphs_xml.append(p)

    body_xml_str = "\n".join(paragraphs_xml)

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="400" w:after="200"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="44"/><w:color w:val="0B3C5D"/></w:rPr><w:t>БИЕ ДААЛТЫН ТАЙЛАН</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="100" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="328CC1"/></w:rPr><w:t>Сэдэв: NoSQL өгөгдлийн сангууд — Онол, Загварчлал, Хэрэглээ, Туршилтын төсөл</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="100" w:after="400"/></w:pPr>
      <w:r><w:rPr><w:i/><w:sz w:val="22"/><w:color w:val="555555"/></w:rPr><w:t>Монгол Улс, Улаанбаатар хот — 2024 он</w:t></w:r>
    </w:p>
    {body_xml_str}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", rels_xml)
        docx.writestr("word/document.xml", document_xml)

    print(f"[OK] Generated Microsoft Word file: '{docx_path}'")


if __name__ == "__main__":
    generate_html_report()
    generate_docx_package()
