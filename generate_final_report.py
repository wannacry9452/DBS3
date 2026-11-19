"""
generate_final_report.py
Generates:
1. REPORT.pdf  - 18-19 page academic report matching the user's PDF with embedded screenshots
2. REPORT.docx - Microsoft Word document with embedded screenshots and full tables
3. REPORT.html - Styled responsive HTML document with embedded base64 screenshots
"""

import os
import base64
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether, HRFlowable, Preformatted
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROD_HEADERS = ["ID", "Бүтээгдэхүүний нэр", "Ангилал", "Үнэ (₮)", "Үлдэгдэл", "Онцлог талбар"]
PROD_ROWS = [
    [101, "Lenovo ThinkPad X1 Carbon", "LAPTOP", "5,450,000", 14, "i7, 16GB RAM"],
    [102, "Apple MacBook Air M3", "LAPTOP", "4,600,000", 20, "M3, 16GB RAM"],
    [103, "ASUS ROG Zephyrus G14", "LAPTOP", "6,200,000", 8, "Ryzen 9, RTX 4060"],
    [104, "Dell XPS 15 9530", "LAPTOP", "5,950,000", 10, "i7, OLED 3.5K"],
    [105, "HP Pavilion 15", "LAPTOP", "2,450,000", 25, "i5, 8GB RAM"],
    [106, "iPhone 15 Pro 256GB", "SMARTPHONE", "4,350,000", 18, "A17 Pro, Titanium"],
    [107, "Samsung Galaxy S24 Ultra", "SMARTPHONE", "4,200,000", 15, "S-Pen, 200MP"],
    [108, "Google Pixel 8 Pro", "SMARTPHONE", "3,150,000", 12, "Tensor G3, AI Cam"],
    [109, "Xiaomi 14 5G", "SMARTPHONE", "2,650,000", 22, "Leica Lens, 90W"],
    [110, "OnePlus 12R", "SMARTPHONE", "1,950,000", 19, "5500mAh, 120Hz"],
    [111, "Sony WH-1000XM5", "AUDIO", "1,280,000", 30, "ANC, 30h Battery"],
    [112, "Apple AirPods Pro 2", "AUDIO", "890,000", 40, "USB-C, H2 Chip"],
    [113, "Bose QuietComfort Ultra", "AUDIO", "1,420,000", 16, "Spatial Audio"],
    [114, "JBL Charge 5 Speaker", "AUDIO", "580,000", 28, "IP67, 20h Play"],
    [115, "Marshall Major IV", "AUDIO", "490,000", 35, "80h Wireless"],
    [116, "Logitech MX Master 3S", "ACCESSORY", "360,000", 45, "8000 DPI, Quiet"],
    [117, "Keychron K2 Pro Keyboard", "ACCESSORY", "390,000", 24, "Hot-swap, RGB"],
    [118, "Anker 737 PowerBank 24K", "ACCESSORY", "440,000", 32, "140W PD Output"],
    [119, "Samsung T7 Shield 1TB", "ACCESSORY", "380,000", 27, "1050MB/s, IP65"],
    [120, "Ugreen 100W GaN Charger", "ACCESSORY", "185,000", 50, "4-Port USB-C"],
    [121, "LG UltraFine 27\" 4K", "MONITOR", "1,750,000", 11, "IPS, 99% sRGB"],
    [122, "Dell UltraSharp U2723QE", "MONITOR", "2,150,000", 9, "IPS Black, Hub"],
    [123, "Samsung Odyssey G7 32\"", "MONITOR", "2,300,000", 7, "240Hz, Curved"],
    [124, "ASUS ProArt PA278QV", "MONITOR", "1,180,000", 15, "Calman Verified"],
    [125, "BenQ GW2480 24\"", "MONITOR", "480,000", 34, "Eye-Care, FHD"]
]

AGG_HEADERS = ["Ангилал", "Барааны тоо", "Дундаж үнэ (₮)", "Хамгийн өндөр үнэ", "Агуулахын нийт дүн (₮)"]
AGG_ROWS = [
    ["LAPTOP", "5", "4,930,000", "6,200,000", "338,650,000"],
    ["SMARTPHONE", "5", "3,260,000", "4,350,000", "274,450,000"],
    ["AUDIO", "5", "932,000", "1,420,000", "130,110,000"],
    ["MONITOR", "5", "1,572,000", "2,300,000", "88,720,000"],
    ["ACCESSORY", "5", "351,000", "440,000", "59,150,000"]
]

# -------------------------------------------------------------
# 1. DOCX GENERATION
# -------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def generate_docx():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Segoe UI'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x21, 0x25, 0x29)

    # Header
    p_top = doc.add_paragraph()
    r_top = p_top.add_run("БИЕ ДААЛТЫН АЖИЛ №3")
    r_top.font.size = Pt(11)
    r_top.font.bold = True
    r_top.font.color.rgb = RGBColor(0x00, 0x66, 0x99)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run("NoSQL өгөгдлийн сангууд: Онол, Загварчлал, Хэрэглээ ба Туршилтын төсөл")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x0B, 0x3C, 0x5D)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Цахим худалдааны бүтээгдэхүүний каталогийн системийг MongoDB платформ дээр хэрэгжүүлсэн судалгааны тайлан")
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    meta_table = doc.add_table(rows=3, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Гүйцэтгэсэн оюутан:", True), ("[Оюутны овог, нэр]", False), ("Оюутны код:", True), ("[Оюутны код]", False)],
        [("Анги, бүлэг:", True), ("[Анги, бүлэг]", False), ("Шалгасан багш:", True), ("[Багшийн нэр]", False)],
        [("Сонгосон систем:", True), ("MongoDB (Document Store)", False), ("Огноо:", True), ("2026 он", False)],
    ]
    for r_idx, row in enumerate(meta_table.rows):
        for c_idx, cell in enumerate(row.cells):
            text, is_bold = meta_data[r_idx][c_idx]
            cell.text = text
            set_cell_background(cell, "F8F9FA" if is_bold else "FFFFFF")
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.bold = is_bold
            if is_bold:
                p.runs[0].font.color.rgb = RGBColor(0x0B, 0x3C, 0x5D)

    doc.add_paragraph()

    abs_table = doc.add_table(rows=1, cols=1)
    abs_cell = abs_table.rows[0].cells[0]
    set_cell_background(abs_cell, "EBF3FA")
    set_cell_margins(abs_cell, top=120, bottom=120, left=180, right=180)
    p_abs = abs_cell.paragraphs[0]
    r_abs_h = p_abs.add_run("Ажлын зорилго ба товч агуулга: ")
    r_abs_h.font.bold = True
    r_abs_h.font.color.rgb = RGBColor(0x00, 0x40, 0x85)
    r_abs_b = p_abs.add_run(
        "Энэхүү бие даалтын ажил нь орчин үеийн их хэмжээний өгөгдөл боловсруулах шаардлагад нийцүүлэн үүссэн "
        "NoSQL өгөгдлийн сангуудын онолын үндэс, дөрвөн үндсэн төрлийн архитектурын ялгаа, CAP теоремийн практик "
        "ач холбогдол болон өгөгдөл загварчлалын зарчмыг системтэйгээр судлахад чиглэнэ. Онолын судалгаанд тулгуурлан "
        "MongoDB платформ болон Python (PyMongo) орчинд цахим худалдааны бүтээгдэхүүний каталогийн туршилтын системийг "
        "загварчлан байгуулж, CRUD үйлдлүүд, индексжүүлэлт, текст хайлт болон Aggregation Pipeline-ийн дэвшилтэт "
        "боломжуудыг туршин, хамааралт (SQL) системтэй харьцуулсан дүн шинжилгээ хийв."
    )
    r_abs_b.font.size = Pt(9.5)

    doc.add_paragraph()

    def add_h1(text):
        h = doc.add_heading(text, level=1)
        h.paragraph_format.space_before = Pt(16)
        h.paragraph_format.space_after = Pt(6)
        for r in h.runs:
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x0B, 0x3C, 0x5D)

    def add_h2(text):
        h = doc.add_heading(text, level=2)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(4)
        for r in h.runs:
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x00, 0x66, 0x99)

    def add_h3(text):
        h = doc.add_heading(text, level=3)
        h.paragraph_format.space_before = Pt(9)
        h.paragraph_format.space_after = Pt(3)
        for r in h.runs:
            r.font.size = Pt(10.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x21, 0x25, 0x29)

    def add_p(text):
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15
        return p

    def add_code(code_text):
        table = doc.add_table(rows=1, cols=1)
        cell = table.rows[0].cells[0]
        set_cell_background(cell, "F6F8FA")
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(code_text)
        run.font.name = 'Consolas'
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
        doc.add_paragraph()

    # Section 1
    add_h1("1. Онолын судалгааны хэсэг")
    add_h2("1.1. SQL ба NoSQL — Суурь ойлголт")
    add_p("Өгөгдлийн сангийн удирдлагын системүүд (DBMS) нь програм хангамжийн инженерчлэлийн суурь дэд бүтэц болж ирсэн. 1970 онд Эдгар Коддын боловсруулсан хамааралт загвар (Codd, 1970) нь өгөгдлийг мөр, баганаас бүрдэх хүснэгт хэлбэрээр зохион байгуулж, хүснэгт хоорондын холбоосыг гадаад түлхүүр (Foreign Key)-ээр илэрхийлэх замаар өгөгдлийн давхардлыг бууруулж, бүрэн бүтэн байдлыг баталгаажуулсан юм. Энэхүү зарчим дээр суурилсан RDBMS (Relational Database Management System — жишээ нь PostgreSQL, MySQL, Oracle, MS SQL Server) нь SQL (Structured Query Language) хэл ашиглан хатуу тодорхойлсон схем (Schema-on-Write) болон ACID (Atomicity, Consistency, Isolation, Durability) гүйлгээний баталгааг хангадаг.")
    add_p("Гэвч 2000-аад оны дунд үеэс интернэт хэрэглэгчдийн тоо огцом өсөж, нийгмийн сүлжээ, цахим худалдаа, үүлэн тооцоолол болон IoT төхөөрөмжүүдээс секунд тутамд асар их хэмжээний бүтэцгүй болон хагас бүтэцтэй өгөгдөл (Big Data) үүсэх болсон. Уламжлалт RDBMS архитектур нь дараах гурван үндсэн сорилттой тулгарсан (Sadalage & Fowler, 2012):")
    add_p("• Хэвтээ өргөтгөлийн хязгаарлалт (Horizontal Scaling): Хамааралт өгөгдлийн сан нь хүснэгт хоорондын JOIN үйлдэл болон ACID гүйлгээг хадгалахын тулд ихэвчлэн нэг хүчирхэг серверийн хүчин чадлыг нэмэгдүүлэх буюу босоо өргөтгөлт (Vertical Scaling / Scale-up) хийдэг. Харин NoSQL нь олон энгийн серверийн кластерт өгөгдлийг тараан байрлуулах (Scale-out) зарчмаар бүтээгдсэн.")
    add_p("• Схемийн уян хатан бус байдал: Бизнесийн шаардлага хурдан өөрчлөгддөг орчин үеийн веб системүүдэд өгөгдлийн бүтэц байнга шинэчлэгддэг. Сая сая бичлэгтэй SQL хүснэгтэд ALTER TABLE хийх нь системийг удаашруулах эсвэл зогсолт (downtime) үүсгэх эрсдэлтэй.")
    add_p("• Их хэмжээний өгөгдөл ба тархсан системийн хэрэгцээ: Өндөр ачаалалтай системүүдэд секундийн мянганы нарийвчлалтай хариу үйлдэл (low latency), тасралтгүй ажиллагаа (high availability), газарзүйн олон бүсэд өгөгдөл хуулах (geo-replication) чадвар шаардагддаг. NoSQL нь ACID-ийн хатуу шаардлагыг тодорхой хэмжээнд сулруулж, BASE загварыг нэвтрүүлснээр тархсан орчинд өндөр гүйцэтгэл гаргах боломжийг олгодог (Kleppmann, 2017).")

    add_h2("1.2. NoSQL системүүдийн 4 үндсэн төрөл")
    add_p("NoSQL өгөгдлийн сангууд нь нэгдмэл ганц архитектур биш бөгөөд өгөгдлийг хадгалах, индексжүүлэх болон хандах аргачлалаасаа хамааран дөрвөн үндсэн төрөлд ангилагддаг (Sadalage & Fowler, 2012):")
    add_h3("1) Key-Value Store (Түлхүүр-утга хадгалагч)")
    add_p("Хамгийн энгийн бүтэцтэй NoSQL төрөл бөгөөд өгөгдлийг хэш хүснэгт (hash table) буюу өвөрмөц түлхүүр (key) ба түүнд харгалзах дурын утгын (value) хослол хэлбэрээр хадгалдаг. Систем нь утгын дотоод бүтцийг задлан шалгах шаардлагагүй тул санах ойд эсвэл диск дээр O(1) хугацаанд уншиж бичдэг. Жишээ: Redis, Amazon DynamoDB, Memcached.")
    add_h3("2) Document Store (Баримтад суурилсан өгөгдлийн сан)")
    add_p("Key-Value загварыг өргөтгөсөн хэлбэр бөгөөд утга хэсэг нь JSON/BSON баримт байна. Баримтын дотоод шаталсан талбар болон массив дээр индекс үүсгэж, нарийн шүүлтүүр, агрегаци хийх чадвартай. Жишээ: MongoDB, CouchDB, Firestore.")
    add_h3("3) Column-Family Store (Өргөн баганын өгөгдлийн сан)")
    add_p("Google Bigtable-д суурилсан, баганын бүлгээр хадгалдаг. Диск дээр өндөр хурдтай бичилт хийдэг тул цагийн цуваа өгөгдөл, IoT сенсорт төгс зохицдог. Жишээ: Apache Cassandra, ScyllaDB, HBase.")
    add_h3("4) Graph Database (Граф өгөгдлийн сан)")
    add_p("Орой (Node) ба Ирмэг (Edge) бүхий математик графын онолд суурилсан. Index-free adjacency бүтэцтэй тул олон шатлалт хамаарлыг миллисекундэд хайдаг. Жишээ: Neo4j, ArangoDB.")

    add_h2("1.3. CAP Теорем — Дэлгэрэнгүй тайлбар")
    add_p("Эрик Брюэрийн CAP теорем ёсоор тархсан өгөгдлийн систем нь Consistency (C), Availability (A), Partition Tolerance (P) гурваас нэгэн зэрэг хамгийн ихдээ хоёрыг л хангана. Бодит сүлжээнд сүлжээний тасралт (P) гарцаагүй тул системүүд CP (MongoDB) эсвэл AP (Cassandra) загварыг сонгодог.")

    add_h2("1.4. NoSQL-ийн өгөгдлийн загварчлал")
    add_p("NoSQL-д Query-driven загварчлал баримталдаг. Бүтээгдэхүүний техникийн үзүүлэлт болон сэтгэгдлийг шигтгэж (Embedding), захиалгын мэдээллийг холбох (Referencing) эрлийз загвар нь хамгийн өндөр үр ашигтай байдаг.")

    add_h2("1.5. Том системүүдийн хэрэгжүүлэлт: Netflix ба Apache Cassandra")
    add_p("Netflix нь 200 сая хэрэглэгчийн үзэлтийн түүхийг Apache Cassandra ашиглан AP зарчмаар 99.999% найдвартай ажиллуулдаг.")

    # Section 2
    add_h1("2. Практик туршилтын төсөл: Цахим худалдааны бүтээгдэхүүний каталог (MongoDB)")
    add_h2("2.1. Системийн зорилго ба 2.2. Загварчлал")
    add_p("Туршилтын хүрээнд MongoDB болон Python (PyMongo) ашиглан Цахим худалдааны бүтээгдэхүүний каталогийн системийг хөгжүүлсэн. products, categories, orders цуглуулгуудтай.")

    add_h2("2.3. Бүтээгдэхүүний жагсаалт (25 баримт)")
    t_prod = doc.add_table(rows=26, cols=6)
    t_prod.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h in enumerate(PROD_HEADERS):
        cell = t_prod.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "0B3C5D")
        set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
        p = cell.paragraphs[0]
        p.runs[0].font.size = Pt(8)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row in enumerate(PROD_ROWS):
        for c_idx, val in enumerate(row):
            cell = t_prod.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            set_cell_background(cell, "F8F9FA" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(7.5)

    doc.add_paragraph()

    add_h2("2.4. CRUD үйлдлүүд (Python / PyMongo)")
    add_code("""# crud_operations.py үр дүн
[CREATE] Нэмэгдсэн _id: 673c41a9f2d8e4b1c091a201

[READ] Шүүсэн зөөврийн компьютерууд:
  {'product_id': 105, 'name': 'HP Pavilion 15', 'price': 2450000, 'stock': 25}
  {'product_id': 102, 'name': 'Apple MacBook Air M3', 'price': 4600000, 'stock': 20}

[UPDATE] Шинэчлэгдсэн баримт: 1
[DELETE] Устгагдсан баримт: 1""")

    add_h2("2.5. Aggregation Pipeline-ийн тооцоолсон үр дүн")
    t_agg = doc.add_table(rows=6, cols=5)
    t_agg.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h in enumerate(AGG_HEADERS):
        cell = t_agg.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "0B3C5D")
        set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
        p = cell.paragraphs[0]
        p.runs[0].font.size = Pt(8)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row in enumerate(AGG_ROWS):
        for c_idx, val in enumerate(row):
            cell = t_agg.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            set_cell_background(cell, "F8F9FA" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(8)

    doc.add_paragraph()

    # Section 2.6: Screenshots
    add_h2("2.6. Ажлын дэлгэцийн зургийн хэсэг (Туршилтын баталгаажуулалт)")
    add_p("Системийг бодитоор ажиллуулж авсан дэлгэцийн зургуудыг (screenshots) доор хавсаргав:")

    add_h3("[Дэлгэцийн зураг 1 — Өгөгдлийн сангийн бүтэц (MongoDB Compass)]")
    add_p("MongoDB Compass график интерфэйс дээр ecommerce_catalog сан, products цуглуулга ба 25 баримтын харагдац:")
    doc.add_picture(os.path.join(BASE_DIR, "screenshot1_database_structure.png"), width=Inches(6.2))
    doc.add_paragraph()

    add_h3("[Дэлгэцийн зураг 2 — CRUD үйлдлийн гүйцэтгэл]")
    add_p("Терминал дээр python crud_operations.py ажиллуулсны дараах Create, Read, Update, Delete үр дүн:")
    doc.add_picture(os.path.join(BASE_DIR, "screenshot2_crud_operations.png"), width=Inches(6.2))
    doc.add_paragraph()

    add_h3("[Дэлгэцийн зураг 3 — Нийлмэл индекс ба Текст хайлтын үр дүн]")
    add_p("MongoDB Compass-ийн Indexes таб (category_code_1_price_-1 ба Text Index) болон $text хайлтын хариу:")
    doc.add_picture(os.path.join(BASE_DIR, "screenshot3_indexes_text_search.png"), width=Inches(6.2))
    doc.add_paragraph()

    add_h3("[Дэлгэцийн зураг 4 — Aggregation Pipeline-ийн гүйцэтгэл]")
    add_p("Терминал дээр python advanced_features.py ажиллуулж, ангилал бүрийн дундаж үнэ, агуулахын нийт дүнг тооцсон дүн:")
    doc.add_picture(os.path.join(BASE_DIR, "screenshot4_aggregation_pipeline.png"), width=Inches(6.2))
    doc.add_paragraph()

    # Section 3 & 4
    add_h1("3. Дүн шинжилгээ ба Харьцуулалт")
    add_p("• MongoDB-ийн SQL-тэй харьцуулсан давуу тал: Олон төрлийн барааны техникийн үзүүлэлтийг specs дэд объектод хялбархан хадгалж, JOIN үйлдэлгүйгээр ганцхан find_one() асуулгаар сэтгэгдэлтэй нь хамт хурдан уншсан.")
    add_p("• Ажиглагдсан сул тал: SQL шиг гадаад түлхүүрийн хатуу хяналт анхдагчаар байхгүй, өгөгдлийн давхардал үүсэх, талбарын нэр давтагддаг тул санах ой илүү зарцуулдаг.")
    add_p("• Өөр хэрэгцээнд илүү тохиромжтой NoSQL төрлүүд: Сагс болон сессэд Redis, зөвлөмжийн системд Neo4j, их хэмжээний логт Cassandra тохиромжтой.")
    add_p("• Хамгийн чухал ойлголт: Төгс өгөгдлийн сан гэж байдаггүй; өгөгдлийн загварчлал болон CAP-ийн сонголт нь бизнесийн уншилт-бичилтийн шаардлагаас шууд хамаардаг.")

    add_h1("4. Дүгнэлт ба Санал зөвлөмж")
    add_p("Энэхүү бие даалтын ажлаар NoSQL өгөгдлийн сангуудын онол, архитектур, CAP теорем, өгөгдөл загварчлалын зарчмыг цогцоор судалж, MongoDB платформ дээр 25 барааны өгөгдөлтэй цахим худалдааны системийг амжилттай хэрэгжүүлэн баталгаажууллаа.")

    add_h1("Ашигласан эх сурвалж")
    add_p("1. Codd, E. F. (1970). A Relational Model of Data for Large Shared Data Banks. CACM.\n2. Brewer, E. A. (2000). Towards Robust Distributed Systems. PODC.\n3. Gilbert, S., & Lynch, N. (2002). Brewer's Conjecture and Feasibility of CAP Services. ACM SIGACT News.\n4. Brewer, E. (2012). CAP Twelve Years Later: How the Rules Have Changed. IEEE Computer.\n5. Sadalage, P. J., & Fowler, M. (2012). NoSQL Distilled. Addison-Wesley.\n6. Kleppmann, M. (2017). Designing Data-Intensive Applications. O'Reilly Media.\n7. Cockcroft, A., & Sheahan, D. (2011). Benchmarking Cassandra Scalability on AWS. Netflix TechBlog.\n8. MongoDB, Inc. MongoDB Manual: Data Modeling, Aggregation Pipeline, and Indexes.")

    docx_path = os.path.join(BASE_DIR, "REPORT.docx")
    doc.save(docx_path)
    print(f"[OK] Generated Word document: '{docx_path}'")


# -------------------------------------------------------------
# 2. PDF GENERATION (Multi-page Full Layout with ReportLab)
# -------------------------------------------------------------
def generate_pdf():
    pdfmetrics.registerFont(TTFont('SegoeUI', 'C:/Windows/Fonts/segoeui.ttf'))
    pdfmetrics.registerFont(TTFont('SegoeUI-Bold', 'C:/Windows/Fonts/segoeuib.ttf'))
    pdfmetrics.registerFont(TTFont('SegoeUI-Italic', 'C:/Windows/Fonts/segoeuii.ttf'))
    pdfmetrics.registerFont(TTFont('Consolas', 'C:/Windows/Fonts/consola.ttf'))

    pdf_path = os.path.join(BASE_DIR, "REPORT.pdf")
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=48,
        rightMargin=48,
        topMargin=48,
        bottomMargin=48
    )

    styles = getSampleStyleSheet()

    s_top = ParagraphStyle('STop', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=10, textColor=colors.HexColor('#006699'), spaceAfter=4)
    s_title = ParagraphStyle('STitle', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=16, leading=20, textColor=colors.HexColor('#0B3C5D'), spaceAfter=6)
    s_sub = ParagraphStyle('SSub', parent=styles['Normal'], fontName='SegoeUI-Italic', fontSize=10, leading=13, textColor=colors.HexColor('#555555'), spaceAfter=12)
    s_h1 = ParagraphStyle('SH1', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=13, leading=16, textColor=colors.HexColor('#0B3C5D'), spaceBefore=14, spaceAfter=6, keepWithNext=True)
    s_h2 = ParagraphStyle('SH2', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=11, leading=14, textColor=colors.HexColor('#006699'), spaceBefore=10, spaceAfter=4, keepWithNext=True)
    s_h3 = ParagraphStyle('SH3', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=9.5, leading=12, textColor=colors.HexColor('#212529'), spaceBefore=8, spaceAfter=3, keepWithNext=True)
    s_body = ParagraphStyle('SBody', parent=styles['Normal'], fontName='SegoeUI', fontSize=9, leading=13, textColor=colors.HexColor('#212529'), spaceAfter=5)
    s_code = ParagraphStyle('SCode', parent=styles['Normal'], fontName='Consolas', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#24292E'))
    s_cell = ParagraphStyle('SCell', parent=styles['Normal'], fontName='SegoeUI', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#212529'))
    s_cell_head = ParagraphStyle('SCellHead', parent=styles['Normal'], fontName='SegoeUI-Bold', fontSize=7.5, leading=9.5, textColor=colors.white)

    story = []

    # PAGE 1: Header + Meta + Abstract + 1.1 Intro
    story.append(Paragraph("БИЕ ДААЛТЫН АЖИЛ №3", s_top))
    story.append(Paragraph("NoSQL өгөгдлийн сангууд: Онол, Загварчлал, Хэрэглээ ба Туршилтын төсөл", s_title))
    story.append(Paragraph("Цахим худалдааны бүтээгдэхүүний каталогийн системийг MongoDB платформ дээр хэрэгжүүлсэн судалгааны тайлан", s_sub))

    m_data = [
        [Paragraph("<b>Гүйцэтгэсэн оюутан:</b>", s_cell), Paragraph("[Оюутны овог, нэр]", s_cell), Paragraph("<b>Оюутны код:</b>", s_cell), Paragraph("[Оюутны код]", s_cell)],
        [Paragraph("<b>Анги, бүлэг:</b>", s_cell), Paragraph("[Анги, бүлэг]", s_cell), Paragraph("<b>Шалгасан багш:</b>", s_cell), Paragraph("[Багшийн нэр]", s_cell)],
        [Paragraph("<b>Сонгосон систем:</b>", s_cell), Paragraph("MongoDB (Document Store)", s_cell), Paragraph("<b>Огноо:</b>", s_cell), Paragraph("2026 он", s_cell)]
    ]
    t_m = Table(m_data, colWidths=[110, 140, 110, 139])
    t_m.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8F9FA')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_m)
    story.append(Spacer(1, 10))

    abs_p = (
        "<b>Ажлын зорилго ба товч агуулга:</b> Энэхүү бие даалтын ажил нь орчин үеийн их хэмжээний өгөгдөл "
        "боловсруулах шаардлагад нийцүүлэн үүссэн NoSQL өгөгдлийн сангуудын онолын үндэс, дөрвөн үндсэн төрлийн "
        "архитектурын ялгаа, CAP теоремийн практик ач холбогдол болон өгөгдөл загварчлалын зарчмыг системтэйгээр судлахад "
        "чиглэнэ. Онолын судалгаанд тулгуурлан MongoDB платформ болон Python (PyMongo) орчинд цахим худалдааны "
        "бүтээгдэхүүний каталогийн туршилтын системийг загварчлан байгуулж, CRUD үйлдлүүд, индексжүүлэлт, текст хайлт "
        "болон Aggregation Pipeline-ийн дэвшилтэт боломжуудыг туршин, хамааралт (SQL) системтэй харьцуулсан дүн шинжилгээ хийв."
    )
    t_abs = Table([[Paragraph(abs_p, s_body)]], colWidths=[499])
    t_abs.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#EBF3FA')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#006699')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_abs)
    story.append(Spacer(1, 10))

    story.append(Paragraph("1. Онолын судалгааны хэсэг", s_h1))
    story.append(Paragraph("1.1. SQL ба NoSQL — Суурь ойлголт", s_h2))
    story.append(Paragraph(
        "Өгөгдлийн сангийн удирдлагын системүүд (DBMS) нь програм хангамжийн инженерчлэлийн суурь дэд бүтэц болж ирсэн. "
        "1970 онд Эдгар Коддын боловсруулсан хамааралт загвар (Codd, 1970) нь өгөгдлийг мөр, баганаас бүрдэх хүснэгт хэлбэрээр "
        "зохион байгуулж, хүснэгт хоорондын холбоосыг гадаад түлхүүр (Foreign Key)-ээр илэрхийлэх замаар өгөгдлийн давхардлыг "
        "бууруулж, бүрэн бүтэн байдлыг баталгаажуулсан юм. Энэхүү зарчим дээр суурилсан RDBMS (Relational Database Management "
        "System — жишээ нь PostgreSQL, MySQL, Oracle, MS SQL Server) нь SQL (Structured Query Language) хэл ашиглан "
        "хатуу тодорхойлсон схем (Schema-on-Write) болон ACID (Atomicity, Consistency, Isolation, Durability) гүйлгээний "
        "баталгааг хангадаг.",
        s_body
    ))
    story.append(Paragraph(
        "Гэвч 2000-аад оны дунд үеэс интернэт хэрэглэгчдийн тоо огцом өсөж, нийгмийн сүлжээ, цахим худалдаа, үүлэн тооцоолол "
        "болон IoT төхөөрөмжүүдээс секунд тутамд асар их хэмжээний бүтэцгүй болон хагас бүтэцтэй өгөгдөл (Big Data) үүсэх болсон. "
        "Уламжлалт RDBMS архитектур нь дараах гурван үндсэн сорилттой тулгарсан (Sadalage & Fowler, 2012):",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 2: 1.1 Bullet points + Table 1 (SQL vs NoSQL) + 1.2 Intro
    story.append(Paragraph(
        "• <b>Хэвтээ өргөтгөлийн хязгаарлалт (Horizontal Scaling):</b> Хамааралт өгөгдлийн сан нь хүснэгт хоорондын JOIN үйлдэл "
        "болон ACID гүйлгээг хадгалахын тулд ихэвчлэн нэг хүчирхэг серверийн хүчин чадлыг нэмэгдүүлэх буюу босоо өргөтгөлт "
        "(Vertical Scaling / Scale-up) хийдэг. Энэ нь өртөг өндөртэй бөгөөд техник хангамжийн физик хязгаарт тулдаг. "
        "Харин NoSQL нь олон энгийн серверийн кластерт өгөгдлийг тараан байрлуулах (Scale-out) зарчмаар бүтээгдсэн.<br/>"
        "• <b>Схемийн уян хатан бус байдал:</b> Бизнесийн шаардлага хурдан өөрчлөгддөг орчин үеийн веб системүүдэд өгөгдлийн "
        "бүтэц байнга шинэчлэгддэг. Сая сая бичлэгтэй SQL хүснэгтэд ALTER TABLE хийх нь системийг удаашруулах эсвэл "
        "зогсолт (downtime) үүсгэх эрсдэлтэй.<br/>"
        "• <b>Их хэмжээний өгөгдөл ба тархсан системийн хэрэгцээ:</b> Өндөр ачаалалтай системүүдэд секундийн мянганы нарийвчлалтай "
        "хариу үйлдэл (low latency), тасралтгүй ажиллагаа (high availability), газарзүйн олон бүсэд өгөгдөл хуулах "
        "(geo-replication) чадвар шаардагддаг. NoSQL нь ACID-ийн хатуу шаардлагыг тодорхой хэмжээнд сулруулж, BASE "
        "(Basically Available, Soft state, Eventual consistency) загварыг нэвтрүүлснээр тархсан орчинд өндөр гүйцэтгэл "
        "гаргах боломжийг олгодог (Kleppmann, 2017).",
        s_body
    ))
    story.append(Spacer(1, 6))

    t1_p_data = [
        [Paragraph("<b>Шалгуур үзүүлэлт</b>", s_cell_head), Paragraph("<b>SQL (RDBMS)</b>", s_cell_head), Paragraph("<b>NoSQL (Non-Relational)</b>", s_cell_head)],
        [Paragraph("Өгөгдлийн загвар", s_cell), Paragraph("Хүснэгт, мөр, багана; хатуу хамаарал ба нормалчлал", s_cell), Paragraph("Баримт (JSON/BSON), түлхүүр-утга, өргөн багана эсвэл граф", s_cell)],
        [Paragraph("Схем (Schema)", s_cell), Paragraph("Урьдчилан хатуу тодорхойлсон (Static / Schema-on-Write)", s_cell), Paragraph("Уян хатан, динамик (Dynamic / Schema-on-Read)", s_cell)],
        [Paragraph("Өргөтгөлт (Scaling)", s_cell), Paragraph("Голчлон босоо өргөтгөлт (CPU, RAM нэмэх)", s_cell), Paragraph("Хэвтээ өргөтгөлт (олон сервер нэмж кластер болгох)", s_cell)],
        [Paragraph("Гүйлгээний баталгаа", s_cell), Paragraph("Хатуу ACID (Atomicity, Consistency, Isolation, Durability)", s_cell), Paragraph("Ихэвчлэн BASE (зарим системд олон баримтын ACID дэмждэг)", s_cell)],
        [Paragraph("Хүсэлтийн хэл", s_cell), Paragraph("Стандартчилагдсан SQL", s_cell), Paragraph("Систем бүрт өвөрмөц API, JSON query, CQL, Cypher гэх мэт", s_cell)],
        [Paragraph("Тохиромжтой хэрэглээ", s_cell), Paragraph("Санхүү, банк, нягтлан бодох бүртгэл, ERP систем", s_cell), Paragraph("Цахим худалдаа, нийгмийн сүлжээ, IoT, кэш, бодит цагийн аналитик", s_cell)]
    ]
    t1_pdf = Table(t1_p_data, colWidths=[110, 194, 195])
    t1_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t1_pdf)
    story.append(Spacer(1, 8))

    story.append(Paragraph("1.2. NoSQL системүүдийн 4 үндсэн төрөл", s_h2))
    story.append(Paragraph(
        "NoSQL өгөгдлийн сангууд нь нэгдмэл ганц архитектур биш бөгөөд өгөгдлийг хадгалах, индексжүүлэх болон хандах "
        "аргачлалаасаа хамааран дөрвөн үндсэн төрөлд ангилагддаг (Sadalage & Fowler, 2012). Төрөл тус бүрийн архитектурын "
        "онцлог, өгөгдлийн бүтэц, давуу ба сул тал, бодит хэрэглээг доор дэлгэрэнгүй авч үзье.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 3: Key-Value, Document, Column-Family Intro
    story.append(Paragraph("1) Key-Value Store (Түлхүүр-утга хадгалагч)", s_h3))
    story.append(Paragraph(
        "<b>Архитектурын онцлог ба өгөгдлийн бүтэц:</b> Хамгийн энгийн бүтэцтэй NoSQL төрөл бөгөөд өгөгдлийг хэш хүснэгт "
        "(hash table) буюу өвөрмөц түлхүүр (key) ба түүнд харгалзах дурын утгын (value — тэмдэгт мөр, JSON, хоёртын өгөгдөл) "
        "хослол хэлбэрээр хадгалдаг. Систем нь утгын дотоод бүтцийг задлан шалгах шаардлагагүй тул санах ойд (in-memory) "
        "эсвэл диск дээр O(1) хугацааны нарийвчлалтайгаар маш хурдан уншилт, бичилт хийдэг.<br/>"
        "• <b>Давуу тал:</b> Маш өндөр хурд (sub-millisecond latency), энгийн API (GET, SET, DELETE), хэвтээ өргөтгөлт хийхэд хялбар.<br/>"
        "• <b>Сул тал:</b> Утгын дотоод талбараар шүүлтүүр хийх, нарийн төвөгтэй нэгтгэл (aggregation) болон холбоост хайлт гүйцэтгэх боломж хязгаарлагдмал.<br/>"
        "• <b>Түгээмэл хэрэгжүүлэлтүүд:</b> Redis, Amazon DynamoDB, Memcached, Riak.<br/>"
        "• <b>Бодит жишээ:</b> Хэрэглэгчийн сесс (session) хадгалах, цахим худалдааны сагс (shopping cart), вэб хуудасны кэшлэлт (caching), тоглоомын онооны самбар.",
        s_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("2) Document Store (Баримтад суурилсан өгөгдлийн сан)", s_h3))
    story.append(Paragraph(
        "<b>Архитектурын онцлог ба өгөгдлийн бүтэц:</b> Key-Value загварыг өргөтгөсөн хэлбэр бөгөөд утга хэсэг нь JSON, BSON "
        "эсвэл XML форматаар бичигдсэн хагас бүтэцтэй баримт (document) байна. Өгөгдлийн сан нь баримтын дотоод бүтцийг "
        "ойлгодог тул шаталсан талбар (nested fields) болон массив (arrays) дээр индекс үүсгэж, нарийн шүүлтүүр, эрэмбэлэлт, "
        "нэгтгэлийн хүсэлт ажиллуулах чадвартай. Ижил зориулалттай баримтууд цуглуулгад (collection) нэгтгэгдэнэ.<br/>"
        "• <b>Давуу тал:</b> Объект хандалтат програмчлалын өгөгдлийн бүтэцтэй шууд нийцдэг (ORM-ийн хүндрэлгүй), уян хатан схемтэй, баримт доторх бүх талбараар хайлт хийх боломжтой.<br/>"
        "• <b>Сул тал:</b> Олон цуглуулга дамнасан нарийн JOIN үйлдэл хийхэд гүйцэтгэл буурдаг; өгөгдлийг давхардуулан хадгалснаар (denormalization) шинэчлэлтийн үед зөрүү гарах эрсдэлтэй.<br/>"
        "• <b>Түгээмэл хэрэгжүүлэлтүүд:</b> MongoDB, CouchDB, Firebase Firestore, Amazon DocumentDB.<br/>"
        "• <b>Бодит жишээ:</b> Цахим худалдааны бүтээгдэхүүний каталог, контент удирдлагын систем (CMS), хэрэглэгчийн профайл, мобайл аппликейшний өгөгдөл хадгалалт.",
        s_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("3) Column-Family Store (Өргөн баганын өгөгдлийн сан)", s_h3))
    story.append(Paragraph(
        "<b>Архитектурын онцлог ба өгөгдлийн бүтэц:</b> Google-ийн Bigtable судалгаанд суурилсан энэхүү төрөл нь өгөгдлийг "
        "мөр бүрээр бус, баганын бүлгээр (column family) хадгалдаг. Нэг мөр (row key) нь сая сая динамик багана агуулж болох ба "
        "хоосон багана диск дээр зай эзэлдэггүй. Тархсан кластерийн бүх зангилаа ижил эрхтэй (peer-to-peer / masterless) ажиллах тул "
        "нэг цэгийн доголдол (single point of failure) үүсдэггүй.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 4: Column Family cont. + Graph DB + Table 2
    story.append(Paragraph(
        "• <b>Давуу тал:</b> Петабайт хэмжээний өгөгдөл дээр маш өндөр хурдтай бичилт (write-heavy workload), өгөгдлийг үр ашигтай "
        "шахах, газарзүйн олон дата төвд тасралтгүй ажиллах чадвар.<br/>"
        "• <b>Сул тал:</b> Хүсэлтийн хэв маягийг (query pattern) урьдчилан нарийн төлөвлөж хүснэгтийн түлхүүрээ зохиох шаардлагатай; "
        "дурын баганаар шүүлт хийх болон гүйлгээ (transaction) хийхэд тохиромжгүй.<br/>"
        "• <b>Түгээмэл хэрэгжүүлэлтүүд:</b> Apache Cassandra, Apache HBase, ScyllaDB, Google Cloud Bigtable.<br/>"
        "• <b>Бодит жишээ:</b> IoT мэдрэгчийн цагийн цуваа (time-series) өгөгдөл, системийн лог бүртгэл, стриминг үзэлтийн түүх (Netflix).",
        s_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("4) Graph Database (Граф өгөгдлийн сан)", s_h3))
    story.append(Paragraph(
        "<b>Архитектурын онцлог ба өгөгдлийн бүтэц:</b> Математикийн графын онолд суурилж, өгөгдлийн объектуудыг орой "
        "(Node / Vertex), тэдгээрийн хоорондын холбоосыг ирмэг (Edge / Relationship), шинж чанарыг атрибут (Property) "
        "хэлбэрээр хадгалдаг. Холбоосыг физик түвшинд шууд заагчаар хадгалдаг (index-free adjacency) тул олон шатлалт "
        "хамаарлыг (жишээ нь найзын найзын найз) хайхад SQL-ийн JOIN-оос хэдэн зуу дахин хурдан ажилладаг.<br/>"
        "• <b>Давуу тал:</b> Нарийн төвөгтэй сүлжээн хамаарлыг илэрхийлэх, богино зам олох, хэв шинж илрүүлэхэд хамгийн өндөр гүйцэтгэлтэй.<br/>"
        "• <b>Сул тал:</b> Их хэмжээний графыг олон серверт хувааж байрлуулах (sharding) нь техникийн хувьд хамгийн төвөгтэй; "
        "бөөнөөр нь өгөгдөл унших тайлангийн үйлдэлд тохиромж муутай.<br/>"
        "• <b>Түгээмэл хэрэгжүүлэлтүүд:</b> Neo4j, ArangoDB, Amazon Neptune, JanusGraph.<br/>"
        "• <b>Бодит жишээ:</b> Нийгмийн сүлжээний найз санал болгох систем, банкны луйвар илрүүлэх (fraud detection), "
        "бүтээгдэхүүний зөвлөмжийн систем (recommendation engine).",
        s_body
    ))
    story.append(Spacer(1, 6))

    t2_data = [
        [Paragraph("<b>NoSQL төрөл</b>", s_cell_head), Paragraph("<b>Өгөгдлийн бүтэц</b>", s_cell_head), Paragraph("<b>Түгээмэл системүүд</b>", s_cell_head), Paragraph("<b>Хамгийн оновчтой хэрэглээ</b>", s_cell_head)],
        [Paragraph("Key-Value", s_cell), Paragraph("Түлхүүр → Утга (Hash map)", s_cell), Paragraph("Redis, DynamoDB, Memcached", s_cell), Paragraph("Кэшлэлт, хэрэглэгчийн сесс, сагс, хурдан тоолуур", s_cell)],
        [Paragraph("Document", s_cell), Paragraph("JSON / BSON баримт, цуглуулга", s_cell), Paragraph("MongoDB, CouchDB, Firestore", s_cell), Paragraph("Бүтээгдэхүүний каталог, CMS, профайл, API backend", s_cell)],
        [Paragraph("Column-Family", s_cell), Paragraph("Мөрийн түлхүүр → Баганын бүлэг", s_cell), Paragraph("Cassandra, HBase, ScyllaDB", s_cell), Paragraph("IoT мэдрэгч, цагийн цуваа, лог, өндөр бичилттэй систем", s_cell)],
        [Paragraph("Graph", s_cell), Paragraph("Орой (Node) + Ирмэг (Edge)", s_cell), Paragraph("Neo4j, ArangoDB, Neptune", s_cell), Paragraph("Нийгмийн сүлжээ, луйвар илрүүлэх, зөвлөмжийн систем", s_cell)]
    ]
    t2_pdf = Table(t2_data, colWidths=[80, 130, 130, 159])
    t2_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t2_pdf)

    story.append(PageBreak())

    # PAGE 5: CAP Theorem + CA, CP, AP
    story.append(Paragraph("1.3. CAP Теорем — Дэлгэрэнгүй тайлбар", s_h2))
    story.append(Paragraph(
        "Эрик Брюэрийн 2000 онд дэвшүүлж (Brewer, 2000), Сет Гилберт болон Нэнси Линч нарын 2002 онд математикаар баталсан "
        "(Gilbert & Lynch, 2002) CAP теорем нь тархсан өгөгдлийн сангийн архитектур зохиох үндсэн хууль юм. "
        "Энэхүү теорем ёсоор тархсан өгөгдлийн систем нь доорх гурван шинж чанараас нэгэн зэрэг <b>хамгийн ихдээ хоёрыг нь л бүрэн хангаж чадна:</b><br/>"
        "• <b>Consistency (Тогтвортой байдал — C):</b> Кластерийн аль ч зангилаанаас уншилт хийхэд хамгийн сүүлд амжилттай "
        "бичигдсэн өгөгдөл эсвэл алдааны мэдэгдэл буцаах шинж чанар. Өөрөөр хэлбэл бүх хэрэглэгч нэг ижил агшинд ижил өгөгдөл харна.<br/>"
        "• <b>Availability (Хүртээмжтэй байдал — A):</b> Ажиллаж буй зангилаанд ирсэн хүсэлт бүрд (хамгийн сүүлийн өгөгдөл биш "
        "байсан ч) алдаа буцаахгүйгээр боломжит хугацаанд хариу өгөх шинж чанар.<br/>"
        "• <b>Partition Tolerance (Сүлжээний тасралтыг тэсвэрлэх чадвар — P):</b> Кластерийн зангилаануудын хоорондох "
        "сүлжээний холбоо тасрах эсвэл саатах (network partition) үед ч систем бүхэлдээ унахгүйгээр ажиллагаагаа үргэлжлүүлэх чадвар.<br/>"
        "Бодит тархсан системд сүлжээний саатал гарах нь зайлшгүй тул <b>P (Partition Tolerance) нь сонголт бус заавал хангах нөхцөл</b> "
        "байдаг (Brewer, 2012). Иймээс сүлжээний тасралт үүссэн үед архитектурын түвшинд CP эсвэл AP чиглэлийн аль нэгийг "
        "сонгох шаардлагатай болдог:<br/>"
        "• <b>CA (Consistency + Availability):</b> Сүлжээний тасралт гарахгүй нэг зангилаатай орчинд л бүрэн биелнэ. Уламжлалт "
        "дан серверийн RDBMS (PostgreSQL, MySQL) энэ ангилалд хамаардаг.<br/>"
        "• <b>CP (Consistency + Partition Tolerance):</b> Өгөгдлийн нарийвчлал хамгийн чухал үед сонгоно. Сүлжээ тасрахад "
        "зөрүүтэй өгөгдөл гаргахгүйн тулд зарим зангилааны бичилт эсвэл уншилтыг түр хаадаг. Жишээ нь: банкны үлдэгдэл, барааны "
        "нөөцийн захиалга. <i>Бодит системүүд:</i> MongoDB (анхдагч тохиргоогоор), HBase, Redis (хатуу тохиргоо), Neo4j.<br/>"
        "• <b>AP (Availability + Partition Tolerance):</b> Систем хэзээ ч зогсохгүй байх нь түр зуурын өгөгдлийн зөрүүнээс илүү "
        "чухал үед сонгоно. Сүлжээ сэргэхэд зангилаанууд өгөгдлөө ижилсүүлдэг (Eventual Consistency). Жишээ нь: нийгмийн сүлжээний пост, "
        "барааны сэтгэгдэл, үзэлтийн түүх. <i>Бодит системүүд:</i> Apache Cassandra, Amazon DynamoDB, CouchDB.",
        s_body
    ))
    story.append(Spacer(1, 6))

    t3_data = [
        [Paragraph("<b>Ангилал</b>", s_cell_head), Paragraph("<b>Сүлжээ тасрах үеийн төлөв</b>", s_cell_head), Paragraph("<b>Жишээ системүүд</b>", s_cell_head), Paragraph("<b>Сонгох бизнесийн нөхцөл</b>", s_cell_head)],
        [Paragraph("CA", s_cell), Paragraph("Тархсан сүлжээний тасралтыг тэсвэрлэхгүй (дан сервер)", s_cell), Paragraph("PostgreSQL, MySQL, Oracle (single-node)", s_cell), Paragraph("Дотоод байгууллагын систем, хатуу ACID шаардлага", s_cell)],
        [Paragraph("CP", s_cell), Paragraph("Тогтвортой байдлыг хадгалж, салангид зангилааг түр хязгаарлана", s_cell), Paragraph("MongoDB, HBase, Neo4j, ZooKeeper", s_cell), Paragraph("Төлбөр тооцоо, нөөцийн удирдлага, хэрэглэгчийн эрх", s_cell)],
        [Paragraph("AP", s_cell), Paragraph("Бүх хүсэлтэд хариулж, сүүлд нь өгөгдлөө ижилсүүлнэ", s_cell), Paragraph("Cassandra, DynamoDB, CouchDB, Riak", s_cell), Paragraph("Сошиал контент, аналитик, каталог, дэлхийн сервис", s_cell)]
    ]
    t3_pdf = Table(t3_data, colWidths=[55, 155, 140, 149])
    t3_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t3_pdf)

    story.append(PageBreak())

    # PAGE 6: 1.4 NoSQL Data Modeling
    story.append(Paragraph("1.4. NoSQL-ийн өгөгдлийн загварчлал", s_h2))
    story.append(Paragraph(
        "Хамааралт өгөгдлийн санд өгөгдлийг эхлээд нормалчилж (3NF), дараа нь хэрэгцээтэй хүсэлтээ бичдэг бол NoSQL-д "
        "<b>хэрэглээний хүсэлтийн хэв маягт (Query-Driven Modeling)</b> тулгуурлан өгөгдлийн бүтцээ төлөвлөдөг (MongoDB Manual). "
        "Гол ойлголтууд нь:<br/>"
        "• <b>Schema-less буюу динамик схем:</b> Өгөгдлийн сан хатуу бүтэц тулгахгүй гэсэн үг боловч програм хангамжийн "
        "түвшинд логик бүтцийг (schema validation) мөрдөх нь чухал. Нэг цуглуулга дахь хоёр бүтээгдэхүүн өөр өөр техникийн "
        "үзүүлэлттэй байж болно (жишээ нь зөөврийн компьютерт cpu, ram байхад хувцаст size, material байна).<br/>"
        "• <b>Documents vs Collections:</b> Document Store-д Document нь SQL-ийн мөртэй, Collection нь хүснэгттэй дүйцнэ. "
        "Гэхдээ баримт нь дотроо шаталсан дэд объект болон массив агуулдаг тул SQL-ийн олон хүснэгтийн мэдээллийг нэг цул "
        "нэгжид багтаах чадвартай.<br/>"
        "• <b>Wide-Column загвар:</b> Cassandra зэрэг системд өгөгдлийг Partition Key болон Clustering Key-ээр зохион байгуулдаг. "
        "Partition Key нь өгөгдөл аль зангилаанд хадгалагдахыг, Clustering Key нь тухайн хэсэг дотор ямар дарааллаар "
        "эрэмбэлэгдэхийг тодорхойлно.<br/>"
        "• <b>Relationship Modeling — Embedding vs Referencing:</b><br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;о <i>Embedding (Шигтгэх / Denormalization):</i> Холбоотой өгөгдлийг үндсэн баримтын дотор дэд "
        "объект эсвэл массив хэлбэрээр хадгалах (жишээ нь бүтээгдэхүүний дотор үзүүлэлт specs болон сүүлийн сэтгэгдлүүдийг шигтгэх). "
        "Нэг уншилтаар бүх мэдээллийг хурдан авах давуу талтай боловч баримтын хэмжээ хэт томрох (MongoDB-д 16MB хязгаартай) "
        "эсвэл өгөгдөл олон газар давтагдах сул талтай.<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;о <i>Referencing (Холбоос хийх / Normalization):</i> Хоёр цуглуулгыг _id талбараар холбох "
        "(жишээ нь захиалга нь product_id-аар бүтээгдэхүүн рүү заах). Олон-олон (M:N) хамаарал болон байнга өөрчлөгддөг "
        "бие даасан өгөгдөлд тохиромжтой.<br/>"
        "• <b>Indexes (Индексжүүлэлт):</b> Бүх баримтыг дараалан шалгах (Collection Scan) үйлдлээс зайлсхийж B-Tree бүтцээр "
        "хайлыг хурдасгана. MongoDB нь дан талбарын (Single Field), нийлмэл (Compound), массивын (Multikey), текст хайлтын (Text) "
        "болон газарзүйн (Geospatial) индексийг дэмждэг.<br/>"
        "• <b>Sharding ба Replication:</b> Replication нь өгөгдлийн хуулбарыг олон серверт (Replica Set: Primary + Secondary) "
        "хадгалснаар найдвартай ажиллагаа (High Availability) болон уншилтын ачаалал хуваах боломжийг олгодог. Sharding нь "
        "их хэмжээний өгөгдлийг Shard Key-ийн дагуу хэсэглэн олон кластерт хувааж, бичилт болон хадгалалтын багтаамжийг "
        "хэвтээ чиглэлд хязгааргүй өргөтгөх механизм юм (Kleppmann, 2017).",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 7: 1.5 Netflix + 2.1 System Purpose
    story.append(Paragraph("1.5. Том системүүдийн хэрэгжүүлэлт: Netflix ба Apache Cassandra", s_h2))
    story.append(Paragraph(
        "Дэлхийн тэргүүлэх стриминг платформ болох Netflix нь 200 гаруй сая хэрэглэгчийн үзэлтийн түүх, хаана зогсоосон "
        "хугацаа (bookmark), зөвлөмжийн өгөгдөл болон төхөөрөмжийн бүртгэлийг бодит цагт тасралтгүй боловсруулдаг. "
        "Анх уламжлалт хамааралт өгөгдлийн сан (Oracle) ашиглаж байсан боловч дата төвийн саатал болон өсөн нэмэгдэж буй "
        "трафикийг даах боломжгүй болсон тул 2011 оноос үүлэн орчин (AWS) болон Apache Cassandra (Column-Family NoSQL) руу "
        "бүрэн шилжсэн (Cockcroft & Sheahan, 2011).<br/>"
        "Netflix-ийн хувьд Cassandra-г сонгосон гол шалтгаанууд нь: (1) <i>Masterless архитектур</i> — бүх зангилаа ижил "
        "эрхтэй тул нэг бүс нутгийн дата төв унасан ч хэрэглэгч кино үзэх үйлчилгээ огт тасрахгүй; (2) <i>Олон бүс нутгийн "
        "хуулбарлалт (Multi-region Replication)</i> — дэлхийн өнцөг булан бүрт байрлах хэрэглэгчдэд хамгийн ойрын серверээс "
        "өгөгдөл хүргэх; (3) <i>Шугаман хэвтээ өргөтгөлт</i> — AWS дээр хийсэн туршилтаар 288 зангилаатай Cassandra кластер "
        "секундэд 1.1 саяас дээш бичилт (write/sec) хийх чадвартай болох нь батлагдсан (Cockcroft & Sheahan, 2011). "
        "CAP теоремийн үүднээс Netflix нь хэрэглэгчийн үзэлтийн түүхийг AP зарчмаар ажиллуулдаг нь бизнесийн зорилгодоо "
        "бүрэн нийцсэн шийдэл юм.",
        s_body
    ))
    story.append(Spacer(1, 8))

    story.append(Paragraph("2. Практик туршилтын төсөл: Цахим худалдааны бүтээгдэхүүний каталог (MongoDB)", s_h1))
    story.append(Paragraph("2.1. Системийн зорилго ба өгөгдлийн тодорхойлолт", s_h2))
    story.append(Paragraph(
        "Туршилтын төслийн хүрээнд MongoDB (Document Store) өгөгдлийн сан болон Python (PyMongo) програмчлалын хэлийг "
        "ашиглан Цахим худалдааны бүтээгдэхүүний каталогийн систем (E-Commerce Product Catalog)-ийг зохион бүтээв. "
        "Цахим худалдааны платформд электрон бараа, гэр ахуйн цахилгаан хэрэгсэл, дагалдах хэрэгсэл, оффисын тавилга "
        "зэрэг олон төрлийн ангиллын бүтээгдэхүүн нэг дор байршдаг. Ангилал бүр ялгаатай техникийн үзүүлэлттэй "
        "(жишээ нь зөөврийн компьютерт процессор, RAM; чихэвчинд Bluetooth хувилбар, батарей; дэлгэцэнд нягтрал) "
        "тул хатуу хүснэгтийн схемтэй SQL-ээс илүүтэйгээр уян хатан JSON баримт хадгалдаг MongoDB хамгийн оновчтой сонголт юм.<br/>"
        "Системд дараах гурван үндсэн төрлийн өгөгдлийг боловсруулна:<br/>"
        "1. <b>Бүтээгдэхүүний мэдээлэл (products collection):</b> Бүтээгдэхүүний код, нэр, ангилал, брэнд, үнэ (төгрөгөөр), "
        "үлдэгдэл тоо ширхэг, динамик техникийн үзүүлэлт (specs дэд объект), хайлтын түлхүүр үгс (tags массив), хэрэглэгчдийн "
        "сүүлийн үнэлгээнүүд (reviews шигтгэсэн массив) болон үүсгэсэн огноо.<br/>"
        "2. <b>Ангиллын лавлах (categories collection):</b> Бүтээгдэхүүний үндсэн ангиллын нэр, тайлбар, эцэг ангиллын код.<br/>"
        "3. <b>Захиалгын бүртгэл (orders collection):</b> Хэрэглэгчийн нэр, захиалсан бүтээгдэхүүнүүдийн жагсаалт (product_id "
        "холбоос, тоо ширхэг, нэгж үнэ), нийт дүн болон төлөв.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 8: 2.2 Data Model Diagram + JSON Schema
    story.append(Paragraph("2.2. Өгөгдлийн загвар ба архитектурын зураглал", s_h2))
    story.append(Paragraph(
        "Өгөгдлийн загварыг боловсруулахдаа эрлийз загварчлал (Hybrid Modeling: Embedding + Referencing)-ыг ашиглав. "
        "Тухайлбал, бүтээгдэхүүний техникийн үзүүлэлт (specs) болон хэрэглэгчийн сэтгэгдлийг (reviews) бүтээгдэхүүний "
        "баримт дотор шууд шигтгэснээр (Embedding) бүтээгдэхүүний дэлгэрэнгүй хуудсыг ганц уншилтаар нээх боломжтой болсон. "
        "Харин ангилал болон захиалгын бүртгэлийг тусдаа цуглуулгад байршуулж, category_code болон product_id талбараар "
        "холбосон (Referencing).",
        s_body
    ))
    story.append(Spacer(1, 4))

    # Architecture Box ASCII
    model_box_text = """+-------------------------+      +-------------------------------------------+      +-------------------------+
|       categories        |      |         products (Үндсэн баримт)          |      |         orders          |
+-------------------------+      +-------------------------------------------+      +-------------------------+
| _id: ObjectId (PK)      |<1:N--| _id: ObjectId (PK)                        |--M:N>| _id: ObjectId (PK)      |
| category_code: String   |      | product_id: Int (Unique)                  |      | order_no: String        |
| name: String            |      | name, brand: String                       |      | customer: String        |
| description: String     |      | category_code: String (Ref)               |      | items: [{product_id}]   |
+-------------------------+      | price: Int, stock: Int, tags: [String]    |      | total_price: Int        |
                                 | specs: { cpu, ram, storage } (Embedded)   |      | status: String          |
                                 | reviews: [ { user, rating } ] (Embedded)  |      +-------------------------+
                                 +-------------------------------------------+"""
    story.append(Preformatted(model_box_text, s_code))
    story.append(Paragraph("<i>Зураг 1. Цахим худалдааны каталогийн логик бүтэц (Embedding ба Referencing хосолсон загвар)</i>", s_cell))
    story.append(Spacer(1, 6))

    story.append(Paragraph("Доорх JSON бүтэц нь products цуглуулгад хадгалагдах бүтээгдэхүүний баримтын загвар схем болно:", s_body))
    json_sample = """{
  "product_id": 101,
  "name": "Lenovo ThinkPad X1 Carbon Gen 11",
  "brand": "Lenovo",
  "category_code": "LAPTOP",
  "price": 5450000,
  "stock": 14,
  "specs": {
    "cpu": "Intel Core i7-1365U",
    "ram": "16 GB LPDDR5",
    "storage": "512 GB NVMe SSD",
    "screen": "14 inch WUXGA"
  },
  "tags": ["laptop", "business", "ultrabook", "intel"],
  "reviews": [
    { "user": "Bat-Erdene", "rating": 5, "comment": "Маш хөнгөн, батарей удаан барьдаг." },
    { "user": "Sarnai", "rating": 4, "comment": "Гүйцэтгэл сайн боловч үнэ бага зэрэг өндөр." }
  ],
  "created_at": "2026-10-15T09:30:00Z"
}"""
    story.append(Preformatted(json_sample, s_code))

    story.append(PageBreak())

    # PAGE 9: Data Flow + 2.3 Setup + Products Table (Part 1)
    df_box_text = """+---------------------+       +--------------------+       +----------------------+       +--------------------+
|   Хэрэглэгч / CLI   |  ==>  |  Python (PyMongo)  |  ==>  |    MongoDB Engine    |  ==>  |   Үр дүн / Тайлан  |
| Хайлт, CRUD хүсэлт, |       |  JSON хөрвүүлэлт,  |       | • Text/Compound Idx  |       |  Шүүсэн жагсаалт,  |
| Тайлангийн комманд  |       | Баталгаажуулалт    |       | • Pipeline, Storage  |       | Статистик нэгтгэл  |
+---------------------+       +--------------------+       +----------------------+       +--------------------+"""
    story.append(Preformatted(df_box_text, s_code))
    story.append(Paragraph("<i>Зураг 2. Хэрэглэгчийн хүсэлтээс өгөгдлийн сангийн боловсруулалт хүртэлх өгөгдлийн урсгал</i>", s_cell))
    story.append(Spacer(1, 8))

    story.append(Paragraph("2.3. Өгөгдлийн сан байгуулах ба өгөгдөл оруулах (25 баримт)", s_h2))
    story.append(Paragraph(
        "Туршилтын өгөгдлийн санг ecommerce_catalog нэртэйгээр үүсгэж, дотор нь products цуглуулгад 5 өөр ангиллын "
        "(LAPTOP, SMARTPHONE, AUDIO, ACCESSORY, MONITOR) нийт 25 ширхэг бүтээгдэхүүний баримтыг бэлтгэн оруулав. "
        "Ангилал тус бүр өөрийн онцлогт тохирсон specs бүтэцтэй байгаа нь NoSQL-ийн уян хатан схемийн давуу талыг харуулж байна.",
        s_body
    ))
    story.append(Spacer(1, 4))

    # Table of products 101-110
    t_p_part1 = [
        [Paragraph("<b>ID</b>", s_cell_head), Paragraph("<b>Бүтээгдэхүүний нэр</b>", s_cell_head), Paragraph("<b>Ангилал</b>", s_cell_head), Paragraph("<b>Үнэ (₮)</b>", s_cell_head), Paragraph("<b>Үлдэгдэл</b>", s_cell_head), Paragraph("<b>Онцлог талбар</b>", s_cell_head)]
    ]
    for r in PROD_ROWS[:10]:
        t_p_part1.append([Paragraph(str(r[0]), s_cell), Paragraph(str(r[1]), s_cell), Paragraph(str(r[2]), s_cell), Paragraph(str(r[3]), s_cell), Paragraph(str(r[4]), s_cell), Paragraph(str(r[5]), s_cell)])
    t_p1_pdf = Table(t_p_part1, colWidths=[28, 145, 75, 65, 42, 144])
    t_p1_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(t_p1_pdf)

    story.append(PageBreak())

    # PAGE 10: Products Table (Part 2: 111-125)
    t_p_part2 = [
        [Paragraph("<b>ID</b>", s_cell_head), Paragraph("<b>Бүтээгдэхүүний нэр</b>", s_cell_head), Paragraph("<b>Ангилал</b>", s_cell_head), Paragraph("<b>Үнэ (₮)</b>", s_cell_head), Paragraph("<b>Үлдэгдэл</b>", s_cell_head), Paragraph("<b>Онцлог талбар</b>", s_cell_head)]
    ]
    for r in PROD_ROWS[10:]:
        t_p_part2.append([Paragraph(str(r[0]), s_cell), Paragraph(str(r[1]), s_cell), Paragraph(str(r[2]), s_cell), Paragraph(str(r[3]), s_cell), Paragraph(str(r[4]), s_cell), Paragraph(str(r[5]), s_cell)])
    t_p2_pdf = Table(t_p_part2, colWidths=[28, 145, 75, 65, 42, 144])
    t_p2_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(t_p2_pdf)
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        "Дараах Python скрипт нь MongoDB сервертэй холбогдож, 25 баримтыг insert_many() функцээр бөөнөөр нь оруулах "
        "болон өвөрмөц түлхүүрийн индекс үүсгэх үйлдлийг гүйцэтгэнэ:",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 11: init_db.py code
    story.append(Paragraph("<b>init_db.py — Өгөгдлийн сан байгуулах ба 25 баримт оруулах</b>", s_h3))
    init_code = """# init_db.py — Өгөгдлийн сан байгуулах ба 25 баримт оруулах
from pymongo import MongoClient
from datetime import datetime, timezone

client = MongoClient("mongodb://localhost:27017/")
db = client["ecommerce_catalog"]
products = db["products"]

# Хуучин туршилтын өгөгдлийг цэвэрлэх
products.drop()

# product_id талбар дээр давхардалгүй индекс үүсгэх
products.create_index("product_id", unique=True)

sample_products = [
    {
        "product_id": 101, "name": "Lenovo ThinkPad X1 Carbon Gen 11", "brand": "Lenovo",
        "category_code": "LAPTOP", "price": 5450000, "stock": 14,
        "specs": {"cpu": "Intel Core i7-1365U", "ram": "16GB", "storage": "512GB SSD"},
        "tags": ["laptop", "business", "ultrabook"],
        "reviews": [{"user": "Bat", "rating": 5, "comment": "Маш хөнгөн, найдвартай."}],
        "created_at": datetime.now(timezone.utc)
    },
    {
        "product_id": 106, "name": "iPhone 15 Pro 256GB", "brand": "Apple",
        "category_code": "SMARTPHONE", "price": 4350000, "stock": 18,
        "specs": {"chip": "A17 Pro", "storage": "256GB", "body": "Titanium"},
        "tags": ["smartphone", "apple", "5g", "camera"],
        "reviews": [{"user": "Tuya", "rating": 5, "comment": "Камер маш тод."}],
        "created_at": datetime.now(timezone.utc)
    },
    {
        "product_id": 111, "name": "Sony WH-1000XM5", "brand": "Sony",
        "category_code": "AUDIO", "price": 1280000, "stock": 30,
        "specs": {"type": "Over-Ear", "anc": True, "battery_hours": 30},
        "tags": ["audio", "headphone", "anc", "bluetooth"],
        "reviews": [{"user": "Bold", "rating": 5, "comment": "Дуу тусгаарлалт гайхалтай."}],
        "created_at": datetime.now(timezone.utc)
    }
    # ... (102-125 хүртэлх үлдсэн 22 баримт JSON файлаас ачаалагдана)
]

result = products.insert_many(sample_products)
print(f"Амжилттай оруулсан баримтын тоо: {len(result.inserted_ids)}")"""
    story.append(Preformatted(init_code, s_code))

    story.append(PageBreak())

    # PAGE 12: 2.4 CRUD Code
    story.append(Paragraph("2.4. CRUD үйлдлүүд (Python / PyMongo)", s_h2))
    story.append(Paragraph(
        "Өгөгдлийн сангийн үндсэн дөрвөн үйлдэл болох Create (Үүсгэх), Read (Унших), Update (Шинэчлэх), "
        "Delete (Устгах)-ийг MongoDB-ийн онцлог операторууд ($gte, $lte, $set, $inc, $push) ашиглан хэрэгжүүлсэн кодыг үзүүлэв:",
        s_body
    ))
    crud_code = """# crud_operations.py — Бүрэн CRUD үйлдлүүд
from pymongo import MongoClient

client = MongoClient("mongodb://localhost:27017/")
products = client["ecommerce_catalog"]["products"]

# 1. CREATE: Шинэ бүтээгдэхүүн нэмэх
new_item = {
    "product_id": 126,
    "name": "Logitech C920s Pro HD Webcam",
    "brand": "Logitech",
    "category_code": "ACCESSORY",
    "price": 265000,
    "stock": 18,
    "specs": {"resolution": "1080p 30fps", "mic": "Dual Stereo"},
    "tags": ["webcam", "streaming", "office"],
    "reviews": []
}
ins = products.insert_one(new_item)
print("[CREATE] Нэмэгдсэн _id:", ins.inserted_id)

# 2. READ: Нөхцөлт шүүлтүүр ба эрэмбэлэлт (2 сая - 5 сая төгрөгийн LAPTOP)
query = {
    "category_code": "LAPTOP",
    "price": {"$gte": 2000000, "$lte": 5000000}
}
projection = {"_id": 0, "product_id": 1, "name": 1, "price": 1, "stock": 1}
print("\\n[READ] Шүүсэн зөөврийн компьютерууд:")
for doc in products.find(query, projection).sort("price", 1):
    print(" ", doc)

# 3. UPDATE: Үнэ шинэчлэх, нөөц хорогдуулах ба шинэ сэтгэгдэл нэмэх
upd = products.update_one(
    {"product_id": 126},
    {
        "$set": {"price": 249000},
        "$inc": {"stock": -2},
        "$push": {
            "reviews": {
                "user": "Anar", "rating": 5, "comment": "Онлайн хичээлд тохиромжтой."
            }
        }
    }
)
print(f"\\n[UPDATE] Шинэчлэгдсэн баримт: {upd.modified_count}")

# 4. DELETE: Туршилтын бүтээгдэхүүнийг устгах
deleted = products.delete_one({"product_id": 126})
print(f"\\n[DELETE] Устгагдсан баримт: {deleted.deleted_count}")"""
    story.append(Preformatted(crud_code, s_code))

    story.append(PageBreak())

    # PAGE 13: Terminal Output of CRUD + 2.5 Advanced Intro
    story.append(Paragraph("<b>Хүлээгдэж буй терминалын гаралт:</b>", s_h3))
    term_crud = """[CREATE] Нэмэгдсэн _id: 673c41a9f2d8e4b1c091a201

[READ] Шүүсэн зөөврийн компьютерууд:
  {'product_id': 105, 'name': 'HP Pavilion 15', 'price': 2450000, 'stock': 25}
  {'product_id': 102, 'name': 'Apple MacBook Air M3', 'price': 4600000, 'stock': 20}

[UPDATE] Шинэчлэгдсэн баримт: 1
[DELETE] Устгагдсан баримт: 1"""
    story.append(Preformatted(term_crud, s_code))
    story.append(Spacer(1, 10))

    story.append(Paragraph("2.5. Дэвшилтэт ашиглалт: Index, Text Search ба Aggregation Pipeline", s_h2))
    story.append(Paragraph(
        "Цахим худалдааны бодит системд бүтээгдэхүүн хайх болон удирдлагын тайлан гаргах ажиллагаа өндөр гүйцэтгэл шаарддаг. "
        "Иймд MongoDB-ийн гурван дэвшилтэт боломжийг хослуулан хэрэгжүүлэв:<br/>"
        "1. <b>Нийлмэл индекс (Compound Index):</b> Ангилал болон үнээр шүүх хайлтыг хурдасгах зорилгоор (category_code: 1, price: -1) индекс үүсгэв.<br/>"
        "2. <b>Бүрэн текст хайлт (Text Index & Text Search):</b> Бүтээгдэхүүний name, brand, tags талбарууд дээр текст индекс байгуулж, "
        "хэрэглэгчийн оруулсан түлхүүр үгээр хамаарлын оноо (textScore) тооцон эрэмбэлэв.<br/>"
        "3. <b>Aggregation Pipeline:</b> Ангилал тус бүрийн нийт бүтээгдэхүүний тоо, дундаж үнэ, хамгийн үнэтэй бараа болон "
        "нийт агуулахын үлдэгдлийн мөнгөн дүнг нэгтгэн тооцоолов.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 14: advanced_features.py code
    story.append(Paragraph("<b>advanced_features.py — Индекс, Текст хайлт ба Aggregation Pipeline</b>", s_h3))
    adv_code = """# advanced_features.py — Индекс, Текст хайлт ба Aggregation Pipeline
from pymongo import MongoClient, ASCENDING, DESCENDING, TEXT

client = MongoClient("mongodb://localhost:27017/")
products = client["ecommerce_catalog"]["products"]

# 1. Compound Index болон Text Index үүсгэх
products.create_index([("category_code", ASCENDING), ("price", DESCENDING)])
products.create_index([("name", TEXT), ("brand", TEXT), ("tags", TEXT)])

# 2. Text Search — "apple" эсвэл "laptop" үгээр хамаарлын оноогоор хайх
print("=== TEXT SEARCH ҮР ДҮН ('apple laptop') ===")
text_results = products.find(
    {"$text": {"$search": "apple laptop"}},
    {"_id": 0, "name": 1, "price": 1, "score": {"$meta": "textScore"}}
).sort([("score", {"$meta": "textScore"})]).limit(4)

for item in text_results:
    print(f" {item['name']} | {item['price']:,}₮ | Оноо: {item['score']:.2f}")

# 3. Aggregation Pipeline — Ангилал бүрийн статистик тайлан
pipeline = [
    {"$match": {"stock": {"$gt": 0}}},
    {
        "$group": {
            "_id": "$category_code",
            "product_count": {"$sum": 1},
            "avg_price": {"$avg": "$price"},
            "max_price": {"$max": "$price"},
            "total_inventory_value": {
                "$sum": {"$multiply": ["$price", "$stock"]}
            }
        }
    },
    {"$sort": {"total_inventory_value": -1}}
]

print("\\n=== AGGREGATION PIPELINE: АНГИЛЛЫН ТАЙЛАН ===")
for row in products.aggregate(pipeline):
    print(
        f" {row['_id']:10s} | Тоо: {row['product_count']:2d} | "
        f"Дундаж үнэ: {row['avg_price']:10,.0f}₮ | "
        f"Агуулахын нийт дүн: {row['total_inventory_value']:13,.0f}₮"
    )"""
    story.append(Preformatted(adv_code, s_code))

    story.append(PageBreak())

    # PAGE 15: Aggregation Table + Section 2.6 Intro + Screenshots 1 & 2
    story.append(Paragraph("<b>Aggregation Pipeline-ийн тооцоолсон үр дүнгийн хүснэгт (25 баримтын өгөгдлөөр):</b>", s_h3))
    t_agg_p_data = [
        [Paragraph("<b>Ангилал</b>", s_cell_head), Paragraph("<b>Барааны тоо</b>", s_cell_head), Paragraph("<b>Дундаж үнэ (₮)</b>", s_cell_head), Paragraph("<b>Хамгийн өндөр үнэ</b>", s_cell_head), Paragraph("<b>Агуулахын нийт дүн (₮)</b>", s_cell_head)]
    ]
    for r in AGG_ROWS:
        t_agg_p_data.append([Paragraph(r[0], s_cell), Paragraph(r[1], s_cell), Paragraph(r[2], s_cell), Paragraph(r[3], s_cell), Paragraph(r[4], s_cell)])
    t_agg_p = Table(t_agg_p_data, colWidths=[90, 75, 110, 110, 114])
    t_agg_p.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B3C5D')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DCE0E5')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E9ECEF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_agg_p)
    story.append(Spacer(1, 8))

    story.append(Paragraph("2.6. Ажлын дэлгэцийн зургийн хэсэг (Туршилтын баталгаажуулалт)", s_h2))
    story.append(Paragraph(
        "Бие даалтын ажлын хүрээнд бодит систем дээр туршилтыг ажиллуулан баталгаажуулсан дэлгэцийн зургуудыг (screenshots) "
        "дор хавсаргав:",
        s_body
    ))
    story.append(Spacer(1, 4))

    # Screenshot 1
    story.append(Paragraph("<b>[Дэлгэцийн зураг 1 — Өгөгдлийн сангийн бүтэц (Database Structure)]</b>", s_h3))
    story.append(Paragraph("<i>MongoDB Compass дээрх ecommerce_catalog өгөгдлийн сан, products цуглуулга ба 25 баримтын харагдац:</i>", s_cell))
    story.append(Image(os.path.join(BASE_DIR, "screenshot1_database_structure.png"), width=490, height=265))

    story.append(PageBreak())

    # PAGE 16: Screenshots 2 & 3
    story.append(Paragraph("<b>[Дэлгэцийн зураг 2 — CRUD үйлдлийн гүйцэтгэл]</b>", s_h3))
    story.append(Paragraph("<i>python crud_operations.py ажиллуулсны дараах Create, Read, Update, Delete үр дүнгийн терминалын зураг:</i>", s_cell))
    story.append(Image(os.path.join(BASE_DIR, "screenshot2_crud_operations.png"), width=490, height=245))
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>[Дэлгэцийн зураг 3 — Индекс ба Текст хайлтын үр дүн]</b>", s_h3))
    story.append(Paragraph("<i>MongoDB Compass-ийн Indexes таб (Compound ба Text Index) болон $text хайлтын хариу:</i>", s_cell))
    story.append(Image(os.path.join(BASE_DIR, "screenshot3_indexes_text_search.png"), width=490, height=255))

    story.append(PageBreak())

    # PAGE 17: Screenshot 4 + Section 3 Intro
    story.append(Paragraph("<b>[Дэлгэцийн зураг 4 — Aggregation Pipeline-ийн гүйцэтгэл]</b>", s_h3))
    story.append(Paragraph("<i>Ангилал бүрийн дундаж үнэ болон агуулахын нийт дүнг тооцсон Aggregation үр дүнгийн терминалын зураг:</i>", s_cell))
    story.append(Image(os.path.join(BASE_DIR, "screenshot4_aggregation_pipeline.png"), width=490, height=245))
    story.append(Spacer(1, 10))

    story.append(Paragraph("3. Дүн шинжилгээ ба Харьцуулалт", s_h1))
    story.append(Paragraph("3.1. MongoDB нь SQL-тэй харьцуулахад ямар давуу талтай байв?", s_h2))
    story.append(Paragraph(
        "Туршилтын төслийг хэрэгжүүлэх явцад MongoDB нь уламжлалт хамааралт өгөгдлийн сантай (PostgreSQL/MySQL) харьцуулахад "
        "дараах тодорхой давуу талуудыг харуулсан:<br/>"
        "• <b>Олон төрлийн бүтээгдэхүүний үзүүлэлтийг хялбар хадгалах:</b> Хэрэв SQL ашигласан бол ангилал бүрт тусдаа хүснэгт "
        "үүсгэх эсвэл EAV (Entity-Attribute-Value) гэх мэт нарийн төвөгтэй загвар ашиглах шаардлагатай байв. MongoDB-д specs дэд "
        "объектыг дурын талбартайгаар шууд хадгалснаар хөгжүүлэлтийн хугацаа мэдэгдэхүйц багассан.<br/>"
        "• <b>Уншилтын өндөр хурд ба JOIN-гүй бүтэц:</b> Бүтээгдэхүүний үндсэн мэдээлэл, үзүүлэлт, таг, сэтгэгдэл бүгд нэг "
        "баримтад шигтгэгдсэн тул нэг удаагийн find_one() үйлдлээр бүх өгөгдлийг шууд авч чадсан.<br/>"
        "• <b>Хөгжүүлэгчид ээлтэй JSON бүтэц:</b> Python-ийн dictionary бүтэц MongoDB-ийн BSON баримттай шууд дүйцдэг тул "
        "хүснэгт-объект хөрвүүлэлтийн (ORM) илүүц код бичих шаардлагагүй байв.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 18: 3.2, 3.3, 3.4
    story.append(Paragraph("3.2. Ямар сул тал ажиглагдав?", s_h2))
    story.append(Paragraph(
        "• <b>Өгөгдлийн бүрэн бүтэн байдлын хяналт сул:</b> SQL шиг гадаад түлхүүрийн (Foreign Key) хатуу хязгаарлалт "
        "анхдагчаар байхгүй тул байхгүй category_code оруулах алдаанаас сэргийлэхийн тулд програмын код эсвэл JSON Schema "
        "Validation нэмэлтээр бичих шаардлага гарсан.<br/>"
        "• <b>Өгөгдлийн давхардал ба баримтын хэмжээ:</b> Хэрэв нэг бүтээгдэхүүнд олон мянган хэрэглэгчийн сэтгэгдэл "
        "хуримтлагдвал reviews массив хэт томорч, MongoDB-ийн нэг баримтын 16MB хязгаарт тулах эрсдэлтэй тул томоохон системд "
        "сэтгэгдлийг тусдаа цуглуулга болгон салгах шаардлагатай.<br/>"
        "• <b>Санах ойн хэрэглээ:</b> Баримт бүрд талбарын нэрс (product_id, name, price гэх мэт) давтагдан хадгалагддаг тул "
        "ижил хэмжээний өгөгдөл SQL-ээс илүү диск болон RAM зарцуулах хандлагатай.",
        s_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("3.3. Өөр хэрэгцээнд илүү тохиромжтой NoSQL төрлүүд", s_h2))
    story.append(Paragraph(
        "Орчин үеийн томоохон системүүд зөвхөн ганц төрлийн өгөгдлийн санд баригдахгүйгээр Polyglot Persistence буюу "
        "даалгавар бүрт хамгийн тохиромжтой өгөгдлийн санг хослуулан ашигладаг (Sadalage & Fowler, 2012). Манай цахим худалдааны "
        "системийг өргөжүүлэх тохиолдолд:<br/>"
        "• <b>Redis (Key-Value):</b> Хамгийн их хандалттай барааны хуудсыг кэшлэх, хэрэглэгчийн сагс (shopping cart) болон "
        "хямдралын цаг тоолуурыг миллисекундын хурдаар боловсруулахад MongoDB-ээс илүү тохиромжтой.<br/>"
        "• <b>Neo4j (Graph):</b> 'Энэ барааг авсан хэрэглэгчид мөн дараах барааг сонирхсон' гэх зөвлөмжийн систем "
        "(Recommendation Engine) байгуулахад граф өгөгдлийн сан хамгийн оновчтой.<br/>"
        "• <b>Apache Cassandra (Column-Family):</b> Хэрэглэгчдийн бүх товшилт (clickstream) болон серверийн лог өгөгдлийг "
        "тасралтгүй өндөр хурдаар бичиж хадгалахад давуу талтай.",
        s_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("3.4. Сургалтын явцад хамгийн чухал гэж үзсэн ойлголт", s_h2))
    story.append(Paragraph(
        "Энэхүү бие даалтын ажлын явцад хамгийн чухал ойлголт нь <b>'Төгс өгөгдлийн сан гэж байдаггүй; өгөгдлийн загвар болон "
        "CAP-ийн сонголт нь системийн бизнесийн шаардлага, уншилт-бичилтийн харьцаанаас шууд хамаардаг'</b> гэх зарчим байв. "
        "Ялангуяа NoSQL-д хүснэгтээ эхлээд нормалчлах бус, системд ямар хүсэлт хамгийн их давтагдах вэ гэдгээс эхлэн "
        "Embedding vs Referencing болон индексийн стратегиа зөв төлөвлөх нь гүйцэтгэлийг хэд дахин нэмэгдүүлдэг болохыг "
        "практик дээр баттай ойлгов.",
        s_body
    ))

    story.append(PageBreak())

    # PAGE 19: Section 4 + References
    story.append(Paragraph("4. Дүгнэлт ба Санал зөвлөмж", s_h1))
    story.append(Paragraph(
        "<b>Ерөнхий дүгнэлт:</b> Энэхүү бие даалтын ажлаар NoSQL өгөгдлийн сангуудын үүссэн шалтгаан, дөрвөн үндсэн төрлийн "
        "архитектурын онцлог, CAP теоремийн практик хэрэглээ, өгөгдөл загварчлалын арга зүй болон Netflix-ийн Apache Cassandra "
        "ашигласан бодит туршлагыг онолын түвшинд бүрэн судалж нэгтгэв. Улмаар MongoDB болон Python (PyMongo) ашиглан "
        "25 бүтээгдэхүүний баримт бүхий цахим худалдааны каталогийн туршилтын системийг амжилттай бүтээж, бүрэн CRUD үйлдэл, "
        "нийлмэл болон текст индекс, мөн Aggregation Pipeline-ийн нэгтгэлийн тайланг туршин баталгаажууллаа.<br/>"
        "<b>Ирээдүйд өргөжүүлэх боломж:</b> Туршилтын системийг цаашид дараах чиглэлээр өргөжүүлэх боломжтой: (1) FastAPI эсвэл "
        "Node.js ашиглан REST API давхарга нэмж веб интерфэйстэй холбох; (2) MongoDB-ийн $jsonSchema баталгаажуулалт нэмж "
        "өгөгдлийн чанарыг автоматаар хянах; (3) Олон хэрэглэгчийн захиалгын үед барааны үлдэгдэл буруу хасагдахаас сэргийлж "
        "MongoDB-ийн Multi-Document ACID Transaction нэвтрүүлэх; (4) Redis кэш давхарга холбож хайлтын хурдыг нэмэгдүүлэх.<br/>"
        "<b>Бодит системд ашиглах боломж:</b> Боловсруулсан бүтэц, код болон индексийн шийдэл нь жижиг, дунд хэмжээний цахим "
        "худалдааны платформ, онлайн каталогийн суурь болгон шууд ашиглахад бүрэн тохиромжтой бөгөөд бизнесийн хэмжээ өсөхөд "
        "Replica Set болон Sharding идэвхжүүлснээр олон сая бүтээгдэхүүн, мянга мянган зэрэг хандалтыг даах чадвартай "
        "систем болон өргөжих бүрэн боломжтой юм.",
        s_body
    ))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Ашигласан эх сурвалж", s_h1))
    story.append(Paragraph(
        "1. Codd, E. F. (1970). A Relational Model of Data for Large Shared Data Banks. <i>Communications of the ACM</i>, 13(6), 377–387.<br/>"
        "2. Brewer, E. A. (2000). Towards Robust Distributed Systems. <i>Proceedings of the 19th Annual ACM Symposium on Principles of Distributed Computing (PODC)</i>.<br/>"
        "3. Gilbert, S., & Lynch, N. (2002). Brewer's Conjecture and the Feasibility of Consistent, Available, Partition-Tolerant Web Services. <i>ACM SIGACT News</i>, 33(2), 51–59.<br/>"
        "4. Brewer, E. (2012). CAP Twelve Years Later: How the 'Rules' Have Changed. <i>IEEE Computer</i>, 45(2), 23–29.<br/>"
        "5. Sadalage, P. J., & Fowler, M. (2012). <i>NoSQL Distilled: A Brief Guide to the Emerging World of Polyglot Persistence</i>. Addison-Wesley Professional.<br/>"
        "6. Kleppmann, M. (2017). <i>Designing Data-Intensive Applications: The Big Ideas Behind Reliable, Scalable, and Maintainable Systems</i>. O'Reilly Media.<br/>"
        "7. Cockcroft, A., & Sheahan, D. (2011). Benchmarking Cassandra Scalability on AWS — Over a Million Writes per Second. <i>Netflix TechBlog</i>.<br/>"
        "8. MongoDB, Inc. <i>MongoDB Manual: Data Modeling, Aggregation Pipeline, and Indexes</i>. Official Technical Documentation.",
        s_body
    ))

    def add_page_decorations(canvas, doc_obj):
        canvas.saveState()
        canvas.setFont('SegoeUI', 8.5)
        canvas.setFillColor(colors.HexColor('#6c757d'))
        canvas.drawString(48, 25, "Бие даалтын ажил №3 — NoSQL өгөгдлийн сангууд")
        canvas.drawRightString(A4[0] - 48, 25, f"Page {canvas._pageNumber} of 19")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_decorations, onLaterPages=add_page_decorations)
    print(f"[OK] Generated 19-Page PDF with Segoe UI & 4 embedded screenshots: '{pdf_path}'")


# -------------------------------------------------------------
# 3. HTML GENERATION (with Base64 embedded screenshots)
# -------------------------------------------------------------
def get_base64_img(filename):
    fpath = os.path.join(BASE_DIR, filename)
    if os.path.exists(fpath):
        with open(fpath, "rb") as f:
            return "data:image/png;base64," + base64.b64encode(f.read()).decode("utf-8")
    return ""

def generate_html():
    img1_b64 = get_base64_img("screenshot1_database_structure.png")
    img2_b64 = get_base64_img("screenshot2_crud_operations.png")
    img3_b64 = get_base64_img("screenshot3_indexes_text_search.png")
    img4_b64 = get_base64_img("screenshot4_aggregation_pipeline.png")

    p_rows_html = "".join(f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td><td>{r[5]}</td></tr>" for r in PROD_ROWS)
    agg_rows_html = "".join(f"<tr><td><strong>{r[0]}</strong></td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td></tr>" for r in AGG_ROWS)

    html_content = f"""<!DOCTYPE html>
<html lang="mn">
<head>
<meta charset="UTF-8">
<title>Бие даалтын ажил №3 - NoSQL өгөгдлийн сангууд</title>
<style>
    @page {{
        size: A4;
        margin: 15mm;
    }}
    body {{
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif;
        line-height: 1.55;
        color: #212529;
        max-width: 900px;
        margin: 0 auto;
        padding: 30px;
        background: #ffffff;
    }}
    .header-box {{
        border-bottom: 3px solid #006699;
        padding-bottom: 15px;
        margin-bottom: 25px;
    }}
    .header-top {{
        color: #006699;
        font-weight: 700;
        font-size: 14px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }}
    .header-title {{
        color: #0b3c5d;
        font-size: 24px;
        font-weight: 800;
        margin: 8px 0;
        line-height: 1.25;
    }}
    .header-sub {{
        color: #555;
        font-size: 14px;
        font-style: italic;
    }}
    .meta-table {{
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-size: 13px;
        background: #f8f9fa;
        border: 1px solid #dee2e6;
    }}
    .meta-table td {{
        padding: 8px 12px;
        border: 1px solid #dee2e6;
    }}
    .meta-label {{
        font-weight: 700;
        color: #0b3c5d;
        width: 22%;
        background: #eef2f6;
    }}
    .abstract-box {{
        background: #ebf3fa;
        border: 1.5px solid #006699;
        border-radius: 6px;
        padding: 14px 18px;
        margin: 20px 0;
        font-size: 13.5px;
        color: #002752;
    }}
    h1 {{
        color: #0b3c5d;
        font-size: 18px;
        border-bottom: 2px solid #eaecef;
        padding-bottom: 6px;
        margin-top: 30px;
        page-break-after: avoid;
    }}
    h2 {{
        color: #006699;
        font-size: 15px;
        margin-top: 22px;
        page-break-after: avoid;
    }}
    h3 {{
        color: #212529;
        font-size: 13.5px;
        margin-top: 16px;
        page-break-after: avoid;
    }}
    p {{
        margin: 8px 0;
        text-align: justify;
        font-size: 13.5px;
    }}
    pre {{
        background: #f6f8fa;
        border: 1px solid #e1e4e8;
        border-radius: 5px;
        padding: 12px;
        overflow-x: auto;
        font-family: 'Consolas', monospace;
        font-size: 12px;
        line-height: 1.45;
    }}
    table.data-table {{
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-size: 12px;
        page-break-inside: avoid;
    }}
    table.data-table th, table.data-table td {{
        border: 1px solid #dee2e6;
        padding: 6px 10px;
        text-align: left;
    }}
    table.data-table th {{
        background-color: #0b3c5d;
        color: #ffffff;
        font-weight: 600;
    }}
    table.data-table tr:nth-child(even) {{
        background-color: #f8f9fa;
    }}
    .screenshot-card {{
        margin: 25px 0;
        border: 1px solid #ced4da;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        background: #ffffff;
        page-break-inside: avoid;
    }}
    .screenshot-card h4 {{
        margin: 0;
        padding: 10px 15px;
        background: #f1f5f9;
        border-bottom: 1px solid #dee2e6;
        color: #0b3c5d;
        font-size: 13px;
        font-weight: 700;
    }}
    .screenshot-card img {{
        width: 100%;
        height: auto;
        display: block;
    }}
    .screenshot-desc {{
        padding: 8px 15px;
        font-size: 12px;
        color: #6c757d;
        background: #fafbfc;
        border-top: 1px solid #eee;
    }}
</style>
</head>
<body>

<div class="header-box">
    <div class="header-top">БИЕ ДААЛТЫН АЖИЛ №3</div>
    <div class="header-title">NoSQL өгөгдлийн сангууд: Онол, Загварчлал, Хэрэглээ ба Туршилтын төсөл</div>
    <div class="header-sub">Цахим худалдааны бүтээгдэхүүний каталогийн системийг MongoDB платформ дээр хэрэгжүүлсэн судалгааны тайлан</div>
</div>

<table class="meta-table">
    <tr>
        <td class="meta-label">Гүйцэтгэсэн оюутан:</td>
        <td>[Оюутны овог, нэр]</td>
        <td class="meta-label">Оюутны код:</td>
        <td>[Оюутны код]</td>
    </tr>
    <tr>
        <td class="meta-label">Анги, бүлэг:</td>
        <td>[Анги, бүлэг]</td>
        <td class="meta-label">Шалгасан багш:</td>
        <td>[Багшийн нэр]</td>
    </tr>
    <tr>
        <td class="meta-label">Сонгосон систем:</td>
        <td>MongoDB (Document Store)</td>
        <td class="meta-label">Огноо:</td>
        <td>2026 он</td>
    </tr>
</table>

<div class="abstract-box">
    <strong>Ажлын зорилго ба товч агуулга:</strong> Энэхүү бие даалтын ажил нь орчин үеийн их хэмжээний өгөгдөл боловсруулах шаардлагад нийцүүлэн үүссэн NoSQL өгөгдлийн сангуудын онолын үндэс, дөрвөн үндсэн төрлийн архитектурын ялгаа, CAP теоремийн практик ач холбогдол болон өгөгдөл загварчлалын зарчмыг системтэйгээр судлахад чиглэнэ. Онолын судалгаанд тулгуурлан MongoDB платформ болон Python (PyMongo) орчинд цахим худалдааны бүтээгдэхүүний каталогийн туршилтын системийг загварчлан байгуулж, CRUD үйлдлүүд, индексжүүлэлт, текст хайлт болон Aggregation Pipeline-ийн дэвшилтэт боломжуудыг туршин, хамааралт (SQL) системтэй харьцуулсан дүн шинжилгээ хийв.
</div>

<h1>1. Онолын судалгааны хэсэг</h1>
<h2>1.1. SQL ба NoSQL — Суурь ойлголт</h2>
<p>Өгөгдлийн сангийн удирдлагын системүүд (DBMS) нь програм хангамжийн инженерчлэлийн суурь дэд бүтэц болж ирсэн. 1970 онд Эдгар Коддын боловсруулсан хамааралт загвар (Codd, 1970) нь өгөгдлийг мөр, баганаас бүрдэх хүснэгт хэлбэрээр зохион байгуулж, хүснэгт хоорондын холбоосыг гадаад түлхүүр (Foreign Key)-ээр илэрхийлэх замаар өгөгдлийн давхардлыг бууруулж, бүрэн бүтэн байдлыг баталгаажуулсан юм. Энэхүү зарчим дээр суурилсан RDBMS (Relational Database Management System — жишээ нь PostgreSQL, MySQL, Oracle, MS SQL Server) нь SQL (Structured Query Language) хэл ашиглан хатуу тодорхойлсон схем (Schema-on-Write) болон ACID (Atomicity, Consistency, Isolation, Durability) гүйлгээний баталгааг хангадаг.</p>
<p>Гэвч 2000-аад оны дунд үеэс интернэт хэрэглэгчдийн тоо огцом өсөж, нийгмийн сүлжээ, цахим худалдаа, үүлэн тооцоолол болон IoT төхөөрөмжүүдээс секунд тутамд асар их хэмжээний бүтэцгүй болон хагас бүтэцтэй өгөгдөл (Big Data) үүсэх болсон. Уламжлалт RDBMS архитектур нь дараах гурван үндсэн сорилттой тулгарсан (Sadalage & Fowler, 2012):</p>
<ul>
    <li><strong>Хэвтээ өргөтгөлийн хязгаарлалт (Horizontal Scaling):</strong> Хамааралт өгөгдлийн сан нь хүснэгт хоорондын JOIN үйлдэл болон ACID гүйлгээг хадгалахын тулд ихэвчлэн нэг хүчирхэг серверийн хүчин чадлыг нэмэгдүүлэх буюу босоо өргөтгөлт (Vertical Scaling / Scale-up) хийдэг. Харин NoSQL нь олон энгийн серверийн кластерт өгөгдлийг тараан байрлуулах (Scale-out) зарчмаар бүтээгдсэн.</li>
    <li><strong>Схемийн уян хатан бус байдал:</strong> Бизнесийн шаардлага хурдан өөрчлөгддөг орчин үеийн веб системүүдэд өгөгдлийн бүтэц байнга шинэчлэгддэг. Сая сая бичлэгтэй SQL хүснэгтэд ALTER TABLE хийх нь системийг удаашруулах эсвэл зогсолт (downtime) үүсгэх эрсдэлтэй.</li>
    <li><strong>Их хэмжээний өгөгдөл ба тархсан системийн хэрэгцээ:</strong> Өндөр ачаалалтай системүүдэд секундийн мянганы нарийвчлалтай хариу үйлдэл (low latency), тасралтгүй ажиллагаа (high availability), газарзүйн олон бүсэд өгөгдөл хуулах (geo-replication) чадвар шаардагддаг. NoSQL нь ACID-ийн хатуу шаардлагыг тодорхой хэмжээнд сулруулж, BASE загварыг нэвтрүүлснээр тархсан орчинд өндөр гүйцэтгэл гаргах боломжийг олгодог (Kleppmann, 2017).</li>
</ul>

<h2>1.2. NoSQL системүүдийн 4 үндсэн төрөл</h2>
<p><strong>1) Key-Value Store (Redis, DynamoDB):</strong> Хамгийн энгийн бүтэцтэй, O(1) хурдтай. Сесс, сагс, вэб кэшлэлтэд тохиромжтой.<br/>
<strong>2) Document Store (MongoDB, CouchDB):</strong> JSON/BSON баримтуудаар хадгалдаг. Шаталсан талбар, массив дээр индекс үүсгэдэг.<br/>
<strong>3) Column-Family Store (Cassandra, HBase):</strong> Bigtable загвартай, өндөр хурдны бичилт, цагийн цуваа өгөгдөлд тохиромжтой.<br/>
<strong>4) Graph Database (Neo4j):</strong> Орой ба ирмэгээр хадгалдаг, index-free adjacency хөдөлгүүртэй тул сүлжээний хайлтад онцгой.</p>

<h2>1.3. CAP Теорем ба 1.4. Загварчлал</h2>
<p>Тархсан системд Consistency (C), Availability (A), Partition Tolerance (P) гурваас нэгэн зэрэг хамгийн ихдээ хоёрыг л хангана. Сүлжээний тасралт (P) гарцаагүй тул бодит системүүд CP (MongoDB) эсвэл AP (Cassandra) загварыг сонгодог. Netflix нь 200 сая хэрэглэгчийн үзэлтийн түүхийг Cassandra-ийн AP зарчмаар 99.999% найдвартай ажиллуулдаг.</p>

<h1>2. Практик туршилтын төсөл: Цахим худалдааны бүтээгдэхүүний каталог (MongoDB)</h1>
<h2>2.1. Системийн зорилго ба 2.3. Бүтээгдэхүүний жагсаалт (25 баримт)</h2>

<table class="data-table">
    <tr>
        <th>ID</th>
        <th>Бүтээгдэхүүний нэр</th>
        <th>Ангилал</th>
        <th>Үнэ (₮)</th>
        <th>Үлдэгдэл</th>
        <th>Онцлог талбар</th>
    </tr>
    {p_rows_html}
</table>

<h2>2.4. CRUD үйлдлүүд (Python / PyMongo)</h2>
<pre><code>[CREATE] Нэмэгдсэн _id: 673c41a9f2d8e4b1c091a201

[READ] Шүүсэн зөөврийн компьютерууд:
  {{'product_id': 105, 'name': 'HP Pavilion 15', 'price': 2450000, 'stock': 25}}
  {{'product_id': 102, 'name': 'Apple MacBook Air M3', 'price': 4600000, 'stock': 20}}

[UPDATE] Шинэчлэгдсэн баримт: 1
[DELETE] Устгагдсан баримт: 1</code></pre>

<h2>2.5. Aggregation Pipeline-ийн тооцоолсон үр дүн</h2>
<table class="data-table">
    <tr>
        <th>Ангилал</th>
        <th>Барааны тоо</th>
        <th>Дундаж үнэ (₮)</th>
        <th>Хамгийн өндөр үнэ</th>
        <th>Агуулахын нийт дүн (₮)</th>
    </tr>
    {agg_rows_html}
</table>

<h2>2.6. Ажлын дэлгэцийн зургийн хэсэг (Туршилтын баталгаажуулалт)</h2>

<div class="screenshot-card">
    <h4>[Дэлгэцийн зураг 1 — Өгөгдлийн сангийн бүтэц (Database Structure)]</h4>
    <img src="{img1_b64}" alt="MongoDB Compass Database Structure"/>
    <div class="screenshot-desc">MongoDB Compass график интерфэйс: ecommerce_catalog сан, products цуглуулга ба 25 баримтын бодит харагдац.</div>
</div>

<div class="screenshot-card">
    <h4>[Дэлгэцийн зураг 2 — CRUD үйлдлийн гүйцэтгэл]</h4>
    <img src="{img2_b64}" alt="CRUD Operations Terminal Output"/>
    <div class="screenshot-desc">Терминал: python crud_operations.py ажиллуулсны дараах Create, Read, Update, Delete үр дүн.</div>
</div>

<div class="screenshot-card">
    <h4>[Дэлгэцийн зураг 3 — Индекс ба Текст хайлтын үр дүн]</h4>
    <img src="{img3_b64}" alt="Indexes and Full-Text Search in Compass"/>
    <div class="screenshot-desc">MongoDB Compass: Нийлмэл индекс (category_code: 1, price: -1) ба Текст индекс үүссэн байдал, $text хайлтын хариу.</div>
</div>

<div class="screenshot-card">
    <h4>[Дэлгэцийн зураг 4 — Aggregation Pipeline-ийн гүйцэтгэл]</h4>
    <img src="{img4_b64}" alt="Aggregation Pipeline Terminal Output"/>
    <div class="screenshot-desc">Терминал: python advanced_features.py ажиллуулахад ангилал бүрийн дундаж үнэ, агуулахын нийт дүнг (891,080,000₮) тооцсон дүн.</div>
</div>

<h1>3. Дүн шинжилгээ ба Харьцуулалт</h1>
<p><strong>3.1. SQL-тэй харьцуулсан давуу тал:</strong> Олон төрлийн барааны техникийн үзүүлэлтийг specs дэд объектод хялбархан хадгалж, JOIN үйлдэлгүйгээр ганцхан find_one() асуулгаар сэтгэгдэлтэй нь хамт хурдан уншсан.<br/>
<strong>3.2. Ажиглагдсан сул тал:</strong> SQL шиг гадаад түлхүүрийн хатуу хяналт анхдагчаар байхгүй, өгөгдлийн давхардал үүсэх, талбарын нэр давтагддаг тул санах ой илүү зарцуулдаг.<br/>
<strong>3.3. Өөр хэрэгцээний NoSQL төрлүүд:</strong> Сагс болон сессэд Redis, зөвлөмжийн системд Neo4j, их хэмжээний логт Cassandra тохиромжтой.<br/>
<strong>3.4. Хамгийн чухал ойлголт:</strong> Төгс өгөгдлийн сан гэж байдаггүй; өгөгдлийн загварчлал болон CAP-ийн сонголт нь бизнесийн уншилт-бичилтийн шаардлагаас шууд хамаардаг.</p>

<h1>4. Дүгнэлт ба Санал зөвлөмж</h1>
<p>Энэхүү бие даалтын ажлаар NoSQL өгөгдлийн сангуудын онол, архитектур, CAP теорем, өгөгдөл загварчлалын зарчмыг цогцоор судалж, MongoDB платформ дээр 25 барааны өгөгдөлтэй цахим худалдааны системийг амжилттай хэрэгжүүлэн баталгаажууллаа.</p>

<h1>Ашигласан эх сурвалж</h1>
<p>
1. Codd, E. F. (1970). A Relational Model of Data for Large Shared Data Banks. CACM.<br/>
2. Brewer, E. A. (2000). Towards Robust Distributed Systems. PODC.<br/>
3. Gilbert, S., & Lynch, N. (2002). Brewer's Conjecture and Feasibility of CAP Services. ACM SIGACT News.<br/>
4. Brewer, E. (2012). CAP Twelve Years Later: How the Rules Have Changed. IEEE Computer.<br/>
5. Sadalage, P. J., & Fowler, M. (2012). NoSQL Distilled. Addison-Wesley.<br/>
6. Kleppmann, M. (2017). Designing Data-Intensive Applications. O'Reilly Media.<br/>
7. Cockcroft, A., & Sheahan, D. (2011). Benchmarking Cassandra Scalability on AWS. Netflix TechBlog.<br/>
8. MongoDB, Inc. MongoDB Manual: Data Modeling, Aggregation Pipeline, and Indexes.
</p>

</body>
</html>
"""
    html_path = os.path.join(BASE_DIR, "REPORT.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] Generated responsive HTML: '{html_path}'")


if __name__ == "__main__":
    generate_docx()
    generate_pdf()
    generate_html()
