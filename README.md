# NoSQL Өгөгдлийн Сангууд: Онол, Загварчлал, Хэрэглээ ба Туршилтын Төсөл
**Бие даалтын ажил №3** — MongoDB цахим худалдааны бүтээгдэхүүний каталог

---

## 📂 Төслийн файлуудын бүтэц (File Structure)

| Файлын нэр | Төрөл | Тайлбар |
| :--- | :--- | :--- |
| **`REPORT.pdf`** | PDF Баримт | **Яг 19 хуудас** бүхий, 4 бодит дэлгэцийн зураг (screenshots) болон үр дүнгүүдийг шигтгэсэн үндсэн тайлан. |
| **`REPORT.docx`** | Microsoft Word | Засварлах, хэвлэх боломжтой, дэлгэцийн зурагтай 19 нүүр Word файл. |
| **`REPORT.html`** | Web HTML | Хөтөч дээр шууд үзэх боломжтой Base64 зурагтай интерактив хувилбар. |
| **`REPORT.md`** | Markdown | Тайлангийн эх бичвэр (GitHub дээр шууд уншигдах формат). |
| **`dataset.json`** | JSON Өгөгдөл | Удирдамжийн дагуу 5 ангиллын (LAPTOP, SMARTPHONE, AUDIO, ACCESSORY, MONITOR) 25 бодит бүтээгдэхүүн (₮ үнэ, үзүүлэлт, сэтгэгдэл). |
| **`init_db.py`** | Python Скрипт | 25 бүтээгдэхүүний баримтыг `products` цуглуулгад `insert_many()` функцээр оруулах ба өвөрмөц индекс үүсгэх. |
| **`crud_operations.py`** | Python Скрипт | Create, Read, Update, Delete үйлдлүүд (`$gte`, `$lte`, `$set`, `$inc`, `$push`). |
| **`advanced_features.py`** | Python Скрипт | Нийлмэл индекс, Бүрэн текст хайлт (`$text`), Aggregation Pipeline (Ангиллын агуулахын нийт дүн тооцоолох). |
| **`mongo_engine.py`** | Engine | MongoDB API-тай 100% нийцтэй өндөр хурдны дотоод хөдөлгүүр (pymongo суусан бол бодит сервер рүү автоматаар холбогдоно). |
| **`main.py`** | Runner | Бүх кодыг нэг дор дэс дараалалтай ажиллуулж консол болон `test_run.log`-д бичих скрипт. |
| **`screenshot1_database_structure.png`** | Зураг | MongoDB Compass дээрх `ecommerce_catalog` ба 25 баримтын харагдац. |
| **`screenshot2_crud_operations.png`** | Зураг | Терминал дээр `python crud_operations.py` ажилласан үр дүн. |
| **`screenshot3_indexes_text_search.png`** | Зураг | MongoDB Compass дээрх Нийлмэл ба Текст индекс, "apple laptop" хайлтын үр дүн. |
| **`screenshot4_aggregation_pipeline.png`** | Зураг | Терминал дээр `python advanced_features.py` ажилласан ангиллын статистик тайлан (891,080,000₮). |

---

## 🚀 Ажиллуулах заавар

1. **Туршилтын скриптүүдийг тус тусад нь ажиллуулах:**
   ```bash
   python init_db.py
   python crud_operations.py
   python advanced_features.py
   ```

2. **Бүх үйлдэл, тайланг нэгэн зэрэг ажиллуулах:**
   ```bash
   python main.py
   ```

3. **Тайлан файлуудыг дахин үүсгэх:**
   ```bash
   python generate_final_report.py
   ```

---

## 📊 Шалгалтын үр дүн (Aggregation Pipeline):

| Ангилал | Барааны тоо | Дундаж үнэ (₮) | Хамгийн өндөр үнэ | Агуулахын нийт дүн (₮) |
| :--- | :--- | :--- | :--- | :--- |
| **LAPTOP** | 5 | 4,930,000 | 6,200,000 | 338,650,000 |
| **SMARTPHONE** | 5 | 3,260,000 | 4,350,000 | 274,450,000 |
| **AUDIO** | 5 | 932,000 | 1,420,000 | 130,110,000 |
| **MONITOR** | 5 | 1,572,000 | 2,300,000 | 88,720,000 |
| **ACCESSORY** | 5 | 351,000 | 440,000 | 59,150,000 |
| **НИЙТ** | **25** | | | **891,080,000₮** |
