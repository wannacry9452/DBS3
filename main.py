"""
main.py
Master executable runner for NoSQL Project:
"NoSQL өгөгдлийн сангууд — Онол, Загварчлал, Хэрэглээ, Туршилтын төсөл"

Executes:
1. Environment and connection check
2. Database initialization and seeding (35 products, 5 customer orders)
3. Full CRUD operations demonstration
4. Advanced NoSQL features: Aggregation pipelines, indexing, text search
5. Console output logging into 'test_run.log'
"""

import sys
import io
import time
from db_setup import get_database_client, seed_database, DB_NAME, PRODUCTS_COLLECTION, ORDERS_COLLECTION
import crud_operations as crud
import advanced_features as adv

class DualLogger:
    """Logs terminal output simultaneously to stdout and test_run.log."""
    def __init__(self, filename="test_run.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()

    def close(self):
        self.log.close()

def run_project_pipeline():
    print("=" * 80)
    print("      NoSQL ӨГӨГДЛИЙН САНГИЙН ТУРШИЛТЫН ТӨСӨЛ — СИСТЕМИЙН АЖИЛЛАГАА")
    print("      Сэдэв: E-Commerce цахим худалдааны өгөгдлийн загварчлал ба MongoDB")
    print("=" * 80)

    # 1. Database Connection
    print("\n[ҮЕ ШАТ 1] Өгөгдлийн сантай холбогдох болон орчин бэлтгэх...")
    client, is_embedded = get_database_client()
    db = client[DB_NAME]

    # 2. Seeding
    print("\n[ҮЕ ШАТ 2] Анхны өгөгдөл бэлтгэх, ачаалах (Seeding)...")
    prod_count, order_count = seed_database(db)
    print(f" -> Баталгаажуулалт: Нийт {prod_count} бүтээгдэхүүн, {order_count} захиалга амжилттай үүслээ.")

    products = db[PRODUCTS_COLLECTION]
    orders = db[ORDERS_COLLECTION]

    # 3. CRUD Operations
    print("\n" + "=" * 80)
    print("[ҮЕ ШАТ 3] CRUD ҮЙЛДЛҮҮДИЙН ТУРШИЛТ")
    print("=" * 80)

    # 3.1 Create (Шинэ бүтээгдэхүүн нэмэх)
    print("\n>>> 3.1 CREATE: Шинэ баримт нэмэх (insert_one)")
    new_item = {
        "product_id": "PROD-036",
        "name": "Apple Vision Pro 512GB Spatial Computer",
        "slug": "apple-vision-pro-512",
        "category": "Smartphones & Tablets",
        "sub_category": "Spatial Computing",
        "brand": "Apple",
        "price": 3699.00,
        "discount_price": 3549.00,
        "stock": 10,
        "status": "in_stock",
        "specs": {
            "display": "Dual Micro-OLED 23M pixels 3D",
            "processor": "M2 + R1 dual chip",
            "sensors": "12 cameras, 5 sensors, 6 mics",
            "spatial_audio": "Spatial Audio with dynamic head tracking"
        },
        "tags": ["apple", "vision-pro", "vr", "ar", "spatial-computing"],
        "rating_summary": {"average_rating": 5.0, "review_count": 1},
        "reviews": [
            {
                "review_id": "REV-NEW-01",
                "user_name": "Эрдэм",
                "rating": 5.0,
                "comment": "Орон зайн тооцоолол үнэхээр дараагийн үеийн технологи байна.",
                "date": "2024-02-20T10:00:00Z"
            }
        ],
        "created_at": "2024-02-20T09:00:00Z"
    }
    crud.create_product(products, new_item)

    # 3.2 Read (Хайлт)
    print("\n>>> 3.2 READ: Төрөл бүрийн шүүлтүүрээр хайх")
    print("\n A) ID-аар хайх:")
    crud.get_product_by_id(products, "PROD-036")

    print("\n B) Ангилал ба Үнийн хязгаараар хайх ($gte, $lte):")
    crud.search_products_by_filter(products, category="Computers & Laptops", min_price=1000, max_price=2600)

    print("\n C) Бүтцийн доторх Nested талбараар хайх (specs.ram_gb >= 32):")
    crud.query_nested_specs(products, min_ram_gb=32)

    print("\n D) Массивын түлхүүр үгээр хайх (tags = 'developer'):")
    crud.query_by_tag(products, "developer")

    # 3.3 Update (Шинэчлэх)
    print("\n>>> 3.3 UPDATE: Талбар шинэчлэх ($set), тоолуур өөрчлөх ($inc), массив нэмэх ($push)")
    print("\n A) Үнэ болон үлдэгдэл шинэчлэх:")
    crud.update_product_price_and_stock(products, "PROD-036", 3499.00, -1)

    print("\n B) Хэрэглэгчийн сэтгэгдэл нэмэх ($push) ба үнэлгээг дахин тооцоолох:")
    crud.add_product_review(products, "PROD-036", "Наранбаатар", 4.8, "Виртуал ажлын орчин маш бүтээмжтэй.")

    # 3.4 Delete (Устгах)
    print("\n>>> 3.4 DELETE: Баримт устгах")
    # Add a temporary dummy item to test delete
    dummy_id = crud.create_product(products, {
        "product_id": "PROD-TEMP-999",
        "name": "Temporary Test Item",
        "price": 10.0,
        "stock": 0,
        "category": "Test"
    })
    print(f" Устгахын өмнөх нийт тоо: {products.count_documents({})}")
    crud.delete_product_by_id(products, "PROD-TEMP-999")
    print(f" Устгасны дараах нийт тоо: {products.count_documents({})}")

    # 4. Advanced Features
    print("\n" + "=" * 80)
    print("[ҮЕ ШАТ 4] ДЭВШИЛТЭТ БОЛОМЖУУДЫН ТУРШИЛТ (AGGREGATION, INDEX, TEXT SEARCH)")
    print("=" * 80)

    adv.aggregate_category_summary(products)
    adv.aggregate_review_sentiment_unwind(products)
    adv.aggregate_orders_by_customer(orders)
    adv.perform_text_search(products, "apple m3")
    adv.analyze_index_performance(products)

    print("\n" + "=" * 80)
    print(" ТУРШИЛТ БҮРЭН АМЖИЛТТАЙ ДУУСЛАА! БҮХ ҮР ДҮН test_run.log ФАЙЛД ХАДГАЛАГДЛАА.")
    print("=" * 80)

if __name__ == "__main__":
    logger = DualLogger("test_run.log")
    sys.stdout = logger
    try:
        run_project_pipeline()
    finally:
        sys.stdout = logger.terminal
        logger.close()
