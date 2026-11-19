# advanced_features.py — Индекс, Текст хайлт ба Aggregation Pipeline
import os

try:
    from pymongo import MongoClient, ASCENDING, DESCENDING, TEXT
    client = MongoClient("mongodb://localhost:27017/", serverSelectionTimeoutMS=1500)
    client.admin.command('ping')
except Exception:
    import mongo_engine
    client = mongo_engine.MongoClient("mongodb://localhost:27017/")
    ASCENDING = mongo_engine.ASCENDING
    DESCENDING = mongo_engine.DESCENDING
    TEXT = mongo_engine.TEXT

db = client["ecommerce_catalog"]
products = db["products"]

# Хэрэв цуглуулга хоосон бол эхлүүлэх
if products.count_documents({}) == 0:
    import json
    dataset_path = os.path.join(os.path.dirname(__file__), "dataset.json")
    with open(dataset_path, "r", encoding="utf-8") as f:
        products.insert_many(json.load(f))

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

print("\n=== AGGREGATION PIPELINE: АНГИЛЛЫН ТАЙЛАН ===")
for row in products.aggregate(pipeline):
    print(
        f" {row['_id']:10s} | Тоо: {row['product_count']:2d} | "
        f"Дундаж үнэ: {row['avg_price']:10,.0f}₮ | "
        f"Агуулахын нийт дүн: {row['total_inventory_value']:13,.0f}₮"
    )
