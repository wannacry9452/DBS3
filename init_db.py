# init_db.py — Өгөгдлийн сан байгуулах ба 25 баримт оруулах
import json
import os
from datetime import datetime, timezone

try:
    from pymongo import MongoClient
    client = MongoClient("mongodb://localhost:27017/", serverSelectionTimeoutMS=1500)
    client.admin.command('ping')
except Exception:
    import mongo_engine
    client = mongo_engine.MongoClient("mongodb://localhost:27017/")

db = client["ecommerce_catalog"]
products = db["products"]

# Хуучин туршилтын өгөгдлийг цэвэрлэх
products.drop()

# product_id талбар дээр давхардалгүй индекс үүсгэх
products.create_index("product_id", unique=True)

# 25 бүтээгдэхүүний өгөгдлийг dataset.json-оос ачаалах
dataset_path = os.path.join(os.path.dirname(__file__), "dataset.json")
with open(dataset_path, "r", encoding="utf-8") as f:
    sample_products = json.load(f)

result = products.insert_many(sample_products)
print(f"Амжилттай оруулсан баримтын тоо: {len(result.inserted_ids)}")
