# crud_operations.py — Бүрэн CRUD үйлдлүүд
import os

try:
    from pymongo import MongoClient
    client = MongoClient("mongodb://localhost:27017/", serverSelectionTimeoutMS=1500)
    client.admin.command('ping')
except Exception:
    import mongo_engine
    client = mongo_engine.MongoClient("mongodb://localhost:27017/")

db = client["ecommerce_catalog"]
products = db["products"]

# Хэрэв цуглуулга хоосон бол эхлүүлэх
if products.count_documents({}) == 0:
    import json
    dataset_path = os.path.join(os.path.dirname(__file__), "dataset.json")
    with open(dataset_path, "r", encoding="utf-8") as f:
        products.insert_many(json.load(f))

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
print("\n[READ] Шүүсэн зөөврийн компьютерууд:")
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
print(f"\n[UPDATE] Шинэчлэгдсэн баримт: {upd.modified_count}")

# 4. DELETE: Туршилтын бүтээгдэхүүнийг устгах
deleted = products.delete_one({"product_id": 126})
print(f"\n[DELETE] Устгагдсан баримт: {deleted.deleted_count}")
