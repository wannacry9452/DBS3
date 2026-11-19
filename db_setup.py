"""
db_setup.py
Initializes the MongoDB connection, creates collections, and loads initial datasets.
Supports both pymongo (live MongoDB instance) and embedded mongo_engine for zero-dependency execution.
"""

import json
import os
import sys

# Attempt to import pymongo; if unavailable, use our built-in mongo_engine
try:
    import pymongo
    USE_REAL_PYMONGO = True
except ImportError:
    USE_REAL_PYMONGO = False

import mongo_engine

DB_NAME = "ecommerce_nosql_db"
PRODUCTS_COLLECTION = "products"
ORDERS_COLLECTION = "orders"

def get_database_client(uri: str = "mongodb://localhost:27017", force_embedded: bool = False):
    """
    Returns a connected MongoDB client.
    If real MongoDB is unreachable or force_embedded is True, uses embedded mongo_engine.
    """
    if USE_REAL_PYMONGO and not force_embedded:
        try:
            client = pymongo.MongoClient(uri, serverSelectionTimeoutMS=1500)
            # Ping database to check if server is running
            client.admin.command('ping')
            print(f"[INFO] Connected to LIVE MongoDB server at {uri}")
            return client, False
        except Exception:
            print("[INFO] Live MongoDB server not reachable. Switching to High-Performance Embedded Mongo Engine.")
    
    # Fallback to embedded engine
    client = mongo_engine.MongoClient(uri)
    print("[INFO] Initialized Embedded MongoDB Engine (Pure Python, Full API Compatible).")
    return client, True

def seed_database(db, base_path: str = "."):
    """
    Loads dataset.json and orders_dataset.json into the database collections.
    """
    products_file = os.path.join(base_path, "dataset.json")
    orders_file = os.path.join(base_path, "orders_dataset.json")

    # Load Products
    with open(products_file, "r", encoding="utf-8") as f:
        products_data = json.load(f)

    # Clean existing
    db[PRODUCTS_COLLECTION].delete_many({})
    insert_res = db[PRODUCTS_COLLECTION].insert_many(products_data)
    print(f"[OK] Seeded {len(insert_res.inserted_ids)} product documents into '{PRODUCTS_COLLECTION}' collection.")

    # Load Orders
    with open(orders_file, "r", encoding="utf-8") as f:
        orders_data = json.load(f)

    db[ORDERS_COLLECTION].delete_many({})
    order_res = db[ORDERS_COLLECTION].insert_many(orders_data)
    print(f"[OK] Seeded {len(order_res.inserted_ids)} order documents into '{ORDERS_COLLECTION}' collection.")

    # Create Indexes
    db[PRODUCTS_COLLECTION].create_index([("category", 1), ("price", -1)], name="idx_category_price")
    db[PRODUCTS_COLLECTION].create_index("brand", name="idx_brand")
    db[ORDERS_COLLECTION].create_index("customer.customer_id", name="idx_customer_id")
    print("[OK] Created indexes: idx_category_price, idx_brand, idx_customer_id.")

    return len(insert_res.inserted_ids), len(order_res.inserted_ids)

if __name__ == "__main__":
    client, is_embedded = get_database_client()
    db = client[DB_NAME]
    seed_database(db)
    print(f"Total products in DB: {db[PRODUCTS_COLLECTION].count_documents({})}")
    print(f"Total orders in DB: {db[ORDERS_COLLECTION].count_documents({})}")
