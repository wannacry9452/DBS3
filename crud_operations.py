"""
crud_operations.py
Implements core CRUD (Create, Read, Update, Delete) operations for the E-commerce NoSQL project.
Includes clean logging and formatted outputs to demonstrate NoSQL document manipulation.
"""

from typing import Dict, Any, List, Optional

def create_product(collection, product_doc: Dict[str, Any]) -> str:
    """
    CREATE: Inserts a single product document into the collection.
    Demonstrates schema-flexibility with nested specifications.
    """
    result = collection.insert_one(product_doc)
    print(f"[CREATE] Product added with ID: '{result.inserted_id}' | Name: {product_doc.get('name')}")
    return result.inserted_id

def create_multiple_products(collection, product_list: List[Dict[str, Any]]) -> List[str]:
    """
    CREATE: Batch inserts multiple products into the collection.
    """
    result = collection.insert_many(product_list)
    print(f"[CREATE BATCH] Successfully inserted {len(result.inserted_ids)} new products.")
    return result.inserted_ids

def get_product_by_id(collection, product_id: str) -> Optional[Dict[str, Any]]:
    """
    READ: Queries a single document by unique ID.
    """
    doc = collection.find_one({"product_id": product_id})
    if doc:
        print(f"[READ BY ID] Found: {doc['name']} (${doc['price']}) - Brand: {doc['brand']}")
    else:
        print(f"[READ BY ID] Product with ID '{product_id}' not found.")
    return doc

def search_products_by_filter(collection, category: Optional[str] = None, min_price: float = 0, max_price: float = float('inf')) -> List[Dict[str, Any]]:
    """
    READ: Queries documents by category and price range using MongoDB operators ($gte, $lte).
    """
    query: Dict[str, Any] = {
        "price": {"$gte": min_price, "$lte": max_price}
    }
    if category:
        query["category"] = category

    cursor = collection.find(query).sort("price", 1)
    results = cursor.to_list()
    print(f"[READ FILTER] Found {len(results)} items in Category: '{category or 'All'}' between ${min_price} and ${max_price}:")
    for item in results[:5]:  # print first 5
        print(f"   -> {item['name']} | Price: ${item['price']} | Stock: {item['stock']}")
    if len(results) > 5:
        print(f"   ... and {len(results) - 5} more items.")
    return results

def query_nested_specs(collection, min_ram_gb: int = 16) -> List[Dict[str, Any]]:
    """
    READ: Queries documents based on embedded nested fields (e.g., specs.ram_gb >= 16).
    Demonstrates MongoDB's native support for rich hierarchical data.
    """
    query = {"specs.ram_gb": {"$gte": min_ram_gb}}
    projection = {"name": 1, "brand": 1, "price": 1, "specs.ram_gb": 1, "specs.processor": 1}
    results = collection.find(query, projection).to_list()
    print(f"[READ NESTED] Products with specs.ram_gb >= {min_ram_gb}GB ({len(results)} items found):")
    for r in results:
        ram = r.get("specs", {}).get("ram_gb", "N/A")
        cpu = r.get("specs", {}).get("processor", "N/A")
        print(f"   * {r['name']} ({r['brand']}) - RAM: {ram}GB | CPU: {cpu}")
    return results

def query_by_tag(collection, tag: str) -> List[Dict[str, Any]]:
    """
    READ: Queries documents by array element match (e.g. tags contains 'gaming').
    """
    results = collection.find({"tags": tag}).to_list()
    print(f"[READ TAG] Products tagged with '{tag}' ({len(results)} items found):")
    for item in results:
        print(f"   # {item['name']} | Brand: {item['brand']} | Price: ${item['price']}")
    return results

def update_product_price_and_stock(collection, product_id: str, new_price: float, stock_change: int) -> bool:
    """
    UPDATE: Demonstrates $set (field update) and $inc (atomic counter increment/decrement).
    """
    update_query = {
        "$set": {"price": new_price, "status": "updated_sale"},
        "$inc": {"stock": stock_change}
    }
    res = collection.update_one({"product_id": product_id}, update_query)
    success = res.modified_count > 0
    print(f"[UPDATE] Product '{product_id}' modified: {res.modified_count} doc(s) updated. (Price set to ${new_price}, Stock delta: {stock_change})")
    return success

def add_product_review(collection, product_id: str, reviewer_name: str, rating: float, comment: str) -> bool:
    """
    UPDATE: Demonstrates $push to append a new embedded review object into the reviews array,
    and recalculates average rating.
    """
    new_review = {
        "review_id": f"REV-NEW-{reviewer_name}",
        "user_name": reviewer_name,
        "rating": rating,
        "comment": comment,
        "date": "2024-02-23T15:00:00Z"
    }

    # Push review
    res = collection.update_one({"product_id": product_id}, {"$push": {"reviews": new_review}})

    # Fetch to recalculate rating_summary
    product = collection.find_one({"product_id": product_id})
    if product and "reviews" in product and product["reviews"]:
        reviews = product["reviews"]
        avg = round(sum(r["rating"] for r in reviews) / len(reviews), 2)
        collection.update_one(
            {"product_id": product_id},
            {"$set": {"rating_summary.average_rating": avg, "rating_summary.review_count": len(reviews)}}
        )
        print(f"[UPDATE EMBEDDED] Added review by '{reviewer_name}' to '{product['name']}'. New Rating: {avg} ({len(reviews)} reviews).")
        return True
    return False

def delete_product_by_id(collection, product_id: str) -> bool:
    """
    DELETE: Deletes a single document by product_id.
    """
    res = collection.delete_one({"product_id": product_id})
    print(f"[DELETE ONE] Product '{product_id}' deleted: {res.deleted_count} document removed.")
    return res.deleted_count > 0

def delete_out_of_stock_products(collection) -> int:
    """
    DELETE: Batch deletes products that are out of stock (stock <= 0).
    """
    res = collection.delete_many({"stock": {"$lte": 0}})
    print(f"[DELETE MANY] Removed {res.deleted_count} out-of-stock items from catalog.")
    return res.deleted_count


if __name__ == "__main__":
    from db_setup import get_database_client, seed_database, DB_NAME, PRODUCTS_COLLECTION
    client, _ = get_database_client()
    db = client[DB_NAME]
    seed_database(db)

    products = db[PRODUCTS_COLLECTION]
    print("\n--- Testing CRUD Operations ---")
    get_product_by_id(products, "PROD-001")
    search_products_by_filter(products, "Audio & Sound", 100, 400)
    query_nested_specs(products, min_ram_gb=32)
    query_by_tag(products, "gaming")
    update_product_price_and_stock(products, "PROD-001", 3399.00, -2)
    add_product_review(products, "PROD-001", "Ганхөлөг", 5.0, "Гайхалтай хурдан, ямар ч өө сэвгүй машин байна!")
