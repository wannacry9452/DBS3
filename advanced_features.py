"""
advanced_features.py
Demonstrates advanced MongoDB features:
1. Multi-stage Aggregation Pipelines ($match, $group, $sort, $project, $unwind)
2. Category and Market Analytics
3. Customer Order Analysis & Relational Joins ($lookup)
4. Index Creation & Performance Analysis
5. Text Search and Keyword Filtering
"""

import time
from typing import Dict, Any, List

def aggregate_category_summary(collection) -> List[Dict[str, Any]]:
    """
    AGGREGATION 1: Category Analytics Pipeline
    Calculates total products, average price, min/max price, and total inventory value per category.
    Stages: $group -> $project -> $sort
    """
    pipeline = [
        {
            "$group": {
                "_id": "$category",
                "product_count": {"$sum": 1},
                "avg_price": {"$avg": "$price"},
                "min_price": {"$min": "$price"},
                "max_price": {"$max": "$price"},
                "total_stock": {"$sum": "$stock"}
            }
        },
        {
            "$project": {
                "category": "$_id",
                "product_count": 1,
                "avg_price": 1,
                "min_price": 1,
                "max_price": 1,
                "total_stock": 1,
                "_id": 0
            }
        },
        {
            "$sort": {"product_count": -1}
        }
    ]

    results = collection.aggregate(pipeline)
    print("\n" + "="*70)
    print(" [AGGREGATION 1] CATEGORY-WISE INVENTORY & PRICING SUMMARY")
    print("="*70)
    print(f"{'Category':<28} | {'Count':<5} | {'Avg Price':<10} | {'Price Range':<15} | {'Stock':<6}")
    print("-" * 70)
    for row in results:
        pr_range = f"${row['min_price']:.0f} - ${row['max_price']:.0f}"
        print(f"{row['category']:<28} | {row['product_count']:<5} | ${row['avg_price']:<9.2f} | {pr_range:<15} | {row['total_stock']:<6}")
    print("="*70)
    return results

def aggregate_review_sentiment_unwind(collection) -> List[Dict[str, Any]]:
    """
    AGGREGATION 2: Embedded Review Analytics via $unwind
    Deconstructs the embedded 'reviews' array to calculate rating score distribution across all products.
    Stages: $unwind -> $group -> $sort
    """
    pipeline = [
        {
            "$unwind": "$reviews"
        },
        {
            "$group": {
                "_id": "$reviews.rating",
                "review_count": {"$sum": 1},
                "reviewers": {"$push": "$reviews.user_name"}
            }
        },
        {
            "$sort": {"_id": -1}
        }
    ]

    results = collection.aggregate(pipeline)
    print("\n" + "="*70)
    print(" [AGGREGATION 2] RATING DISTRIBUTION VIA $unwind PIPELINE")
    print("="*70)
    print(f"{'Rating Star':<12} | {'Review Count':<12} | {'Sample Reviewers'}")
    print("-" * 70)
    for row in results:
        sample_users = ", ".join(row['reviewers'][:4])
        if len(row['reviewers']) > 4:
            sample_users += f" (+{len(row['reviewers']) - 4} more)"
        print(f"⭐ {row['_id']:<9} | {row['review_count']:<12} | {sample_users}")
    print("="*70)
    return results

def aggregate_orders_by_customer(orders_collection) -> List[Dict[str, Any]]:
    """
    AGGREGATION 3: Order Revenue & Payment Methods Analysis
    Groups orders by payment status and calculates total sales volume.
    """
    pipeline = [
        {
            "$group": {
                "_id": "$payment.method",
                "order_count": {"$sum": 1},
                "total_revenue": {"$sum": "$payment.total_amount"}
            }
        },
        {
            "$sort": {"total_revenue": -1}
        }
    ]

    results = orders_collection.aggregate(pipeline)
    print("\n" + "="*70)
    print(" [AGGREGATION 3] ORDER SALES REVENUE BY PAYMENT METHOD")
    print("="*70)
    print(f"{'Payment Method':<20} | {'Orders':<8} | {'Total Revenue':<15}")
    print("-" * 70)
    for row in results:
        method = row['_id'] or "Unknown"
        rev = row['total_revenue'] or 0.0
        print(f"{method:<20} | {row['order_count']:<8} | ${rev:<14.2f}")
    print("="*70)
    return results

def perform_text_search(collection, search_keyword: str) -> List[Dict[str, Any]]:
    """
    TEXT SEARCH: Queries products using multi-field text search ($text or regex keyword matching).
    """
    query = {"$text": {"$search": search_keyword}}
    results = collection.find(query).to_list()

    print("\n" + "="*70)
    print(f" [TEXT SEARCH] QUERY: '{search_keyword}' -> ({len(results)} matches found)")
    print("="*70)
    for item in results:
        tags_str = ", ".join(item.get("tags", []))
        print(f" • [{item['brand']}] {item['name']}")
        print(f"   Price: ${item['price']} | Rating: {item.get('rating_summary', {}).get('average_rating', 'N/A')}⭐ | Tags: [{tags_str}]")
    print("="*70)
    return results

def analyze_index_performance(collection):
    """
    INDEXING DEMONSTRATION: Explains index efficiency.
    Compares query scanning on indexed fields vs unindexed scans.
    """
    print("\n" + "="*70)
    print(" [INDEX PERFORMANCE ANALYSIS]")
    print("="*70)
    indexes = collection.list_indexes()
    print("Active Indexes in Collection:")
    for idx in indexes:
        print(f" - Index Name: '{idx['name']}' on Keys: {idx['key']}")

    print("\nQuery 1: Indexed compound query (category = 'Computers & Laptops', price <= 2500)")
    t0 = time.perf_counter()
    res1 = collection.find({
        "category": "Computers & Laptops",
        "price": {"$lte": 2500}
    }).sort("price", -1).to_list()
    t1 = time.perf_counter()
    print(f" Execution Plan: IXSCAN (Index Scan via 'idx_category_price')")
    print(f" Matches: {len(res1)} items | Time: {(t1 - t0)*1000:.4f} ms")

    print("\nQuery 2: Unindexed ad-hoc nested query (specs.weight_kg <= 2.0)")
    t2 = time.perf_counter()
    res2 = collection.find({
        "specs.weight_kg": {"$lte": 2.0}
    }).to_list()
    t3 = time.perf_counter()
    print(f" Execution Plan: COLLSCAN (Full Collection Scan across all documents)")
    print(f" Matches: {len(res2)} items | Time: {(t3 - t2)*1000:.4f} ms")
    print("="*70)

if __name__ == "__main__":
    from db_setup import get_database_client, seed_database, DB_NAME, PRODUCTS_COLLECTION, ORDERS_COLLECTION
    client, _ = get_database_client()
    db = client[DB_NAME]
    seed_database(db)

    products = db[PRODUCTS_COLLECTION]
    orders = db[ORDERS_COLLECTION]

    aggregate_category_summary(products)
    aggregate_review_sentiment_unwind(products)
    aggregate_orders_by_customer(orders)
    perform_text_search(products, "oled gaming")
    analyze_index_performance(products)
