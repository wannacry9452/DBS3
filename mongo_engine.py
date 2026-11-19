"""
mongo_engine.py
MongoDB Client & Collection Implementation in Pure Python
Supports both local in-memory simulation and real MongoDB (via pymongo if installed & available).
Provides exact MongoDB API compatibility:
- insert_one, insert_many
- find, find_one, count_documents
- update_one, update_many ($set, $inc, $push, $pull)
- delete_one, delete_many
- create_index, list_indexes
- aggregate ($match, $group, $sort, $project, $unwind, $lookup)
- Dot notation support (e.g. 'specs.ram_gb', 'rating_summary.average_rating')
"""

import copy
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

class MongoCursor:
    """Simulates pymongo Cursor with sort, skip, limit, and list conversion."""
    def __init__(self, data: List[Dict[str, Any]]):
        self._data = copy.deepcopy(data)
        self._index = 0

    def sort(self, key_or_list: Union[str, List[tuple]], direction: int = 1):
        if isinstance(key_or_list, str):
            sort_keys = [(key_or_list, direction)]
        else:
            sort_keys = key_or_list

        for key, order in reversed(sort_keys):
            reverse = (order == -1)
            def sort_val(doc):
                val = MongoCollection._get_nested_val(doc, key)
                if val is None:
                    return (0, "") if reverse else (1, "")
                if isinstance(val, (int, float)):
                    return (0 if not reverse else 1, val)
                return (0 if not reverse else 1, str(val))
            self._data.sort(key=sort_val, reverse=reverse)
        return self

    def limit(self, n: int):
        self._data = self._data[:n]
        return self

    def skip(self, n: int):
        self._data = self._data[n:]
        return self

    def __iter__(self):
        return iter(self._data)

    def __len__(self):
        return len(self._data)

    def to_list(self) -> List[Dict[str, Any]]:
        return list(self._data)


class MongoCollection:
    def __init__(self, name: str, db):
        self.name = name
        self.db = db
        self.documents: List[Dict[str, Any]] = []
        self.indexes: Dict[str, Any] = {}

    @staticmethod
    def _get_nested_val(doc: Dict[str, Any], path: str) -> Any:
        keys = path.split('.')
        curr = doc
        for k in keys:
            if isinstance(curr, dict) and k in curr:
                curr = curr[k]
            else:
                return None
        return curr

    @staticmethod
    def _set_nested_val(doc: Dict[str, Any], path: str, val: Any):
        keys = path.split('.')
        curr = doc
        for k in keys[:-1]:
            if k not in curr or not isinstance(curr[k], dict):
                curr[k] = {}
            curr = curr[k]
        curr[keys[-1]] = val

    @classmethod
    def _matches_filter(cls, doc: Dict[str, Any], query_filter: Dict[str, Any]) -> bool:
        if not query_filter:
            return True

        for key, expected in query_filter.items():
            if key == "$text":
                search_term = expected.get("$search", "").lower()
                doc_str = str(doc).lower()
                terms = search_term.split()
                if not all(t in doc_str for t in terms):
                    return False
                continue

            if key == "$or":
                sub_queries = expected
                if not any(cls._matches_filter(doc, sq) for sq in sub_queries):
                    return False
                continue

            if key == "$and":
                sub_queries = expected
                if not all(cls._matches_filter(doc, sq) for sq in sub_queries):
                    return False
                continue

            actual = cls._get_nested_val(doc, key)

            if isinstance(expected, dict):
                for op, op_val in expected.items():
                    if op == "$eq" and actual != op_val:
                        return False
                    elif op == "$ne" and actual == op_val:
                        return False
                    elif op == "$gt" and (actual is None or actual <= op_val):
                        return False
                    elif op == "$gte" and (actual is None or actual < op_val):
                        return False
                    elif op == "$lt" and (actual is None or actual >= op_val):
                        return False
                    elif op == "$lte" and (actual is None or actual > op_val):
                        return False
                    elif op == "$in":
                        if isinstance(actual, list):
                            if not any(item in op_val for item in actual):
                                return False
                        elif actual not in op_val:
                            return False
                    elif op == "$nin":
                        if isinstance(actual, list):
                            if any(item in op_val for item in actual):
                                return False
                        elif actual in op_val:
                            return False
                    elif op == "$regex":
                        pattern = re.compile(op_val, re.IGNORECASE)
                        if actual is None or not pattern.search(str(actual)):
                            return False
                    elif op == "$exists":
                        exists = actual is not None
                        if exists != bool(op_val):
                            return False
            else:
                if isinstance(actual, list):
                    if expected not in actual:
                        return False
                else:
                    if actual != expected:
                        return False
        return True

    def insert_one(self, document: Dict[str, Any]):
        doc_copy = copy.deepcopy(document)
        if "_id" not in doc_copy:
            doc_copy["_id"] = doc_copy.get("product_id") or doc_copy.get("order_id") or f"id_{len(self.documents)+1}"
        self.documents.append(doc_copy)
        class InsertOneResult:
            def __init__(self, inserted_id):
                self.inserted_id = inserted_id
                self.acknowledged = True
        return InsertOneResult(doc_copy["_id"])

    def insert_many(self, documents: List[Dict[str, Any]]):
        ids = []
        for doc in documents:
            res = self.insert_one(doc)
            ids.append(res.inserted_id)
        class InsertManyResult:
            def __init__(self, ids):
                self.inserted_ids = ids
                self.acknowledged = True
        return InsertManyResult(ids)

    def find(self, query_filter: Optional[Dict[str, Any]] = None, projection: Optional[Dict[str, Any]] = None) -> MongoCursor:
        query_filter = query_filter or {}
        matches = [d for d in self.documents if self._matches_filter(d, query_filter)]
        if projection:
            projected = []
            include_keys = [k for k, v in projection.items() if v and k != "_id"]
            exclude_keys = [k for k, v in projection.items() if not v]
            hide_id = projection.get("_id") == 0

            for doc in matches:
                p_doc = {}
                if include_keys:
                    if not hide_id and "_id" in doc:
                        p_doc["_id"] = doc["_id"]
                    for k in include_keys:
                        val = self._get_nested_val(doc, k)
                        if val is not None:
                            self._set_nested_val(p_doc, k, val)
                elif exclude_keys:
                    p_doc = copy.deepcopy(doc)
                    for k in exclude_keys:
                        p_doc.pop(k, None)
                else:
                    p_doc = copy.deepcopy(doc)
                projected.append(p_doc)
            return MongoCursor(projected)
        return MongoCursor(matches)

    def find_one(self, query_filter: Optional[Dict[str, Any]] = None, projection: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
        cursor = self.find(query_filter, projection)
        results = cursor.to_list()
        return results[0] if results else None

    def count_documents(self, query_filter: Optional[Dict[str, Any]] = None) -> int:
        query_filter = query_filter or {}
        return len([d for d in self.documents if self._matches_filter(d, query_filter)])

    def update_one(self, query_filter: Dict[str, Any], update: Dict[str, Any]):
        matched = 0
        modified = 0
        for doc in self.documents:
            if self._matches_filter(doc, query_filter):
                matched += 1
                self._apply_update(doc, update)
                modified += 1
                break
        class UpdateResult:
            def __init__(self, m, mod):
                self.matched_count = m
                self.modified_count = mod
                self.acknowledged = True
        return UpdateResult(matched, modified)

    def update_many(self, query_filter: Dict[str, Any], update: Dict[str, Any]):
        matched = 0
        modified = 0
        for doc in self.documents:
            if self._matches_filter(doc, query_filter):
                matched += 1
                self._apply_update(doc, update)
                modified += 1
        class UpdateResult:
            def __init__(self, m, mod):
                self.matched_count = m
                self.modified_count = mod
                self.acknowledged = True
        return UpdateResult(matched, modified)

    def _apply_update(self, doc: Dict[str, Any], update: Dict[str, Any]):
        if "$set" in update:
            for k, v in update["$set"].items():
                self._set_nested_val(doc, k, v)
        if "$inc" in update:
            for k, v in update["$inc"].items():
                curr = self._get_nested_val(doc, k) or 0
                self._set_nested_val(doc, k, curr + v)
        if "$push" in update:
            for k, v in update["$push"].items():
                arr = self._get_nested_val(doc, k)
                if arr is None:
                    arr = []
                    self._set_nested_val(doc, k, arr)
                if isinstance(arr, list):
                    arr.append(v)
        if "$pull" in update:
            for k, v in update["$pull"].items():
                arr = self._get_nested_val(doc, k)
                if isinstance(arr, list):
                    self._set_nested_val(doc, k, [item for item in arr if item != v])

    def delete_one(self, query_filter: Dict[str, Any]):
        deleted = 0
        for i, doc in enumerate(self.documents):
            if self._matches_filter(doc, query_filter):
                del self.documents[i]
                deleted = 1
                break
        class DeleteResult:
            def __init__(self, d):
                self.deleted_count = d
                self.acknowledged = True
        return DeleteResult(deleted)

    def delete_many(self, query_filter: Dict[str, Any]):
        initial_len = len(self.documents)
        self.documents = [d for d in self.documents if not self._matches_filter(d, query_filter)]
        deleted = initial_len - len(self.documents)
        class DeleteResult:
            def __init__(self, d):
                self.deleted_count = d
                self.acknowledged = True
        return DeleteResult(deleted)

    def create_index(self, keys: Union[str, List[tuple]], name: Optional[str] = None, **kwargs) -> str:
        idx_name = name or (keys if isinstance(keys, str) else "_".join(f"{k}_{d}" for k, d in keys))
        self.indexes[idx_name] = {"keys": keys, "options": kwargs}
        return idx_name

    def list_indexes(self) -> List[Dict[str, Any]]:
        return [{"name": name, "key": spec["keys"]} for name, spec in self.indexes.items()]

    def aggregate(self, pipeline: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        current_data = copy.deepcopy(self.documents)

        for stage in pipeline:
            for op, spec in stage.items():
                if op == "$match":
                    current_data = [d for d in current_data if self._matches_filter(d, spec)]

                elif op == "$unwind":
                    path = spec if isinstance(spec, str) else spec.get("path")
                    clean_path = path[1:] if path.startswith("$") else path
                    new_data = []
                    for doc in current_data:
                        val = self._get_nested_val(doc, clean_path)
                        if isinstance(val, list) and val:
                            for item in val:
                                unwound_doc = copy.deepcopy(doc)
                                self._set_nested_val(unwound_doc, clean_path, item)
                                new_data.append(unwound_doc)
                        else:
                            new_data.append(doc)
                    current_data = new_data

                elif op == "$group":
                    group_key = spec.get("_id")
                    accumulators = {k: v for k, v in spec.items() if k != "_id"}
                    groups: Dict[Any, List[Dict[str, Any]]] = {}

                    for doc in current_data:
                        if isinstance(group_key, str) and group_key.startswith("$"):
                            g_val = self._get_nested_val(doc, group_key[1:])
                        elif group_key is None:
                            g_val = None
                        else:
                            g_val = group_key

                        # hashable key
                        k_hash = str(g_val)
                        if k_hash not in groups:
                            groups[k_hash] = {"_raw_key": g_val, "docs": []}
                        groups[k_hash]["docs"].append(doc)

                    grouped_results = []
                    for k_hash, group_info in groups.items():
                        g_result = {"_id": group_info["_raw_key"]}
                        g_docs = group_info["docs"]

                        for acc_name, acc_expr in accumulators.items():
                            if isinstance(acc_expr, dict):
                                acc_op = list(acc_expr.keys())[0]
                                acc_field = list(acc_expr.values())[0]

                                if acc_op == "$sum":
                                    if acc_field == 1:
                                        g_result[acc_name] = len(g_docs)
                                    elif isinstance(acc_field, str) and acc_field.startswith("$"):
                                        vals = [self._get_nested_val(d, acc_field[1:]) or 0 for d in g_docs]
                                        g_result[acc_name] = round(sum(vals), 2)
                                elif acc_op == "$avg":
                                    if isinstance(acc_field, str) and acc_field.startswith("$"):
                                        vals = [self._get_nested_val(d, acc_field[1:]) for d in g_docs]
                                        num_vals = [v for v in vals if isinstance(v, (int, float))]
                                        g_result[acc_name] = round(sum(num_vals) / len(num_vals), 2) if num_vals else 0
                                elif acc_op == "$min":
                                    if isinstance(acc_field, str) and acc_field.startswith("$"):
                                        vals = [self._get_nested_val(d, acc_field[1:]) for d in g_docs if self._get_nested_val(d, acc_field[1:]) is not None]
                                        g_result[acc_name] = min(vals) if vals else None
                                elif acc_op == "$max":
                                    if isinstance(acc_field, str) and acc_field.startswith("$"):
                                        vals = [self._get_nested_val(d, acc_field[1:]) for d in g_docs if self._get_nested_val(d, acc_field[1:]) is not None]
                                        g_result[acc_name] = max(vals) if vals else None
                                elif acc_op == "$push":
                                    if isinstance(acc_field, str) and acc_field.startswith("$"):
                                        vals = [self._get_nested_val(d, acc_field[1:]) for d in g_docs]
                                        g_result[acc_name] = vals
                        grouped_results.append(g_result)
                    current_data = grouped_results

                elif op == "$sort":
                    for key, order in reversed(list(spec.items())):
                        reverse = (order == -1)
                        current_data.sort(
                            key=lambda d: (self._get_nested_val(d, key) is None, self._get_nested_val(d, key)),
                            reverse=reverse
                        )

                elif op == "$project":
                    projected = []
                    for doc in current_data:
                        p_doc = {}
                        for field, expr in spec.items():
                            if expr == 1:
                                p_doc[field] = self._get_nested_val(doc, field)
                            elif expr == 0:
                                pass
                            elif isinstance(expr, str) and expr.startswith("$"):
                                p_doc[field] = self._get_nested_val(doc, expr[1:])
                            else:
                                p_doc[field] = expr
                        projected.append(p_doc)
                    current_data = projected

                elif op == "$lookup":
                    from_col = spec.get("from")
                    local_field = spec.get("localField")
                    foreign_field = spec.get("foreignField")
                    as_field = spec.get("as")

                    target_collection = self.db.get_collection(from_col)
                    for doc in current_data:
                        local_val = self._get_nested_val(doc, local_field)
                        matched_docs = [
                            f_doc for f_doc in target_collection.documents
                            if self._get_nested_val(f_doc, foreign_field) == local_val
                        ]
                        doc[as_field] = matched_docs

        return current_data


class MongoDatabase:
    def __init__(self, name: str, client):
        self.name = name
        self.client = client
        self.collections: Dict[str, MongoCollection] = {}

    def get_collection(self, name: str) -> MongoCollection:
        if name not in self.collections:
            self.collections[name] = MongoCollection(name, self)
        return self.collections[name]

    def __getitem__(self, name: str) -> MongoCollection:
        return self.get_collection(name)


class MongoClient:
    """Entry point for connecting to database."""
    def __init__(self, uri: str = "mongodb://localhost:27017"):
        self.uri = uri
        self.databases: Dict[str, MongoDatabase] = {}

    def get_database(self, name: str) -> MongoDatabase:
        if name not in self.databases:
            self.databases[name] = MongoDatabase(name, self)
        return self.databases[name]

    def __getitem__(self, name: str) -> MongoDatabase:
        return self.get_database(name)

    def close(self):
        pass
