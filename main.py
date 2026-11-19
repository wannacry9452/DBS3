# main.py — NoSQL төслийн нэгдсэн гүйцэтгэгч
import sys

class DualLogger:
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

def run_all():
    print("=" * 75)
    print(" БИЕ ДААЛТЫН АЖИЛ №3: NoSQL ӨГӨГДЛИЙН САНГУУД (MongoDB)")
    print(" Цахим худалдааны бүтээгдэхүүний каталогийн туршилтын систем")
    print("=" * 75)

    print("\n--- 1. ӨГӨГДЛИЙН САНГ ЭХЛҮҮЛЭХ БА 25 БАРИМТ ОРУУЛАХ (init_db.py) ---")
    import init_db

    print("\n--- 2. БҮРЭН CRUD ҮЙЛДЛҮҮДИЙН ГҮЙЦЭТГЭЛ (crud_operations.py) ---")
    import crud_operations

    print("\n--- 3. ДЭВШИЛТЭТ БОЛОМЖ: ИНДЕКС, ТЕКСТ ХАЙЛТ, AGGREGATION (advanced_features.py) ---")
    import advanced_features

    print("\n" + "=" * 75)
    print(" БҮХ ТУРШИЛТ АМЖИЛТТАЙ ГҮЙЦЭТГЭГДЛЭЭ!")
    print("=" * 75)

if __name__ == "__main__":
    logger = DualLogger("test_run.log")
    sys.stdout = logger
    try:
        run_all()
    finally:
        sys.stdout = logger.terminal
        logger.close()
