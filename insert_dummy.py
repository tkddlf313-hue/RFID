"""
테스트용 더미 데이터 삽입 스크립트
실행: python insert_dummy.py
"""
import sqlite3
import datetime
import random

DB_PATH = "rfid_inventory.db"

ITEMS = [
    ("SUS304 판재 2T",    "LOT-2024-001"),
    ("SUS316 환봉 Ø20",   "LOT-2024-002"),
    ("알루미늄 앵글 40x40","LOT-2024-003"),
    ("SS400 각관 50x50",  "LOT-2024-004"),
    ("SPCC 냉연코일 1.2T", "LOT-2024-005"),
    ("황동봉 Ø15",        "LOT-2024-006"),
    ("철판 3.2T",         "LOT-2024-007"),
    ("동파이프 Ø25",      "LOT-2024-008"),
    ("ABS 플라스틱 판",   "LOT-2024-009"),
    ("탄소강 볼트 M12",   "LOT-2024-010"),
]

EPCS = [
    "A1B2C3D4E5F60001",
    "A1B2C3D4E5F60002",
    "A1B2C3D4E5F60003",
    "A1B2C3D4E5F60004",
    "A1B2C3D4E5F60005",
    "A1B2C3D4E5F60006",
    "A1B2C3D4E5F60007",
    "A1B2C3D4E5F60008",
    "A1B2C3D4E5F60009",
    "A1B2C3D4E5F60010",
]

def random_date(days_ago_max=30):
    delta = datetime.timedelta(
        days=random.randint(0, days_ago_max),
        hours=random.randint(8, 17),
        minutes=random.randint(0, 59),
    )
    return (datetime.datetime.now() - delta).strftime("%Y-%m-%d %H:%M:%S")

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# 테이블 없으면 생성
cur.executescript("""
    CREATE TABLE IF NOT EXISTS tags (
        tag_id      TEXT PRIMARY KEY,
        item_name   TEXT NOT NULL,
        lot_number  TEXT NOT NULL,
        quantity    INTEGER DEFAULT 0,
        issued_at   TEXT NOT NULL,
        location    TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS scan_log (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        tag_id      TEXT NOT NULL,
        scanned_at  TEXT NOT NULL,
        matched     INTEGER DEFAULT 0
    );
""")

# 기존 더미 데이터 초기화 여부 확인
existing = cur.execute("SELECT COUNT(*) FROM tags").fetchone()[0]
if existing > 0:
    answer = input(f"이미 {existing}건의 태그가 있습니다. 추가로 삽입할까요? (y/n): ")
    if answer.strip().lower() != "y":
        print("취소됨.")
        conn.close()
        exit()

inserted = 0
for epc, (item_name, lot_number) in zip(EPCS, ITEMS):
    qty = random.randint(10, 500)
    issued_at = random_date()
    try:
        cur.execute(
            "INSERT OR IGNORE INTO tags(tag_id, item_name, lot_number, quantity, issued_at, location) VALUES (?,?,?,?,?,?)",
            (epc, item_name, lot_number, qty, issued_at, f"A-{random.randint(1,5)}-{random.randint(1,10):02d}"),
        )
        if cur.rowcount:
            print(f"  [OK] {epc} | {item_name:<20} | {lot_number} | {qty:>4} EA | {issued_at}")
            inserted += 1
        else:
            print(f"  [SKIP] {epc} 이미 존재 -- 건너뜀")
    except Exception as e:
        print(f"  [ERR] {e}")

conn.commit()
conn.close()
print(f"\n[완료] {inserted}건 삽입됨 -> rfid_inventory.db")
