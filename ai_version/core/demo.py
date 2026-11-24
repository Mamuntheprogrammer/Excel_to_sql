"""Generate a small demo database so the app is useful on first launch."""

from __future__ import annotations

import random
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

FIRST_NAMES = [
    "Ava", "Liam", "Mia", "Noah", "Sofia", "Ethan", "Isla", "Lucas", "Amara",
    "Mateo", "Zara", "Oliver", "Yuki", "Idris", "Nora", "Ravi", "Elena",
    "Kai", "Sana", "Hugo", "Leila", "Oscar", "Aisha", "Felix", "Priya",
]
LAST_NAMES = [
    "Chen", "Patel", "Garcia", "Kim", "Müller", "Silva", "Okafor", "Novak",
    "Haddad", "Rossi", "Nguyen", "Kowalski", "Andersson", "Tanaka", "Mensah",
    "Lopez", "Ibrahim", "Dubois", "Petrov", "Ali",
]
CITIES = [
    ("London", "UK"), ("Berlin", "Germany"), ("Toronto", "Canada"),
    ("Sydney", "Australia"), ("Tokyo", "Japan"), ("Austin", "USA"),
    ("Lisbon", "Portugal"), ("Nairobi", "Kenya"), ("Mumbai", "India"),
    ("São Paulo", "Brazil"), ("Paris", "France"), ("Seoul", "South Korea"),
]
PRODUCTS = [
    ("Laptop Pro 14", "Electronics", 1899.00),
    ("Wireless Mouse", "Electronics", 39.90),
    ("Mechanical Keyboard", "Electronics", 129.00),
    ("27\" 4K Monitor", "Electronics", 549.00),
    ("USB-C Dock", "Electronics", 189.00),
    ("Noise-Cancel Headphones", "Audio", 299.00),
    ("Studio Microphone", "Audio", 149.00),
    ("Bluetooth Speaker", "Audio", 89.50),
    ("Ergonomic Chair", "Furniture", 449.00),
    ("Standing Desk", "Furniture", 699.00),
    ("Desk Lamp", "Furniture", 59.00),
    ("Bookshelf", "Furniture", 219.00),
    ("Notebook Set", "Stationery", 14.50),
    ("Gel Pen Pack", "Stationery", 8.90),
    ("Whiteboard", "Stationery", 79.00),
    ("Coffee Maker", "Kitchen", 129.00),
    ("Kettle", "Kitchen", 49.00),
    ("Water Bottle", "Kitchen", 22.00),
    ("Phones Pro", "Audio", 249.00),
    ("Webcam 4K", "Electronics", 159.00),
    ("Monitor Arm", "Furniture", 99.00),
    ("Cable Organizer", "Accessories", 12.00),
    ("Laptop Sleeve", "Accessories", 34.00),
    ("Travel Backpack", "Accessories", 89.00),
    ("Smart Watch", "Electronics", 329.00),
]
STATUSES = ["completed", "completed", "completed", "shipped", "pending", "cancelled"]


def create_demo_database(path: str) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()

    rng = random.Random(42)
    conn = sqlite3.connect(target)
    try:
        conn.executescript(
            """
            CREATE TABLE customers (
                id          INTEGER PRIMARY KEY,
                name        TEXT NOT NULL,
                email       TEXT NOT NULL UNIQUE,
                city        TEXT,
                country     TEXT,
                joined_at   TEXT
            );
            CREATE TABLE products (
                id          INTEGER PRIMARY KEY,
                name        TEXT NOT NULL,
                category    TEXT,
                price       REAL,
                stock       INTEGER
            );
            CREATE TABLE orders (
                id          INTEGER PRIMARY KEY,
                customer_id INTEGER REFERENCES customers(id),
                product_id  INTEGER REFERENCES products(id),
                quantity    INTEGER,
                unit_price  REAL,
                status      TEXT,
                ordered_at  TEXT
            );
            CREATE INDEX idx_orders_customer ON orders(customer_id);
            CREATE INDEX idx_orders_product ON orders(product_id);
            CREATE VIEW order_summary AS
                SELECT o.id AS order_id,
                       c.name AS customer,
                       c.country,
                       p.name AS product,
                       p.category,
                       o.quantity,
                       ROUND(o.quantity * o.unit_price, 2) AS total,
                       o.status,
                       o.ordered_at
                FROM orders o
                JOIN customers c ON c.id = o.customer_id
                JOIN products p ON p.id = o.product_id;
            """
        )

        start = datetime(2025, 1, 1)
        customers = []
        for index in range(1, 41):
            first = rng.choice(FIRST_NAMES)
            last = rng.choice(LAST_NAMES)
            city, country = rng.choice(CITIES)
            email = f"{first.lower()}.{last.lower().replace('ü', 'u')}{index}@example.com"
            joined = start + timedelta(days=rng.randint(0, 300))
            customers.append(
                (index, f"{first} {last}", email, city, country, joined.date().isoformat())
            )
        conn.executemany("INSERT INTO customers VALUES (?,?,?,?,?,?)", customers)

        products = []
        for index, (name, category, price) in enumerate(PRODUCTS, start=1):
            products.append((index, name, category, price, rng.randint(4, 180)))
        conn.executemany("INSERT INTO products VALUES (?,?,?,?,?)", products)

        orders = []
        for index in range(1, 161):
            customer_id = rng.randint(1, len(customers))
            product_id = rng.randint(1, len(products))
            quantity = rng.randint(1, 5)
            unit_price = products[product_id - 1][3]
            status = rng.choice(STATUSES)
            ordered = start + timedelta(
                days=rng.randint(0, 330), hours=rng.randint(0, 23), minutes=rng.randint(0, 59)
            )
            orders.append(
                (
                    index,
                    customer_id,
                    product_id,
                    quantity,
                    unit_price,
                    status,
                    ordered.strftime("%Y-%m-%d %H:%M:%S"),
                )
            )
        conn.executemany("INSERT INTO orders VALUES (?,?,?,?,?,?,?)", orders)
        conn.commit()
    finally:
        conn.close()
