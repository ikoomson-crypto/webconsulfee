import sqlite3
from datetime import datetime

DATABASE = 'payment_system.db'


def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    db = get_db()

    # Create suppliers table
    db.execute('''
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT,
            telephone TEXT,
            bank_name TEXT,
            account_number TEXT,
            account_name TEXT,
            swift_code TEXT,
            bank_address TEXT,
            id_type TEXT,
            id_number TEXT,
            street_address TEXT,
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Create services table with supplier_id and status
    db.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER,
            service_description TEXT NOT NULL,
            rate REAL,
            total REAL,
            status TEXT DEFAULT 'Pending',
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (supplier_id) REFERENCES suppliers (id)
        )
    ''')

    # Create enhanced payments table
    db.execute('''
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER,
            service_id INTEGER,
            amount REAL,
            payment_method TEXT,
            reference_number TEXT,
            payment_date DATE,
            notes TEXT,
            status TEXT DEFAULT 'Completed',
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (supplier_id) REFERENCES suppliers (id),
            FOREIGN KEY (service_id) REFERENCES services (id)
        )
    ''')

    db.commit()
    db.close()