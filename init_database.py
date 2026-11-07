import sqlite3
import openpyxl
import os
import json
from datetime import datetime, date

DB_PATH = r"D:\yash\Website\morbi_swm.db"
EXCEL_PATH = r"D:\yash\Website\Tractor_SSI_Annual_Dashboard.xlsx"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Table: wards
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS wards (
        ward_id INTEGER PRIMARY KEY,
        ward_name TEXT NOT NULL,
        ssi_name TEXT NOT NULL,
        ssi_phone TEXT NOT NULL,
        si_name TEXT DEFAULT '',
        si_phone TEXT DEFAULT '',
        subject TEXT DEFAULT 'Sankalp Patra Activity in School',
        remarks TEXT DEFAULT ''
    )
    """)

    # Table: tractors (master list: 2 tractors per ward)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tractors (
        tractor_id INTEGER PRIMARY KEY AUTOINCREMENT,
        ward_id INTEGER NOT NULL,
        tractor_label TEXT NOT NULL,
        plate_number TEXT NOT NULL,
        driver_phone TEXT DEFAULT '',
        status TEXT DEFAULT 'સક્રિય',
        FOREIGN KEY (ward_id) REFERENCES wards (ward_id)
    )
    """)

    # Table: daily_entries (trip records by date and tractor)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS daily_entries (
        entry_id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_date TEXT NOT NULL,
        tractor_id INTEGER NOT NULL,
        ward_id INTEGER NOT NULL,
        plate_number TEXT NOT NULL,
        driver_phone TEXT DEFAULT '',
        trip1 INTEGER DEFAULT 0,
        trip2 INTEGER DEFAULT 0,
        trip3 INTEGER DEFAULT 0,
        total_trips INTEGER DEFAULT 0,
        attendance TEXT DEFAULT 'હાજર',
        remarks TEXT DEFAULT '',
        updated_by TEXT DEFAULT 'SSI',
        updated_at TEXT NOT NULL,
        UNIQUE(entry_date, tractor_id)
    )
    """)

    # Table: users (for login / roles)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_pin TEXT NOT NULL,
        role TEXT NOT NULL,
        ward_id INTEGER,
        display_name TEXT NOT NULL
    )
    """)

    conn.commit()
    conn.close()
    print("Database tables initialized successfully.")

if __name__ == "__main__":
    init_db()

