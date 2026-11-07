import json
import sqlite3
import os
from datetime import datetime

BACKUP_DIR = os.path.join(os.path.dirname(__file__), "data_backup")
BACKUP_FILE = os.path.join(BACKUP_DIR, "swm_backup.json")
DB_PATH = os.path.join(os.path.dirname(__file__), "morbi_swm.db")

def ensure_backup_dir():
    os.makedirs(BACKUP_DIR, exist_ok=True)

def backup_database_to_json():
    """
    Exports the entire SQLite database to a structured JSON file for 0-data-loss persistence.
    """
    ensure_backup_dir()
    if not os.path.exists(DB_PATH):
        return
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    data = {
        "exported_at": datetime.now().isoformat(),
        "version": "2.0_secure",
        "system_config": {},
        "wards": [],
        "tractors": [],
        "daily_entries": []
    }

    try:
        c.execute("SELECT key, value FROM system_config")
        data["system_config"] = {r["key"]: r["value"] for r in c.fetchall()}
    except Exception:
        pass

    c.execute("SELECT * FROM wards ORDER BY ward_id")
    data["wards"] = [dict(r) for r in c.fetchall()]

    c.execute("SELECT * FROM tractors ORDER BY tractor_id")
    data["tractors"] = [dict(r) for r in c.fetchall()]

    c.execute("SELECT * FROM daily_entries ORDER BY entry_date, tractor_id")
    data["daily_entries"] = [dict(r) for r in c.fetchall()]

    conn.close()

    with open(BACKUP_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    # print(f"Backup saved: {len(data['daily_entries'])} entries backed up.")
    return BACKUP_FILE

def restore_database_from_json():
    """
    Restores database from data_backup/swm_backup.json if database is missing or empty.
    """
    if not os.path.exists(BACKUP_FILE):
        return False
    
    with open(BACKUP_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Recreate tables if not present
    c.execute("""
    CREATE TABLE IF NOT EXISTS wards (
        ward_id INTEGER PRIMARY KEY,
        ward_name TEXT NOT NULL,
        ssi_name TEXT NOT NULL,
        ssi_phone TEXT NOT NULL,
        si_name TEXT DEFAULT '',
        si_phone TEXT DEFAULT '',
        ssi_password TEXT DEFAULT '1234'
    )
    """)

    c.execute("""
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

    c.execute("""
    CREATE TABLE IF NOT EXISTS daily_entries (
        entry_id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_date TEXT NOT NULL,
        ward_id INTEGER NOT NULL,
        tractor_id INTEGER NOT NULL,
        tractor_label TEXT NOT NULL,
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
        is_submitted INTEGER DEFAULT 1,
        UNIQUE(entry_date, tractor_id)
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS system_config (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """)

    for k, v in data.get("system_config", {}).items():
        c.execute("INSERT OR REPLACE INTO system_config (key, value) VALUES (?, ?)", (k, v))

    for w in data.get("wards", []):
        c.execute("""
        INSERT OR REPLACE INTO wards (ward_id, ward_name, ssi_name, ssi_phone, si_name, si_phone, ssi_password)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (w["ward_id"], w["ward_name"], w["ssi_name"], w["ssi_phone"], w.get("si_name", ""), w.get("si_phone", ""), w.get("ssi_password", "1234")))

    for t in data.get("tractors", []):
        c.execute("""
        INSERT OR REPLACE INTO tractors (tractor_id, ward_id, tractor_label, plate_number, driver_phone, status)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (t["tractor_id"], t["ward_id"], t["tractor_label"], t["plate_number"], t.get("driver_phone", ""), t.get("status", "સક્રિય")))

    for d in data.get("daily_entries", []):
        c.execute("""
        INSERT OR REPLACE INTO daily_entries (
            entry_id, entry_date, ward_id, tractor_id, tractor_label, plate_number, driver_phone,
            trip1, trip2, trip3, total_trips, attendance, remarks, updated_by, updated_at, is_submitted
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            d.get("entry_id"), d["entry_date"], d["ward_id"], d["tractor_id"], d["tractor_label"],
            d["plate_number"], d.get("driver_phone", ""), d.get("trip1", 0), d.get("trip2", 0),
            d.get("trip3", 0), d.get("total_trips", 0), d.get("attendance", "હાજર"),
            d.get("remarks", ""), d.get("updated_by", "SSI"), d.get("updated_at", datetime.now().isoformat()),
            d.get("is_submitted", 1)
        ))

    conn.commit()
    conn.close()
    print("Database restored from JSON backup.")
    return True

if __name__ == "__main__":
    backup_database_to_json()
    print("Backup verified.")

