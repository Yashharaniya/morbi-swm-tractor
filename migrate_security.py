import sqlite3
import os

DB_PATH = r"D:\yash\Website\morbi_swm.db"

def migrate():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Check if ssi_password exists in wards
    c.execute("PRAGMA table_info(wards)")
    cols = [r[1] for r in c.fetchall()]
    if "ssi_password" not in cols:
        print("Adding ssi_password column to wards table...")
        c.execute("ALTER TABLE wards ADD COLUMN ssi_password TEXT DEFAULT '1234'")
        # Give each ward default pin '1234'
        c.execute("UPDATE wards SET ssi_password = '1234'")
        conn.commit()
        print("ssi_password column added.")
    else:
        print("ssi_password column already exists.")

    # Check master admin password in a config table
    c.execute("""
    CREATE TABLE IF NOT EXISTS system_config (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """)
    c.execute("INSERT OR REPLACE INTO system_config (key, value) VALUES ('admin_pin', '12345678')")
    c.execute("INSERT OR REPLACE INTO system_config (key, value) VALUES ('admin_hint', '1 - 8')")
    
    conn.commit()
    conn.close()
    print("Migration finished successfully.")

if __name__ == "__main__":
    migrate()

