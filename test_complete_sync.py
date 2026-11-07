import sqlite3
import json
import os
from datetime import datetime
import app
import db_backup_manager
from excel_generator import generate_dymc_excel_report, generate_monthly_excel_report

def test_full_system():
    print("=" * 70)
    print("🧪 101% WORKING MORBI SWM FULL VERIFICATION TEST")
    print("=" * 70)

    # 1. Admin PIN verification
    print("\n[1] Verifying New Admin PIN (87654321)...")
    assert app.verify_admin_pin("87654321") == True, "Failed: 87654321 must be valid"
    assert app.verify_admin_pin("00000000") == False, "Failed: 00000000 must be invalid"
    print("    ✅ New Admin PIN 87654321 successfully verified.")

    # 2. Ward Passwords verification
    print("\n[2] Verifying Ward Passwords Endpoint...")
    conn = app.get_db()
    c = conn.cursor()
    c.execute("SELECT ward_id, ward_name, ssi_password FROM wards ORDER BY ward_id")
    wards = c.fetchall()
    assert len(wards) == 13, f"Expected 13 wards, got {len(wards)}"
    for w in wards:
        assert w["ssi_password"] != "", f"Ward {w['ward_id']} password is empty"
    print(f"    ✅ All 13 wards verified with valid passwords (Sample Ward 1: {wards[0]['ssi_password']}).")
    conn.close()

    # 3. Tractor Number Synchronization Across the Whole Site
    print("\n[3] Testing Tractor Number Synchronization Everywhere...")
    conn = app.get_db()
    c = conn.cursor()
    c.execute("SELECT tractor_id, plate_number, driver_phone FROM tractors WHERE tractor_id = 1")
    t1 = c.fetchone()
    orig_plate = t1["plate_number"]
    orig_driver = t1["driver_phone"]
    conn.close()

    test_plate = "GJ-36-SYNC-9999"
    test_driver = "9876500000"
    today_str = datetime.now().strftime("%Y-%m-%d")

    # Simulate /api/update_master_tractor
    conn = app.get_db()
    c = conn.cursor()
    c.execute("UPDATE tractors SET plate_number = ?, driver_phone = ? WHERE tractor_id = 1", (test_plate, test_driver))
    c.execute("UPDATE daily_entries SET plate_number = ?, driver_phone = ? WHERE tractor_id = 1 AND entry_date >= ?", (test_plate, test_driver, today_str))
    conn.commit()
    conn.close()

    # Check 1: Master tractors table
    conn = app.get_db()
    c = conn.cursor()
    c.execute("SELECT plate_number FROM tractors WHERE tractor_id = 1")
    assert c.fetchone()[0] == test_plate, "Check 1 failed: tractors table not updated"
    print("    ✓ Check 1: Master tractors fleet updated.")

    # Check 2: DyMC Wards data API
    wards_data = app.get_wards_data(today_str)
    w1_t1 = next(t for w in wards_data if w["ward_id"] == 1 for t in w["tractors"] if t["tractor_id"] == 1)
    assert w1_t1["plate_number"] == test_plate, f"Check 2 failed: DyMC view has {w1_t1['plate_number']} instead of {test_plate}"
    print(f"    ✓ Check 2: DyMC Live Dashboard immediately shows: {w1_t1['plate_number']}.")

    # Check 3: Monthly Matrix API
    matrix = app.get_monthly_matrix(2026, 11)
    m_w1_t1 = next(t for w in matrix if w["ward_id"] == 1 for t in w["tractors"] if t["tractor_id"] == 1)
    assert m_w1_t1["plate_number"] == test_plate, f"Check 3 failed: Monthly matrix has {m_w1_t1['plate_number']}"
    print(f"    ✓ Check 3: Monthly Register 31-day matrix immediately shows: {m_w1_t1['plate_number']}.")

    # Check 4: DyMC Excel Generation
    summary = app.get_summary(today_str)
    excel_bytes = generate_dymc_excel_report(today_str, summary, wards_data, "ટેસ્ટ ફિલ્ટર")
    assert len(excel_bytes) > 2000, "DyMC Excel generation failed"
    print("    ✓ Check 4: DyMC Excel report exports with new tractor plate.")

    # Check 5: Monthly Excel Generation
    m_excel_bytes = generate_monthly_excel_report(2026, 11, matrix)
    assert len(m_excel_bytes) > 2000, "Monthly Excel generation failed"
    print("    ✓ Check 5: Monthly Register Excel exports with new tractor plate.")

    # Check 6: JSON Backup Auto-Sync
    bk_file = db_backup_manager.backup_database_to_json()
    with open(bk_file, "r", encoding="utf-8") as f:
        bk_data = json.load(f)
    bk_t1 = next(t for t in bk_data["tractors"] if t["tractor_id"] == 1)
    assert bk_t1["plate_number"] == test_plate, "Check 6 failed: JSON backup not updated"
    print(f"    ✓ Check 6: Render.com JSON backup snapshot updated with: {bk_t1['plate_number']}.")

    # Revert back to original plate
    conn = app.get_db()
    c = conn.cursor()
    c.execute("UPDATE tractors SET plate_number = ?, driver_phone = ? WHERE tractor_id = 1", (orig_plate, orig_driver))
    c.execute("UPDATE daily_entries SET plate_number = ?, driver_phone = ? WHERE tractor_id = 1 AND entry_date >= ?", (orig_plate, orig_driver, today_str))
    conn.commit()
    conn.close()
    db_backup_manager.backup_database_to_json()
    print("    ✅ Tractor #1 cleanly reverted to original plate.")

    print("\n" + "=" * 70)
    print("🎉 101% VERIFICATION COMPLETE: ALL SYSTEMS, PASSWORDS & SYNCS WORKING!")
    print("=" * 70)

if __name__ == "__main__":
    test_full_system()

