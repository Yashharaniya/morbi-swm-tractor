import sqlite3
import openpyxl
from datetime import datetime

DB_PATH = r"D:\yash\Website\morbi_swm.db"
EXCEL_PATH = r"D:\yash\Website\Tractor_SSI_Annual_Dashboard.xlsx"

# 13 Wards Master Information (SSI & SI names and phone numbers)
WARDS_INFO = [
    {"ward_id": 1, "ward_name": "વોર્ડ - 1", "ssi_name": "જીગ્નેશભાઈ ચૌહાણ", "ssi_phone": "9974599703", "si_name": "ભૂમિબા વાળા", "si_phone": "9974566346"},
    {"ward_id": 2, "ward_name": "વોર્ડ - 2", "ssi_name": "અરુણભાઈ માથુર", "ssi_phone": "9925182979", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 3, "ward_name": "વોર્ડ - 3", "ssi_name": "સુરેશ ટોયટા", "ssi_phone": "7069151493", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 4, "ward_name": "વોર્ડ - 4", "ssi_name": "બ્રિજરાજસિંહ વાળા", "ssi_phone": "9998441077", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 5, "ward_name": "વોર્ડ - 5", "ssi_name": "ધ્રુવ રામાવત", "ssi_phone": "8238235007", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 6, "ward_name": "વોર્ડ - 6", "ssi_name": "પૃથ્વીબેન તલસાણીયા", "ssi_phone": "9313902658", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 7, "ward_name": "વોર્ડ - 7", "ssi_name": "ધવલભાઈ પોપટણી", "ssi_phone": "9898262933", "si_name": "સંજયભાઇ", "si_phone": "7600162383"},
    {"ward_id": 8, "ward_name": "વોર્ડ - 8", "ssi_name": "રુચિતા", "ssi_phone": "9624776125", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 9, "ward_name": "વોર્ડ - 9", "ssi_name": "જાનવી વાઢેર", "ssi_phone": "8140305133", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 10, "ward_name": "વોર્ડ - 10", "ssi_name": "મયુરભાઈ ઘોરડીયા", "ssi_phone": "6354992588", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 11, "ward_name": "વોર્ડ - 11", "ssi_name": "મીનાબેન વાઘેલા", "ssi_phone": "9327639393", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 12, "ward_name": "વોર્ડ - 12", "ssi_name": "કેતનભાઈ પટેલ", "ssi_phone": "8141992028", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""},
    {"ward_id": 13, "ward_name": "વોર્ડ - 13", "ssi_name": "શૈલેષભાઈ રંગપરા", "ssi_phone": "9904724280", "si_name": "સેનિટરી ઇન્સ્પેક્ટર", "si_phone": ""}
]

# 26 Tractors Master Information (directly from Excel Daily sheet)
TRACTORS_INFO = [
    # Ward 1
    {"ward_id": 1, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "RJ-03-0076", "driver_phone": "6355573349"},
    {"ward_id": 1, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "GJ-36-AR-1273", "driver_phone": "8511397286"},
    # Ward 2
    {"ward_id": 2, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "GJ-20-CA-4064", "driver_phone": "9328970617"},
    {"ward_id": 2, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-ZK-1908", "driver_phone": "9302149430"},
    # Ward 3
    {"ward_id": 3, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-ZH-7342", "driver_phone": "9752536104"},
    {"ward_id": 3, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-AU-5620", "driver_phone": "82004 17723"},
    # Ward 4
    {"ward_id": 4, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "RJ-03-RB-9846", "driver_phone": "8209282113"},
    {"ward_id": 4, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-ZE-4594", "driver_phone": "9685453542"},
    # Ward 5
    {"ward_id": 5, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-ZG-2402", "driver_phone": "7874196096"},
    {"ward_id": 5, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-AA-9173", "driver_phone": "7898813070"},
    # Ward 6
    {"ward_id": 6, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-ZJ-7714", "driver_phone": "9179092942"},
    {"ward_id": 6, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-ZE-7877", "driver_phone": "7016743033"},
    # Ward 7
    {"ward_id": 7, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-ZE-7529", "driver_phone": "7489861932"},
    {"ward_id": 7, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-AA-8190", "driver_phone": "9328905817"},
    # Ward 8
    {"ward_id": 8, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-32-ZB-7412", "driver_phone": "8780726232"},
    {"ward_id": 8, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-11-ZM-1026", "driver_phone": "7470516280"},
    # Ward 9
    {"ward_id": 9, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-ZG-4443", "driver_phone": "9510149067"},
    {"ward_id": 9, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-45-ZU-7277", "driver_phone": "6355499069"},
    # Ward 10
    {"ward_id": 10, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-AA-7242", "driver_phone": "7247648244"},
    {"ward_id": 10, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "નવું મહેન્દ્રા(નંબર નથી)", "driver_phone": "7089335465"},
    # Ward 11
    {"ward_id": 11, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-13-ZN-1352", "driver_phone": "9510141005"},
    {"ward_id": 11, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "નવું મહેન્દ્રા(નંબર નથી)", "driver_phone": "9630987201"},
    # Ward 12
    {"ward_id": 12, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-AC-2855", "driver_phone": "9238673410"},
    {"ward_id": 12, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-20-AB-9978", "driver_phone": "9274306405"},
    # Ward 13
    {"ward_id": 13, "tractor_label": "ટ્રેક્ટર ૧", "plate_number": "MP-45-AA-5418", "driver_phone": "8359083551"},
    {"ward_id": 13, "tractor_label": "ટ્રેક્ટર ૨", "plate_number": "MP-43-ZE-9622", "driver_phone": "8815354464"}
]

def seed():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("DROP TABLE IF EXISTS audit_logs")
    cursor.execute("DROP TABLE IF EXISTS daily_entries")
    cursor.execute("DROP TABLE IF EXISTS tractors")
    cursor.execute("DROP TABLE IF EXISTS wards")
    cursor.execute("DROP TABLE IF EXISTS users")

    cursor.execute("""
    CREATE TABLE wards (
        ward_id INTEGER PRIMARY KEY,
        ward_name TEXT NOT NULL,
        ssi_name TEXT NOT NULL,
        ssi_phone TEXT NOT NULL,
        si_name TEXT DEFAULT '',
        si_phone TEXT DEFAULT ''
    )
    """)

    cursor.execute("""
    CREATE TABLE tractors (
        tractor_id INTEGER PRIMARY KEY AUTOINCREMENT,
        ward_id INTEGER NOT NULL,
        tractor_label TEXT NOT NULL,
        plate_number TEXT NOT NULL,
        driver_phone TEXT DEFAULT '',
        status TEXT DEFAULT 'સક્રિય',
        FOREIGN KEY (ward_id) REFERENCES wards (ward_id)
    )
    """)

    cursor.execute("""
    CREATE TABLE daily_entries (
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

    cursor.execute("""
    CREATE TABLE audit_logs (
        log_id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT NOT NULL,
        ward_id INTEGER,
        details TEXT NOT NULL,
        user_role TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    cursor.execute("""
    CREATE TABLE users (
        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_pin TEXT NOT NULL,
        role TEXT NOT NULL,
        ward_id INTEGER,
        display_name TEXT NOT NULL
    )
    """)

    # Populate Wards
    for w in WARDS_INFO:
        cursor.execute("""
        INSERT INTO wards (ward_id, ward_name, ssi_name, ssi_phone, si_name, si_phone)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (w["ward_id"], w["ward_name"], w["ssi_name"], w["ssi_phone"], w["si_name"], w["si_phone"]))

    # Populate Tractors
    for t in TRACTORS_INFO:
        cursor.execute("""
        INSERT INTO tractors (ward_id, tractor_label, plate_number, driver_phone)
        VALUES (?, ?, ?, ?)
        """, (t["ward_id"], t["tractor_label"], t["plate_number"], t["driver_phone"]))

    # Populate Users (DyMC, HOD, SI, SSI)
    users = [
        ("dymc", "1234", "DYMC", None, "ડેપ્યુટી મ્યુનિસિપલ કમિશ્નર (DyMC)"),
        ("hod", "1234", "HOD", None, "HOD મોનિટર (સોલિડ વેસ્ટ)"),
        ("si_bhumiba", "1234", "SI", 1, "SI ભૂમિબા વાળા"),
        ("si_sanjay", "1234", "SI", 7, "SI સંજયભાઇ")
    ]
    for ward_id in range(1, 14):
        ward_info = WARDS_INFO[ward_id - 1]
        users.append((f"ward{ward_id}", "1234", "SSI", ward_id, f"SSI {ward_info['ssi_name']} ({ward_info['ward_name']})"))

    for u in users:
        cursor.execute("""
        INSERT INTO users (username, password_pin, role, ward_id, display_name)
        VALUES (?, ?, ?, ?, ?)
        """, u)

    # Import initial data directly from Excel sheet 'દૈનિક એન્ટ્રી (Daily)'
    today_str = datetime.now().strftime("%Y-%m-%d")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    daily_ws = wb['દૈનિક એન્ટ્રી (Daily)']

    for idx, r in enumerate(range(5, 31)):
        tractor_id = idx + 1
        ward_id = ((tractor_id - 1) // 2) + 1
        tractor_label = "ટ્રેક્ટર ૧" if tractor_id % 2 != 0 else "ટ્રેક્ટર ૨"

        plate = str(daily_ws.cell(r, 5).value or "").strip() or TRACTORS_INFO[idx]["plate_number"]
        driver = daily_ws.cell(r, 6).value
        driver_str = str(int(driver)) if isinstance(driver, float) else str(driver or TRACTORS_INFO[idx]["driver_phone"]).strip()

        t1_val = str(daily_ws.cell(r, 7).value or "").strip().lower()
        t2_val = str(daily_ws.cell(r, 8).value or "").strip().lower()
        t3_val = str(daily_ws.cell(r, 9).value or "").strip().lower()

        t1 = 1 if t1_val in ["yes", "હા", "1", "true"] else 0
        t2 = 1 if t2_val in ["yes", "હા", "1", "true"] else 0
        t3 = 1 if t3_val in ["yes", "હા", "1", "true"] else 0
        tot = t1 + t2 + t3

        status = str(daily_ws.cell(r, 11).value or "હાજર").strip()
        remarks = str(daily_ws.cell(r, 12).value or "").strip()

        cursor.execute("""
        INSERT INTO daily_entries (
            entry_date, ward_id, tractor_id, tractor_label, plate_number, driver_phone,
            trip1, trip2, trip3, total_trips, attendance, remarks, updated_by, updated_at, is_submitted
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'EXCEL_SYNC', ?, 1)
        """, (
            today_str, ward_id, tractor_id, tractor_label, plate, driver_str,
            t1, t2, t3, tot, status, remarks, datetime.now().isoformat()
        ))

    cursor.execute("""
    INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
    VALUES ('EXCEL_SEED', NULL, 'Database synchronized directly from Tractor_SSI_Annual_Dashboard.xlsx', 'SYSTEM', ?)
    """, (datetime.now().isoformat(),))

    conn.commit()
    conn.close()
    print("Database reseeded successfully from Excel.")

if __name__ == "__main__":
    seed()
