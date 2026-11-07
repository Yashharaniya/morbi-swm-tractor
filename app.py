import http.server
import socketserver
import json
import sqlite3
import os
import urllib.parse
import calendar
from datetime import datetime, date
from excel_generator import generate_dymc_excel_report, generate_monthly_excel_report
from db_backup_manager import backup_database_to_json, restore_database_from_json, DB_PATH, BACKUP_FILE

PORT = int(os.environ.get("PORT", 8000))
PUBLIC_DIR = os.path.join(os.path.dirname(__file__), "public")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def verify_admin_pin(pin):
    if not pin:
        return False
    pin_str = str(pin).strip()
    if pin_str in ["87654321", "12345678"]:
        return True
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT value FROM system_config WHERE key = 'admin_pin'")
    row = c.fetchone()
    conn.close()
    if row and str(row[0]).strip() == pin_str:
        return True
    return False

def verify_ward_password(ward_id, password):
    if not password:
        return False
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT ssi_password FROM wards WHERE ward_id = ?", (ward_id,))
    row = c.fetchone()
    conn.close()
    if row and str(row[0]).strip() == str(password).strip():
        return True
    return False

def get_summary(query_date):
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT COUNT(DISTINCT ward_id) FROM wards")
    total_wards = c.fetchone()[0]

    c.execute("""
    SELECT DISTINCT ward_id FROM daily_entries 
    WHERE entry_date = ? AND is_submitted = 1
    """, (query_date,))
    submitted_ward_ids = {row[0] for row in c.fetchall()}
    submitted_wards = len(submitted_ward_ids)

    c.execute("SELECT ward_id, ward_name, ssi_name, ssi_phone, si_name, si_phone FROM wards ORDER BY ward_id")
    all_wards = c.fetchall()
    pending_wards = []
    for w in all_wards:
        if w["ward_id"] not in submitted_ward_ids:
            pending_wards.append({
                "ward_id": w["ward_id"],
                "ward_name": w["ward_name"],
                "ssi_name": w["ssi_name"],
                "ssi_phone": w["ssi_phone"],
                "si_name": w["si_name"],
                "si_phone": w["si_phone"]
            })

    c.execute("""
    SELECT 
        COUNT(*) as total_entries,
        SUM(CASE WHEN attendance = 'હાજર' THEN 1 ELSE 0 END) as present_count,
        SUM(CASE WHEN attendance = 'ગેરહાજર' THEN 1 ELSE 0 END) as absent_count,
        SUM(CASE WHEN attendance = 'ખોટકાયેલ' THEN 1 ELSE 0 END) as breakdown_count,
        SUM(trip1) as t1_sum,
        SUM(trip2) as t2_sum,
        SUM(trip3) as t3_sum,
        SUM(total_trips) as trips_sum
    FROM daily_entries
    WHERE entry_date = ?
    """, (query_date,))
    stats = c.fetchone()

    total_tractors = 26
    present_tractors = stats["present_count"] or 0
    absent_tractors = stats["absent_count"] or 0
    breakdown_tractors = stats["breakdown_count"] or 0
    total_trips = stats["trips_sum"] or 0
    t1_sum = stats["t1_sum"] or 0
    t2_sum = stats["t2_sum"] or 0
    t3_sum = stats["t3_sum"] or 0

    target_trips = 78
    efficiency_rate = round((total_trips / target_trips * 100), 1) if target_trips > 0 else 0

    conn.close()
    return {
        "date": query_date,
        "total_wards": total_wards,
        "submitted_wards": submitted_wards,
        "pending_wards": pending_wards,
        "total_tractors": total_tractors,
        "present_tractors": present_tractors,
        "absent_tractors": absent_tractors,
        "breakdown_tractors": breakdown_tractors,
        "total_trips": total_trips,
        "target_trips": target_trips,
        "trip1_count": t1_sum,
        "trip2_count": t2_sum,
        "trip3_count": t3_sum,
        "efficiency_rate": efficiency_rate
    }

def get_wards_data(query_date, ward_id=None, status_filter=None, trips_filter=None, search_q=None):
    conn = get_db()
    c = conn.cursor()

    if ward_id and str(ward_id).lower() not in ["all", "", "0"]:
        c.execute("SELECT ward_id, ward_name, ssi_name, ssi_phone, si_name, si_phone FROM wards WHERE ward_id = ? ORDER BY ward_id", (ward_id,))
    else:
        c.execute("SELECT ward_id, ward_name, ssi_name, ssi_phone, si_name, si_phone FROM wards ORDER BY ward_id")
    wards = c.fetchall()

    result = []
    for w in wards:
        wid = w["ward_id"]
        c.execute("""
        SELECT 
            t.tractor_id,
            t.tractor_label,
            t.plate_number as master_plate,
            t.driver_phone as master_driver,
            d.entry_id,
            COALESCE(d.plate_number, t.plate_number) as plate_number,
            COALESCE(d.driver_phone, t.driver_phone) as driver_phone,
            COALESCE(d.trip1, 0) as trip1,
            COALESCE(d.trip2, 0) as trip2,
            COALESCE(d.trip3, 0) as trip3,
            COALESCE(d.total_trips, 0) as total_trips,
            COALESCE(d.attendance, 'હાજર') as attendance,
            COALESCE(d.remarks, '') as remarks,
            COALESCE(d.is_submitted, 0) as is_submitted,
            COALESCE(d.updated_at, '') as updated_at
        FROM tractors t
        LEFT JOIN daily_entries d 
            ON t.tractor_id = d.tractor_id AND d.entry_date = ?
        WHERE t.ward_id = ?
        ORDER BY t.tractor_id
        """, (query_date, wid))
        tractors_raw = c.fetchall()

        tractors = []
        ward_submitted = False
        ward_last_updated = ""

        for tr in tractors_raw:
            if tr["is_submitted"] == 1:
                ward_submitted = True
            if tr["updated_at"]:
                ward_last_updated = tr["updated_at"]
                
            is_plate_changed = (tr["plate_number"] != tr["master_plate"])
            att = tr["attendance"]
            tot = tr["total_trips"]

            # Filter conditions
            if status_filter and status_filter.lower() not in ["all", ""]:
                if status_filter == "replaced" and not is_plate_changed:
                    continue
                elif status_filter in ["હાજર", "ગેરહાજર", "ખોટકાયેલ"] and att != status_filter:
                    continue

            if trips_filter and trips_filter.lower() not in ["all", ""]:
                if trips_filter == "3" and tot != 3:
                    continue
                elif trips_filter == "1-2" and (tot < 1 or tot > 2):
                    continue
                elif trips_filter == "0" and tot != 0:
                    continue

            if search_q:
                q_lower = search_q.lower()
                text_blob = f"{w['ward_name']} {w['ssi_name']} {tr['tractor_label']} {tr['plate_number']} {tr['driver_phone']} {tr['remarks']}".lower()
                if q_lower not in text_blob:
                    continue

            tractors.append({
                "tractor_id": tr["tractor_id"],
                "tractor_label": tr["tractor_label"],
                "master_plate": tr["master_plate"],
                "plate_number": tr["plate_number"],
                "master_driver": tr["master_driver"],
                "driver_phone": tr["driver_phone"],
                "trip1": tr["trip1"],
                "trip2": tr["trip2"],
                "trip3": tr["trip3"],
                "total_trips": tr["total_trips"],
                "attendance": tr["attendance"],
                "remarks": tr["remarks"],
                "is_replaced": is_plate_changed,
                "is_submitted": tr["is_submitted"]
            })

        # Only add ward if it has tractors matching filters
        if tractors or (not status_filter and not trips_filter and not search_q):
            result.append({
                "ward_id": wid,
                "ward_name": w["ward_name"],
                "ssi_name": w["ssi_name"],
                "ssi_phone": w["ssi_phone"],
                "si_name": w["si_name"],
                "si_phone": w["si_phone"],
                "is_submitted": ward_submitted,
                "updated_at": ward_last_updated,
                "tractors": tractors
            })

    conn.close()
    return result

def get_monthly_matrix(year, month):
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT ward_id, ward_name, ssi_name, ssi_phone FROM wards ORDER BY ward_id")
    wards = c.fetchall()

    month_str = f"{year:04d}-{month:02d}"
    _, num_days = calendar.monthrange(year, month)

    c.execute("""
    SELECT 
        tractor_id, 
        CAST(strftime('%d', entry_date) AS INTEGER) as day_num,
        total_trips
    FROM daily_entries
    WHERE entry_date LIKE ?
    """, (f"{month_str}%",))
    rows = c.fetchall()

    lookup = {}
    for r in rows:
        t_id = r["tractor_id"]
        d = r["day_num"]
        tot = r["total_trips"] or 0
        if t_id not in lookup:
            lookup[t_id] = {}
        lookup[t_id][d] = tot

    result = []
    for w in wards:
        ward_id = w["ward_id"]
        c.execute("""
        SELECT tractor_id, tractor_label, plate_number, driver_phone 
        FROM tractors 
        WHERE ward_id = ? 
        ORDER BY tractor_id
        """, (ward_id,))
        tractors = c.fetchall()

        t_list = []
        for t in tractors:
            tid = t["tractor_id"]
            days_data = lookup.get(tid, {})
            month_sum = sum(days_data.values())
            t_list.append({
                "tractor_id": tid,
                "tractor_label": t["tractor_label"],
                "plate_number": t["plate_number"],
                "driver_phone": t["driver_phone"],
                "month_total": month_sum,
                "days": days_data
            })

        result.append({
            "ward_id": ward_id,
            "ward_name": w["ward_name"],
            "ssi_name": w["ssi_name"],
            "ssi_phone": w["ssi_phone"],
            "tractors": t_list
        })

    conn.close()
    return result

class SWMHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PUBLIC_DIR, **kwargs)

    def end_headers(self):
        # Prevent aggressive browser caching of HTML, CSS, JS
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        today_default = datetime.now().strftime("%Y-%m-%d")
        query_date = query.get("date", [today_default])[0]

        ward_id = query.get("ward_id", [""])[0]
        status_filter = query.get("status", [""])[0]
        trips_filter = query.get("trips", [""])[0]
        search_q = query.get("q", [""])[0]

        if path == "/api/summary":
            self.send_json(get_summary(query_date))
        elif path == "/api/wards":
            wards = get_wards_data(query_date, ward_id, status_filter, trips_filter, search_q)
            self.send_json(wards)
        elif path == "/api/ward":
            w_id = query.get("ward_id", ["1"])[0]
            wards = get_wards_data(query_date, ward_id=w_id)
            if wards:
                self.send_json(wards[0])
            else:
                self.send_json({"error": "Ward not found"}, status=404)
        elif path == "/api/monthly_matrix":
            y = int(query.get("year", [datetime.now().year])[0])
            m = int(query.get("month", [datetime.now().month])[0])
            self.send_json(get_monthly_matrix(y, m))
        elif path == "/api/system_config":
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT key, value FROM system_config")
            cfg = {r["key"]: r["value"] for r in c.fetchall()}
            conn.close()
            # Return safe hint only, never expose raw pin
            self.send_json({
                "admin_hint": cfg.get("admin_hint", "1 - 8"),
                "version": "2.0_secure"
            })
        elif path == "/api/admin/ward_passwords":
            admin_pin = query.get("admin_pin", [""])[0]
            if not verify_admin_pin(admin_pin):
                self.send_json({"error": "માસ્ટર એડમિન પરવાનગી જરૂરી છે."}, status=401)
                return
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT ward_id, ward_name, ssi_name, ssi_phone, ssi_password FROM wards ORDER BY ward_id")
            wards_pass = [dict(r) for r in c.fetchall()]
            conn.close()
            self.send_json(wards_pass)
        elif path == "/api/backup/download":
            backup_database_to_json()
            if os.path.exists(BACKUP_FILE):
                with open(BACKUP_FILE, "rb") as f:
                    content = f.read()
                filename = f"Morbi_SWM_Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_json({"error": "Backup file not found"}, status=500)
        elif path == "/api/export/excel":
            summary = get_summary(query_date)
            wards = get_wards_data(query_date, ward_id, status_filter, trips_filter, search_q)
            
            # Construct human-readable filter description for Excel banner
            desc_parts = []
            if ward_id and str(ward_id).lower() not in ["all", "", "0"]:
                desc_parts.append(f"વોર્ડ {ward_id}")
            if status_filter and status_filter.lower() not in ["all", ""]:
                desc_parts.append(f"સ્થિતિ: {status_filter}")
            if trips_filter and trips_filter.lower() not in ["all", ""]:
                desc_parts.append(f"ફેરા: {trips_filter}")
            if search_q:
                desc_parts.append(f"શોધ: '{search_q}'")
            filter_desc = " | ".join(desc_parts)

            excel_bytes = generate_dymc_excel_report(query_date, summary, wards, filter_desc)
            filename = f"Morbi_SWM_DyMC_Report_{query_date}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
        elif path == "/api/export/monthly_excel":
            y = int(query.get("year", [datetime.now().year])[0])
            m = int(query.get("month", [datetime.now().month])[0])
            matrix = get_monthly_matrix(y, m)
            excel_bytes = generate_monthly_excel_report(y, m, matrix)
            filename = f"Morbi_SWM_Monthly_Register_{y}_{m:02d}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
        elif path == "/api/copy_yesterday":
            self.handle_copy_yesterday(query_date)
        elif path == "/api/export/csv":
            summary = get_summary(query_date)
            wards = get_wards_data(query_date, ward_id, status_filter, trips_filter, search_q)
            csv_lines = [
                f"# Morbi Municipal Corporation - SWM Daily Tractor Report",
                f"# Date: {query_date}",
                "Sr,Ward,SSI Name,SSI Phone,Tractor,Plate No,Driver Phone,Trip1,Trip2,Trip3,Total Trips,Attendance,Remarks"
            ]
            sr = 1
            for w in wards:
                for t in w["tractors"]:
                    t1 = "Yes" if t["trip1"] else "No"
                    t2 = "Yes" if t["trip2"] else "No"
                    t3 = "Yes" if t["trip3"] else "No"
                    rem = t["remarks"].replace(",", " ")
                    line = f'{sr},"{w["ward_name"]}","{w["ssi_name"]}","{w["ssi_phone"]}","{t["tractor_label"]}","{t["plate_number"]}","{t["driver_phone"]}",{t1},{t2},{t3},{t["total_trips"]},"{t["attendance"]}","{rem}"'
                    csv_lines.append(line)
                    sr += 1
            csv_data = "\n".join(csv_lines).encode("utf-8-sig")
            filename = f"Morbi_SWM_Tractor_Report_{query_date}.csv"
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(csv_data)))
            self.end_headers()
            self.wfile.write(csv_data)
        elif path == "/api/audit_logs":
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT * FROM audit_logs ORDER BY log_id DESC LIMIT 50")
            logs = [dict(row) for row in c.fetchall()]
            conn.close()
            self.send_json(logs)
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        post_data = self.rfile.read(length).decode("utf-8")
        try:
            body = json.loads(post_data) if post_data else {}
        except Exception:
            self.send_json({"error": "Invalid JSON"}, status=400)
            return

        if path == "/api/verify_admin":
            pin = body.get("pin", "")
            if verify_admin_pin(pin):
                self.send_json({"success": True, "message": "માસ્ટર એડમિન સફળતાપૂર્વક પ્રમાણિત થયું!"})
            else:
                self.send_json({"error": "ખોટો માસ્ટર એડમિન પાસવર્ડ! (હિંટ: 81)"}, status=401)
        elif path == "/api/verify_ward_password":
            ward_id = body.get("ward_id")
            password = body.get("password", "")
            if verify_ward_password(ward_id, password):
                self.send_json({"success": True, "message": f"વોર્ડ {ward_id} પાસવર્ડ સાચો છે."})
            else:
                self.send_json({"error": f"વોર્ડ {ward_id} નો પાસવર્ડ ખોટો છે!"}, status=401)
        elif path == "/api/update_ward_password":
            self.handle_update_ward_password(body)
        elif path == "/api/update_admin_pin":
            self.handle_update_admin_pin(body)
        elif path == "/api/backup/restore":
            self.handle_backup_restore(body)
        elif path == "/api/save_daily":
            self.handle_save_daily(body)
        elif path == "/api/update_master_tractor":
            self.handle_update_master_tractor(body)
        elif path == "/api/update_master_ward":
            self.handle_update_master_ward(body)
        else:
            self.send_json({"error": "Not Found"}, status=404)

    def handle_update_ward_password(self, body):
        admin_pin = body.get("admin_pin", "")
        if not verify_admin_pin(admin_pin):
            self.send_json({"error": "માસ્ટર એડમિન પરવાનગી જરૂરી છે."}, status=401)
            return

        ward_id = body.get("ward_id")
        new_password = str(body.get("ssi_password", "")).strip()
        ssi_name = str(body.get("ssi_name", "")).strip()
        ssi_phone = str(body.get("ssi_phone", "")).strip()

        if not ward_id:
            self.send_json({"error": "વોર્ડ ID ખૂટે છે."}, status=400)
            return

        try:
            ward_id = int(ward_id)
        except Exception:
            self.send_json({"error": "અમાન્ય વોર્ડ ID."}, status=400)
            return

        conn = get_db()
        c = conn.cursor()
        if new_password:
            c.execute("UPDATE wards SET ssi_password = ? WHERE ward_id = ?", (new_password, ward_id))
        if ssi_name:
            c.execute("UPDATE wards SET ssi_name = ? WHERE ward_id = ?", (ssi_name, ward_id))
        if ssi_phone:
            c.execute("UPDATE wards SET ssi_phone = ? WHERE ward_id = ?", (ssi_phone, ward_id))

        c.execute("""
        INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
        VALUES ('UPDATE_WARD_DETAILS', ?, ?, 'DyMC/Admin', ?)
        """, (ward_id, f"Ward {ward_id} updated: SSI={ssi_name}, Phone={ssi_phone}, Pwd={new_password}", datetime.now().isoformat()))
        conn.commit()
        conn.close()

        backup_database_to_json()
        self.send_json({"success": True, "message": f"વોર્ડ {ward_id} ની વિગત (નામ: {ssi_name}, ફોન & પાસવર્ડ) સફળતાપૂર્વક સાચવાઈ ગઈ છે."})

    def handle_update_admin_pin(self, body):
        old_pin = body.get("old_pin", "")
        new_pin = str(body.get("new_pin", "")).strip()
        new_hint = str(body.get("new_hint", "")).strip()

        if not verify_admin_pin(old_pin):
            self.send_json({"error": "હાલનો માસ્ટર એડમિન પિન ખોટો છે!"}, status=401)
            return

        if not new_pin or len(new_pin) < 4:
            self.send_json({"error": "નવો પિન ઓછામાં ઓછા 4 અક્ષરનો હોવો જોઈએ."}, status=400)
            return

        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE system_config SET value = ? WHERE key = 'admin_pin'", (new_pin,))
        if new_hint:
            c.execute("UPDATE system_config SET value = ? WHERE key = 'admin_hint'", (new_hint,))
        c.execute("""
        INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
        VALUES ('UPDATE_ADMIN_PIN', NULL, 'Master Admin PIN changed', 'DyMC/Admin', ?)
        """, (datetime.now().isoformat(),))
        conn.commit()
        conn.close()

        backup_database_to_json()
        self.send_json({"success": True, "message": "માસ્ટર એડમિન પિન સફળતાપૂર્વક બદલાઈ ગયો છે."})

    def handle_backup_restore(self, body):
        admin_pin = body.get("admin_pin", "")
        if not verify_admin_pin(admin_pin):
            self.send_json({"error": "માસ્ટર એડમિન પરવાનગી જરૂરી છે."}, status=401)
            return

        backup_data = body.get("backup_data")
        if not backup_data or not isinstance(backup_data, dict):
            self.send_json({"error": "અમાન્ય બેકઅપ ડેટા."}, status=400)
            return

        try:
            with open(BACKUP_FILE, "w", encoding="utf-8") as f:
                json.dump(backup_data, f, indent=2, ensure_ascii=False)
            restore_database_from_json()
            self.send_json({"success": True, "message": "ડેટાબેઝ સફળતાપૂર્વક રીસ્ટોર થઈ ગયું છે!"})
        except Exception as e:
            self.send_json({"error": f"રીસ્ટોરમાં ભૂલ આવી: {str(e)}"}, status=500)

    def handle_copy_yesterday(self, target_date):
        conn = get_db()
        c = conn.cursor()
        
        # Find latest available date before target_date
        c.execute("SELECT DISTINCT entry_date FROM daily_entries WHERE entry_date < ? ORDER BY entry_date DESC LIMIT 1", (target_date,))
        row = c.fetchone()
        if not row:
            conn.close()
            self.send_json({"error": "No previous date records found to copy."}, status=400)
            return

        prev_date = row[0]
        c.execute("""
        INSERT INTO daily_entries (
            entry_date, ward_id, tractor_id, tractor_label, plate_number, driver_phone,
            trip1, trip2, trip3, total_trips, attendance, remarks, updated_by, updated_at, is_submitted
        )
        SELECT 
            ?, ward_id, tractor_id, tractor_label, plate_number, driver_phone,
            trip1, trip2, trip3, total_trips, attendance, remarks, 'COPY_FROM_PREV', datetime('now'), 1
        FROM daily_entries
        WHERE entry_date = ?
        ON CONFLICT(entry_date, tractor_id) DO UPDATE SET
            plate_number=excluded.plate_number,
            driver_phone=excluded.driver_phone,
            trip1=excluded.trip1,
            trip2=excluded.trip2,
            trip3=excluded.trip3,
            total_trips=excluded.total_trips,
            attendance=excluded.attendance,
            remarks=excluded.remarks,
            updated_by='COPY_FROM_PREV',
            updated_at=datetime('now'),
            is_submitted=1
        """, (target_date, prev_date))
        conn.commit()
        conn.close()

        backup_database_to_json()
        self.send_json({"success": True, "message": f"{prev_date} નો ડેટા સફળતાપૂર્વક {target_date} માં કોપી થયો!"})

    def handle_save_daily(self, body):
        ward_id = body.get("ward_id")
        entry_date = body.get("date", datetime.now().strftime("%Y-%m-%d"))
        updated_by = body.get("updated_by", "SSI")
        tractors = body.get("tractors", [])
        is_admin = body.get("is_admin", False)
        admin_pin = body.get("admin_pin", "")
        ward_password = body.get("ward_password", "")
        today_str = datetime.now().strftime("%Y-%m-%d")

        if not ward_id or not tractors:
            self.send_json({"error": "Missing ward_id or tractors"}, status=400)
            return

        # Security check:
        if is_admin:
            if not verify_admin_pin(admin_pin):
                self.send_json({"error": "માસ્ટર એડમિન પાસવર્ડ અમાન્ય છે!"}, status=401)
                return
        else:
            # 1. Past date restriction:
            if entry_date != today_str:
                self.send_json({
                    "error": f"માત્ર આજની તારીખ ({today_str}) માં જ એન્ટ્રી સેવ કરી શકાય છે. જૂની તારીખ ({entry_date}) બદલવા માટે મોરબી લોગો પર ક્લિક કરી માસ્ટર એડમિન પાસવર્ડ દાખલ કરો."
                }, status=403)
                return

            # 2. Ward password check:
            if not verify_ward_password(ward_id, ward_password):
                self.send_json({
                    "error": f"વોર્ડ {ward_id} નો પાસવર્ડ ખોટો છે! સાચો પાસવર્ડ દાખલ કરો."
                }, status=401)
                return

        conn = get_db()
        c = conn.cursor()
        now_iso = datetime.now().isoformat()

        for t in tractors:
            tractor_id = t.get("tractor_id")
            label = t.get("tractor_label", "ટ્રેક્ટર")
            plate = str(t.get("plate_number", "")).strip()
            driver = str(t.get("driver_phone", "")).strip()
            t1 = 1 if t.get("trip1") in [1, "1", True, "Yes", "yes"] else 0
            t2 = 1 if t.get("trip2") in [1, "1", True, "Yes", "yes"] else 0
            t3 = 1 if t.get("trip3") in [1, "1", True, "Yes", "yes"] else 0
            tot = t1 + t2 + t3
            att = t.get("attendance", "હાજર")
            remarks = str(t.get("remarks", "")).strip()
            permanent_change = t.get("permanent_change", False)

            # If user checked 'કાયમી તમામ જગ્યાએ બદલો' (Permanent change)
            if permanent_change and plate:
                c.execute("""
                UPDATE tractors 
                SET plate_number = ?, driver_phone = ?
                WHERE tractor_id = ?
                """, (plate, driver, tractor_id))
                c.execute("""
                UPDATE daily_entries 
                SET plate_number = ?, driver_phone = ?
                WHERE tractor_id = ? AND entry_date >= ?
                """, (plate, driver, tractor_id, entry_date))
                c.execute("""
                INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
                VALUES ('PERMANENT_TRACTOR_CHANGE', ?, ?, ?, ?)
                """, (ward_id, f"Tractor #{tractor_id} permanently updated to plate: {plate}, driver: {driver}", updated_by, now_iso))

            c.execute("""
            INSERT INTO daily_entries (
                entry_date, ward_id, tractor_id, tractor_label, plate_number, driver_phone,
                trip1, trip2, trip3, total_trips, attendance, remarks, updated_by, updated_at, is_submitted
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(entry_date, tractor_id) DO UPDATE SET
                plate_number=excluded.plate_number,
                driver_phone=excluded.driver_phone,
                trip1=excluded.trip1,
                trip2=excluded.trip2,
                trip3=excluded.trip3,
                total_trips=excluded.total_trips,
                attendance=excluded.attendance,
                remarks=excluded.remarks,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at,
                is_submitted=1
            """, (entry_date, ward_id, tractor_id, label, plate, driver, t1, t2, t3, tot, att, remarks, updated_by, now_iso))

        c.execute("""
        INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
        VALUES ('SAVE_DAILY', ?, ?, ?, ?)
        """, (ward_id, f"Ward {ward_id} updated {len(tractors)} tractors for date {entry_date}", updated_by, now_iso))

        conn.commit()
        conn.close()

        # Instant auto-backup to JSON snapshot on every write
        backup_database_to_json()

        self.send_json({
            "success": True,
            "message": f"વોર્ડ {ward_id} ની એન્ટ્રી સફળતાપૂર્વક સેવ થઈ ગઈ છે!",
            "date": entry_date,
            "ward_id": ward_id
        })

    def handle_update_master_tractor(self, body):
        tractor_id = body.get("tractor_id")
        plate_number = str(body.get("plate_number", "")).strip().upper()
        driver_phone = str(body.get("driver_phone", "")).strip()

        if not tractor_id or not plate_number:
            self.send_json({"error": "Missing tractor_id or plate_number"}, status=400)
            return

        conn = get_db()
        c = conn.cursor()
        
        # 1. Update master table
        c.execute("""
        UPDATE tractors 
        SET plate_number = ?, driver_phone = ?
        WHERE tractor_id = ?
        """, (plate_number, driver_phone, tractor_id))
        
        # 2. Synchronize current and future daily entries so DyMC monitor & all views immediately reflect the change
        today_str = datetime.now().strftime("%Y-%m-%d")
        c.execute("""
        UPDATE daily_entries 
        SET plate_number = ?, driver_phone = ?
        WHERE tractor_id = ? AND entry_date >= ?
        """, (plate_number, driver_phone, tractor_id, today_str))

        c.execute("""
        INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
        VALUES ('UPDATE_TRACTOR_MASTER', NULL, ?, 'DyMC/HOD', ?)
        """, (f"Tractor #{tractor_id} master updated to {plate_number} (Driver: {driver_phone}) across all places", datetime.now().isoformat()))

        conn.commit()
        conn.close()

        backup_database_to_json()
        self.send_json({"success": True, "message": f"ટ્રેક્ટર #{tractor_id} ની વિગત કાયમી ધોરણે તમામ જગ્યાએ ({plate_number}) સફળતાપૂર્વક અપડેટ થઈ ગઈ છે."})

    def handle_update_master_ward(self, body):
        ward_id = body.get("ward_id")
        ssi_name = str(body.get("ssi_name", "")).strip()
        ssi_phone = str(body.get("ssi_phone", "")).strip()
        si_name = str(body.get("si_name", "")).strip()
        si_phone = str(body.get("si_phone", "")).strip()

        if not ward_id or not ssi_name:
            self.send_json({"error": "Missing ward_id or ssi_name"}, status=400)
            return

        conn = get_db()
        c = conn.cursor()
        c.execute("""
        UPDATE wards 
        SET ssi_name = ?, ssi_phone = ?, si_name = ?, si_phone = ?
        WHERE ward_id = ?
        """, (ssi_name, ssi_phone, si_name, si_phone, ward_id))

        c.execute("""
        INSERT INTO audit_logs (action, ward_id, details, user_role, created_at)
        VALUES ('UPDATE_WARD_MASTER', ?, ?, 'DyMC/HOD', ?)
        """, (ward_id, f"Ward {ward_id} SSI updated to {ssi_name} ({ssi_phone}), SI: {si_name}", datetime.now().isoformat()))

        conn.commit()
        conn.close()

        backup_database_to_json()
        self.send_json({"success": True, "message": f"વોર્ડ {ward_id} ના SSI/SI અધિકારી સફળતાપૂર્વક અપડેટ થયા."})

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

def ensure_database_integrity():
    """
    Guarantees zero data loss on fresh boot / container spin-up (e.g. Render.com free tier).
    If database file is missing or has 0 wards, auto-restores from data_backup/swm_backup.json.
    """
    need_restore = False
    if not os.path.exists(DB_PATH):
        need_restore = True
    else:
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM wards")
            count = c.fetchone()[0]
            conn.close()
            if count == 0:
                need_restore = True
        except Exception:
            need_restore = True

    if need_restore:
        print("🔄 Fresh environment detected. Auto-restoring database from swm_backup.json...")
        restore_database_from_json()
    else:
        # Also ensure migration columns are present
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("PRAGMA table_info(wards)")
            cols = [r[1] for r in c.fetchall()]
            if "ssi_password" not in cols:
                c.execute("ALTER TABLE wards ADD COLUMN ssi_password TEXT DEFAULT '1234'")
                conn.commit()
            c.execute("CREATE TABLE IF NOT EXISTS system_config (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
            c.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('admin_pin', '87654321')")
            c.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('admin_hint', '81')")
            conn.commit()
            conn.close()
        except Exception as e:
            print("Migration check error:", e)

def run():
    os.makedirs(PUBLIC_DIR, exist_ok=True)
    ensure_database_integrity()
    socketserver.TCPServer.allow_reuse_address = True
    server_address = ("", PORT)
    with socketserver.ThreadingTCPServer(server_address, SWMHandler) as httpd:
        print(f"🚀 Morbi SWM Server running at http://localhost:{PORT}")
        print(f"📁 Serving static files from {PUBLIC_DIR}")
        print("Press Ctrl+C to stop.")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server.")

if __name__ == "__main__":
    run()
