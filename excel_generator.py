import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import calendar
from datetime import datetime

# Common Styling Palettes
navy_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Navy Blue
sub_header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid") # Royal Blue
kpi_bg_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid") # Light Blue Tint
col_header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # Slate Dark
alt_row_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Slate Light
highlight_trip_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Light Green (Yes)
no_trip_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Slate (No)
absent_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Light Red
total_row_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Light Amber

font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
font_sub = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_col_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_kpi_val = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
font_data = Font(name="Calibri", size=10, bold=False, color="000000")
font_bold = Font(name="Calibri", size=10, bold=True, color="000000")

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")

thin_border_side = Side(border_style="thin", color="CBD5E1")
border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
thick_bottom_side = Side(border_style="medium", color="1E293B")
border_total = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

GUJARATI_MONTHS = {
    1: "જાન્યુઆરી", 2: "ફેબ્રુઆરી", 3: "માર્ચ", 4: "એપ્રિલ",
    5: "મે", 6: "જૂન", 7: "જુલાઈ", 8: "ઓગસ્ટ",
    9: "સપ્ટેમ્બર", 10: "ઓક્ટોબર", 11: "નવેમ્બર", 12: "ડિસેમ્બર"
}

def generate_dymc_excel_report(date_str, summary, wards_data, filter_desc=""):
    """
    Generates a beautifully styled, executive-ready Excel workbook for DyMC / HOD.
    Supports Filter-wise Export (by Ward, Status, Trip counts, etc.)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"DyMC_{date_str}"
    ws.views.sheetView[0].showGridLines = True

    # 1. Main Title Header (Row 1)
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = "🏛️ મોરબી મહાનગરપાલિકા - સોલિડ વેસ્ટ મેનેજમેન્ટ (SWM) - દૈનિક ટ્રેક્ટર ફેરા રિપોર્ટ"
    title_cell.font = font_title
    title_cell.fill = navy_header_fill
    title_cell.alignment = align_center
    ws.row_dimensions[1].height = 36

    # 2. Subtitle / Metadata (Row 2)
    ws.merge_cells("A2:K2")
    sub_cell = ws["A2"]
    filter_label = f" | ફિલ્ટર: {filter_desc}" if filter_desc else ""
    sub_cell.value = f"📅 રિપોર્ટ તારીખ: {date_str} | DyMC / HOD મોનિટરિંગ એક્સપોર્ટ{filter_label}"
    sub_cell.font = font_sub
    sub_cell.fill = sub_header_fill
    sub_cell.alignment = align_center
    ws.row_dimensions[2].height = 24

    # Calculate actual counts from filtered wards_data
    filtered_tractors_count = sum(len(w.get("tractors", [])) for w in wards_data)
    filtered_present_count = 0
    filtered_trips_sum = 0
    for w in wards_data:
        for t in w.get("tractors", []):
            if t.get("attendance") == "હાજર":
                filtered_present_count += 1
            filtered_trips_sum += t.get("total_trips", 0)

    # 3. KPI Summary Bar (Row 4)
    ws.row_dimensions[4].height = 28
    eff_rate = round((filtered_trips_sum / (filtered_tractors_count * 3) * 100), 1) if filtered_tractors_count > 0 else 0
    kpis = [
        ("A4", "B4", "પસંદ કરેલ વોર્ડ", f"{len(wards_data)} વોર્ડ"),
        ("C4", "D4", "કુલ ટ્રેક્ટર", f"{filtered_tractors_count} ટ્રેક્ટર"),
        ("E4", "F4", "હાજર ટ્રેક્ટર", f"{filtered_present_count} / {filtered_tractors_count}"),
        ("G4", "H4", "કુલ ફેરા", f"{filtered_trips_sum} ફેરા"),
        ("I4", "J4", "કાર્યક્ષમતા %", f"{eff_rate}%"),
        ("K4", "K4", "બાકી વોર્ડ", f"{len(summary.get('pending_wards', []))}")
    ]
    for start_c, end_c, label, val in kpis:
        if start_c != end_c:
            ws.merge_cells(f"{start_c}:{end_c}")
        cell = ws[start_c]
        cell.value = f"{label}: {val}"
        cell.font = font_kpi_val
        cell.fill = kpi_bg_fill
        cell.alignment = align_center
        
        col_start = openpyxl.utils.column_index_from_string(start_c[0])
        col_end = openpyxl.utils.column_index_from_string(end_c[0]) if len(end_c) > 1 else col_start
        for c in range(col_start, col_end + 1):
            ws.cell(4, c).border = border_cell

    # 4. Table Header (Row 6)
    headers = [
        ("અ.નં.", 6),
        ("વોર્ડ નં.", 12),
        ("એસ.એસ.આઈ. (SSI) નામ & મો.", 24),
        ("ટ્રેક્ટર", 12),
        ("ગાડી નંબર", 18),
        ("ડ્રાઈવર / માલિક મો.", 16),
        ("ફેરો ૧\n(Yes/No)", 11),
        ("ફેરો ૨\n(Yes/No)", 11),
        ("ફેરો ૩\n(Yes/No)", 11),
        ("કુલ ફેરા\n(Total)", 12),
        ("હાજરી સ્થિતિ & નોંધ", 32)
    ]
    ws.row_dimensions[6].height = 30
    for col_idx, (h_name, width) in enumerate(headers, start=1):
        c = ws.cell(row=6, column=col_idx, value=h_name)
        c.font = font_col_header
        c.fill = col_header_fill
        c.alignment = align_center
        c.border = border_cell
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    # 5. Populate Data Rows
    current_row = 7
    sr_no = 1
    total_trips_sum = 0
    total_t1 = 0
    total_t2 = 0
    total_t3 = 0

    for ward in wards_data:
        ward_start_row = current_row
        tractors = ward.get("tractors", [])
        
        for t_idx, t in enumerate(tractors):
            ws.row_dimensions[current_row].height = 24
            is_alt = (sr_no % 2 == 0)
            row_fill = alt_row_fill if is_alt else PatternFill(fill_type=None)

            t1_val = "Yes" if t.get("trip1") else "No"
            t2_val = "Yes" if t.get("trip2") else "No"
            t3_val = "Yes" if t.get("trip3") else "No"
            trips_tot = t.get("total_trips", 0)

            total_trips_sum += trips_tot
            if t.get("trip1"): total_t1 += 1
            if t.get("trip2"): total_t2 += 1
            if t.get("trip3"): total_t3 += 1

            att = t.get("attendance", "હાજર")
            remarks = t.get("remarks", "")
            att_remarks_text = f"[{att}]" + (f" {remarks}" if remarks else "")

            ws.cell(row=current_row, column=1, value=sr_no)
            ws.cell(row=current_row, column=2, value=ward.get("ward_name", ""))
            ws.cell(row=current_row, column=3, value=f"{ward.get('ssi_name', '')}\n📞 {ward.get('ssi_phone', '')}")
            ws.cell(row=current_row, column=4, value=t.get("tractor_label", ""))
            ws.cell(row=current_row, column=5, value=t.get("plate_number", ""))
            ws.cell(row=current_row, column=6, value=str(t.get("driver_phone", "")))
            
            c_t1 = ws.cell(row=current_row, column=7, value=t1_val)
            c_t2 = ws.cell(row=current_row, column=8, value=t2_val)
            c_t3 = ws.cell(row=current_row, column=9, value=t3_val)
            c_tot = ws.cell(row=current_row, column=10, value=trips_tot)
            c_att = ws.cell(row=current_row, column=11, value=att_remarks_text)

            for c_idx in range(1, 12):
                cell = ws.cell(row=current_row, column=c_idx)
                cell.font = font_data
                cell.border = border_cell
                if row_fill.fill_type:
                    cell.fill = row_fill

            ws.cell(row=current_row, column=1).alignment = align_center
            ws.cell(row=current_row, column=2).alignment = align_center
            ws.cell(row=current_row, column=3).alignment = align_center
            ws.cell(row=current_row, column=4).alignment = align_center
            ws.cell(row=current_row, column=5).alignment = align_center
            ws.cell(row=current_row, column=6).alignment = align_center
            c_t1.alignment = align_center
            c_t2.alignment = align_center
            c_t3.alignment = align_center
            c_tot.alignment = align_center
            c_att.alignment = align_left

            c_t1.fill = highlight_trip_fill if t1_val == "Yes" else no_trip_fill
            c_t2.fill = highlight_trip_fill if t2_val == "Yes" else no_trip_fill
            c_t3.fill = highlight_trip_fill if t3_val == "Yes" else no_trip_fill
            if att == "ગેરહાજર": c_att.fill = absent_fill

            c_tot.font = font_bold
            sr_no += 1
            current_row += 1

        if len(tractors) > 1:
            ws.merge_cells(start_row=ward_start_row, start_column=2, end_row=ward_start_row + len(tractors) - 1, end_column=2)
            ws.merge_cells(start_row=ward_start_row, start_column=3, end_row=ward_start_row + len(tractors) - 1, end_column=3)

    # 6. Total Summary Row
    ws.row_dimensions[current_row].height = 28
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
    tot_label = ws.cell(row=current_row, column=1, value="દૈનિક કુલ સરવાળો (Daily Total Sum):")
    tot_label.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    tot_label.alignment = Alignment(horizontal="right", vertical="center")
    
    tot_t1 = ws.cell(row=current_row, column=7, value=total_t1)
    tot_t2 = ws.cell(row=current_row, column=8, value=total_t2)
    tot_t3 = ws.cell(row=current_row, column=9, value=total_t3)
    tot_all = ws.cell(row=current_row, column=10, value=total_trips_sum)
    tot_status = ws.cell(row=current_row, column=11, value=f"કુલ {filtered_tractors_count} ટ્રેક્ટર પૈકી {filtered_present_count} હાજર")

    for c_idx in range(1, 12):
        c = ws.cell(row=current_row, column=c_idx)
        c.fill = total_row_fill
        c.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        c.border = border_total
        if c_idx >= 7 and c_idx <= 10:
            c.alignment = align_center
        elif c_idx == 11:
            c.alignment = align_left

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()

def generate_monthly_excel_report(year, month, monthly_matrix):
    """
    Generates the Full Monthly Register Excel matching the original 12 monthly tabs!
    Columns: Sr, Ward, SSI, Tractor, Plate, Driver, Monthly Total, Avg/Day, Status, Day 1 to Day 30/31.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    month_name_gu = GUJARATI_MONTHS.get(month, f"{month}")
    ws.title = f"{month}-{month_name_gu}"
    ws.views.sheetView[0].showGridLines = True

    _, num_days = calendar.monthrange(year, month)

    total_cols = 9 + num_days
    last_col_letter = openpyxl.utils.get_column_letter(total_cols)
    ws.merge_cells(f"A1:{last_col_letter}1")
    t1 = ws["A1"]
    t1.value = f"🏛️ મોરબી મહાનગરપાલિકા - {month_name_gu} {year} ટ્રેક્ટર ફેરા રજિસ્ટર ({num_days} દિવસ રજિસ્ટર)"
    t1.font = font_title
    t1.fill = navy_header_fill
    t1.alignment = align_center
    ws.row_dimensions[1].height = 36

    ws.merge_cells(f"A2:{last_col_letter}2")
    t2 = ws["A2"]
    t2.value = f"{month_name_gu} {year} | વોર્ડ ૧ થી ૧૩ (૨૬ ટ્રેક્ટર) | દરેક ટ્રેક્ટર ૩ ફેરા ક્ષમતા | DyMC & HOD માસિક મોનિટરિંગ"
    t2.font = font_sub
    t2.fill = sub_header_fill
    t2.alignment = align_center
    ws.row_dimensions[2].height = 24

    headers = [
        ("અ.નં.", 6),
        ("વોર્ડ નં.", 12),
        ("એસ.એસ.આઈ. (SSI) નામ & મો.", 24),
        ("ટ્રેક્ટર", 12),
        ("ગાડી નંબર", 18),
        ("ડ્રાઈવર / માલિક મો.", 16),
        ("માસિક કુલ ફેરા", 14),
        ("દૈનિક સરેરાશ", 12),
        ("સ્થિતિ", 12)
    ]
    for d in range(1, num_days + 1):
        headers.append((f"{d}-{month_name_gu[:3]}", 9))

    ws.row_dimensions[4].height = 28
    for col_idx, (h_name, width) in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=h_name)
        c.font = font_col_header
        c.fill = col_header_fill
        c.alignment = align_center
        c.border = border_cell
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    current_row = 5
    sr_no = 1
    daily_sums = [0] * num_days
    grand_month_total = 0

    for ward in monthly_matrix:
        ward_start_row = current_row
        tractors = ward.get("tractors", [])
        for t in tractors:
            ws.row_dimensions[current_row].height = 22
            is_alt = (sr_no % 2 == 0)
            row_fill = alt_row_fill if is_alt else PatternFill(fill_type=None)

            month_trips = t.get("month_total", 0)
            grand_month_total += month_trips
            avg_per_day = round(month_trips / num_days, 1)

            ws.cell(row=current_row, column=1, value=sr_no)
            ws.cell(row=current_row, column=2, value=ward.get("ward_name", ""))
            ws.cell(row=current_row, column=3, value=f"{ward.get('ssi_name', '')}\n📞 {ward.get('ssi_phone', '')}")
            ws.cell(row=current_row, column=4, value=t.get("tractor_label", ""))
            ws.cell(row=current_row, column=5, value=t.get("plate_number", ""))
            ws.cell(row=current_row, column=6, value=str(t.get("driver_phone", "")))
            
            c_m_tot = ws.cell(row=current_row, column=7, value=month_trips)
            c_m_avg = ws.cell(row=current_row, column=8, value=avg_per_day)
            c_m_st = ws.cell(row=current_row, column=9, value="સક્રિય")

            c_m_tot.font = font_bold
            c_m_tot.fill = kpi_bg_fill

            day_trips = t.get("days", {})
            for d in range(1, num_days + 1):
                col_i = 9 + d
                val = day_trips.get(d, 0)
                daily_sums[d - 1] += val
                cell = ws.cell(row=current_row, column=col_i, value=val if val > 0 else "")
                cell.alignment = align_center
                cell.border = border_cell
                if val > 0:
                    cell.fill = highlight_trip_fill
                    cell.font = font_bold
                else:
                    if row_fill.fill_type: cell.fill = row_fill

            for c_idx in range(1, 10):
                cell = ws.cell(row=current_row, column=c_idx)
                cell.border = border_cell
                if c_idx != 7 and row_fill.fill_type:
                    cell.fill = row_fill
                if c_idx in [1, 2, 4, 5, 6, 7, 8, 9]:
                    cell.alignment = align_center
                elif c_idx == 3:
                    cell.alignment = align_left

            sr_no += 1
            current_row += 1

        if len(tractors) > 1:
            ws.merge_cells(start_row=ward_start_row, start_column=2, end_row=ward_start_row + len(tractors) - 1, end_column=2)
            ws.merge_cells(start_row=ward_start_row, start_column=3, end_row=ward_start_row + len(tractors) - 1, end_column=3)

    ws.row_dimensions[current_row].height = 28
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
    tot_label = ws.cell(row=current_row, column=1, value="માસિક તથા દૈનિક કુલ ફેરા સરવાળો:")
    tot_label.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    tot_label.alignment = Alignment(horizontal="right", vertical="center")
    
    tot_month = ws.cell(row=current_row, column=7, value=grand_month_total)
    tot_month.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    tot_month.alignment = align_center

    ws.cell(row=current_row, column=8, value=round(grand_month_total / num_days, 1)).alignment = align_center
    ws.cell(row=current_row, column=9, value=f"{sr_no - 1} ટ્રેક્ટર").alignment = align_center

    for d in range(1, num_days + 1):
        col_i = 9 + d
        c = ws.cell(row=current_row, column=col_i, value=daily_sums[d - 1])
        c.font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
        c.alignment = align_center
        c.fill = total_row_fill
        c.border = border_total

    for c_idx in range(1, 10):
        c = ws.cell(row=current_row, column=c_idx)
        c.fill = total_row_fill
        c.border = border_total

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()
