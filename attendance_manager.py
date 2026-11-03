import re
import io
from datetime import datetime, date, time, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Regular expression to parse raw facial recognition / biometric punch lines
# Example line from machine:
# 242    2026-09-09 18:57:13    4    1    Md Ershad Prodhan Rahat I    0    1
PUNCH_REGEX = re.compile(
    r'^\s*(?P<emp_id>[A-Za-z0-9_-]+)\s+(?P<datetime>\d{4}[-/]\d{2}[-/]\d{2}\s+\d{2}:\d{2}(?::\d{2})?)(?P<rest>.*)$'
)

def parse_notepad_punch_lines(content):
    """
    Parses raw Notepad file content or pasted text line by line.
    Returns a list of parsed punch dicts:
    [{ 'device_user_id': '242', 'punch_time': datetime(...), 'device_code': '4 1', 'raw_line': ... }]
    """
    if not content:
        return []
    
    lines = content.splitlines() if isinstance(content, str) else content
    parsed = []
    
    for line in lines:
        if not line or not line.strip():
            continue
        line_str = line.strip()
        m = PUNCH_REGEX.match(line_str)
        if m:
            raw_emp_id = m.group('emp_id').strip()
            dt_str = m.group('datetime').strip().replace('/', '-')
            rest_part = m.group('rest').strip()
            
            try:
                # Handle YYYY-MM-DD HH:MM:SS or YYYY-MM-DD HH:MM
                if len(dt_str) == 16:
                    p_time = datetime.strptime(dt_str, '%Y-%m-%d %H:%M')
                else:
                    p_time = datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
            except Exception:
                continue

            # Extract device code or mode if present (e.g. "4 1" or "109 15")
            rest_tokens = rest_part.split()
            device_code = " ".join(rest_tokens[:2]) if len(rest_tokens) >= 2 else (rest_tokens[0] if rest_tokens else "-")
            
            parsed.append({
                'device_user_id': raw_emp_id,
                'punch_time': p_time,
                'device_code': device_code,
                'raw_line': line_str
            })
            
    return parsed


def analyze_employee_punches(punches, standard_hours=8.0, grace_period_mins=15):
    """
    AI Shift and Duty Analyzer Engine.
    Intelligently clusters biometric punches, pairs Check-In and Check-Out,
    classifies Day vs Night vs Double Shifts, calculates working hours,
    overtime, and generates human-readable AI explanation notes.
    """
    if not punches:
        return []

    # Sort strictly chronologically
    sorted_punches = sorted(punches, key=lambda x: x['punch_time'] if isinstance(x, dict) else x.punch_time)
    
    # Step 1: De-duplicate rapid punches (< 5 minutes apart)
    filtered = []
    for p in sorted_punches:
        p_time = p['punch_time'] if isinstance(p, dict) else p.punch_time
        if not filtered:
            filtered.append(p)
        else:
            prev_time = filtered[-1]['punch_time'] if isinstance(filtered[-1], dict) else filtered[-1].punch_time
            if (p_time - prev_time).total_seconds() > 300: # > 5 min
                filtered.append(p)

    sessions = []
    i = 0
    n = len(filtered)
    
    while i < n:
        curr_p = filtered[i]
        p_in = curr_p['punch_time'] if isinstance(curr_p, dict) else curr_p.punch_time
        in_hour = p_in.hour + p_in.minute / 60.0
        
        # Look ahead for out punch within max duty window (18 hours)
        candidate_outs = []
        j = i + 1
        while j < n:
            next_p = filtered[j]
            next_time = next_p['punch_time'] if isinstance(next_p, dict) else next_p.punch_time
            diff_secs = (next_time - p_in).total_seconds()
            if diff_secs <= 18 * 3600:
                candidate_outs.append((j, next_p, next_time))
                j += 1
            else:
                break
                
        p_out = None
        out_idx = None
        if candidate_outs:
            # Candidates at least 45 minutes after in
            valid_outs = [c for c in candidate_outs if (c[2] - p_in).total_seconds() >= 2700]
            if valid_outs:
                # Furthest punch within 18 hours is the check-out
                out_idx, out_p, p_out = valid_outs[-1]
                i = out_idx + 1
            else:
                # Only quick successive punch < 45 min
                out_idx, out_p, p_out = candidate_outs[-1]
                i = out_idx + 1
        else:
            i += 1
            
        duty_date = p_in.date()
        is_cross_day = (p_out is not None and p_out.date() > p_in.date())
        
        if p_out:
            duration = (p_out - p_in).total_seconds() / 3600.0
        else:
            duration = 0.0

        # AI Classification Logic
        is_night = False
        
        if in_hour >= 18.0 or in_hour < 4.0:
            # Night entry (18:00 - 04:00)
            is_night = True
            if is_cross_day and duration >= 14.0:
                shift_type = 'Double Shift (Night-to-Day)'
                status = 'Night Duty + OT'
                ai_note = f'🤖 Night-to-Day Continuous Duty: Checked in {p_in.strftime("%H:%M")} (Night), stayed through following day exiting at {p_out.strftime("%H:%M")}. Total: {duration:.1f} hrs.'
            elif is_cross_day or (p_out and p_out.hour <= 11):
                shift_type = 'Night Shift'
                status = 'Night Duty'
                ai_note = f'🤖 Night Shift (Cross-Day): Checked in at {p_in.strftime("%H:%M")}, worked through overnight shift, exiting next morning at {p_out.strftime("%H:%M")} ({duration:.1f} hrs).'
            else:
                shift_type = 'Night Shift'
                status = 'Night Duty' if p_out else 'Missing Out'
                ai_note = f'🤖 Night Shift: Checked in at {p_in.strftime("%H:%M")}.' + (f' Exited at {p_out.strftime("%H:%M")} ({duration:.1f} hrs).' if p_out else ' Single punch (missing checkout).')
        elif 13.0 <= in_hour < 18.0:
            # Afternoon / Evening shift (13:00 - 18:00)
            shift_type = 'Evening Shift'
            status = 'Present' if p_out else 'Missing Out'
            ai_note = f'🤖 Evening Shift: Checked in at {p_in.strftime("%H:%M")}.' + (f' Exited at {p_out.strftime("%H:%M")} ({duration:.1f} hrs).' if p_out else ' Single punch (missing checkout).')
        else:
            # Morning / Day shift (04:00 - 13:00)
            if is_cross_day or duration >= 14.0:
                shift_type = 'Double Shift (Day-to-Night)'
                status = 'Double Shift'
                is_night = True
                ai_note = f'🤖 Day-to-Night Double Shift: Entered morning at {p_in.strftime("%H:%M")}, worked continuously through night exiting at {p_out.strftime("%H:%M")}. Total: {duration:.1f} hrs.'
            else:
                shift_type = 'Day Shift'
                is_late = (in_hour > (9.0 + grace_period_mins / 60.0))
                status = 'Late' if is_late else ('Present' if p_out else 'Missing Out')
                ai_note = f'🤖 Regular Day Shift: In at {p_in.strftime("%H:%M")}, Out at {p_out.strftime("%H:%M") if p_out else "--:--"} ({duration:.1f} hrs).' + (' [Late Arrival]' if is_late else ' [Punctual Entry]')

        ot_hours = max(0.0, round(duration - standard_hours, 2)) if p_out else 0.0
        
        # Calculate human-readable duration
        hours_int = int(duration)
        mins_int = int(round((duration - hours_int) * 60))
        duration_fmt = f"{hours_int}h {mins_int:02d}m" if p_out else "-"
        
        sessions.append({
            'attendance_date': duty_date,
            'shift_type': shift_type,
            'in_time': p_in,
            'out_time': p_out,
            'duration_hours': round(duration, 2),
            'duration_formatted': duration_fmt,
            'overtime_hours': ot_hours,
            'status': status,
            'is_night_shift': is_night,
            'is_cross_day': is_cross_day,
            'ai_analysis': ai_note,
            'punch_count': 2 if p_out else 1
        })
        
    return sessions


def build_monthly_attendance_matrix(attendance_records, employees, year, month):
    """
    Constructs a 1-to-31 days calendar matrix for all employees in a given year and month.
    Seamlessly incorporates Government Holiday intelligence from GovtHoliday database.
    Holidays are automatically marked as 'H' (Paid Holiday), preventing false absences,
    and duties performed on holidays are tracked as 'H*' (Holiday Duty).
    """
    import calendar
    days_in_month = calendar.monthrange(year, month)[1]
    date_list = [date(year, month, d) for d in range(1, days_in_month + 1)]
    
    # Retrieve Govt Holidays for this month
    try:
        from govt_holiday_manager import govt_holiday_manager
        holiday_map = govt_holiday_manager.get_holiday_map_for_month(year, month)
    except Exception:
        holiday_map = {}

    # Map attendance by (employee_id, attendance_date)
    att_map = {}
    for r in attendance_records:
        key = (r.employee_id, r.attendance_date)
        att_map[key] = r
        
    matrix = []
    
    for emp in employees:
        emp_row = {
            'employee': emp,
            'days': {},
            'total_present': 0,
            'total_holidays': 0,
            'total_holiday_worked': 0,
            'total_weekends': 0,
            'total_leaves': 0,
            'total_absent': 0,
            'total_payable_days': 0,
            'total_night': 0,
            'total_day': 0,
            'total_double': 0,
            'total_hours': 0.0,
            'total_ot': 0.0,
            'late_days': 0
        }
        
        for d in date_list:
            day_num = d.day
            weekday = d.weekday() # 4 is Friday in Bangladesh weekend
            is_weekend = (weekday in [4, 5]) # Friday & Saturday (or Friday)
            
            rec = att_map.get((emp.id, d))
            h_obj = holiday_map.get(d)
            is_govt_holiday = (h_obj is not None)
            
            if rec:
                st_lower = (rec.status or '').lower()
                sh_lower = (rec.shift_type or '').lower()
                rem_lower = (rec.remarks or '').lower()

                # 1. Leave Posting Check
                if any(lv in st_lower for lv in ['leave', 'casual', 'sick', 'earned', 'maternity', 'paternity']) or 'leave' in sh_lower:
                    emp_row['total_leaves'] += 1
                    if 'casual' in st_lower or 'cl' in st_lower:
                        badge = 'CL'
                    elif 'sick' in st_lower or 'sl' in st_lower:
                        badge = 'SL'
                    elif 'earned' in st_lower or 'el' in st_lower:
                        badge = 'EL'
                    else:
                        badge = 'LV'
                    css_class = 'badge-leave'
                    shift_title = rec.remarks or f"Leave: {rec.status}"

                # 2. Duty Swap Check
                elif 'swap off' in st_lower or 'swap off' in sh_lower or 'swap off' in rem_lower:
                    emp_row['total_weekends'] += 1
                    badge = 'SO'
                    css_class = 'badge-weekend'
                    shift_title = rec.remarks or "Duty Swap: Off-Day"
                elif 'swap' in st_lower or 'swap' in sh_lower or 'swap' in rem_lower:
                    emp_row['total_present'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'SW'
                    css_class = 'badge-swap'
                    shift_title = rec.remarks or f"Duty Swap: {rec.shift_type}"

                # 3. Manual Adjustment Check
                elif 'adjust' in st_lower or 'force' in st_lower or 'tour' in st_lower or 'waived' in st_lower:
                    emp_row['total_present'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    if 'half' in st_lower:
                        badge = 'HD'
                        css_class = 'badge-warning'
                    else:
                        badge = 'AD'
                        css_class = 'badge-adjusted'
                    shift_title = rec.remarks or f"Adjusted: {rec.status}"

                # 4. Standard Shifts & Holidays
                elif is_govt_holiday:
                    emp_row['total_present'] += 1
                    emp_row['total_holiday_worked'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'H*'
                    css_class = 'badge-holiday-work'
                    shift_title = f"Govt Holiday Duty: {h_obj.name}"
                elif 'Night' in rec.shift_type:
                    emp_row['total_present'] += 1
                    emp_row['total_night'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'N'
                    css_class = 'badge-night'
                    shift_title = rec.shift_type
                elif 'Double' in rec.shift_type:
                    emp_row['total_present'] += 1
                    emp_row['total_double'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'DN'
                    css_class = 'badge-double'
                    shift_title = rec.shift_type
                elif 'Late' in (rec.status or ''):
                    emp_row['total_present'] += 1
                    emp_row['total_day'] += 1
                    emp_row['late_days'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'L'
                    css_class = 'badge-late'
                    shift_title = rec.shift_type
                elif rec.status == 'Missing Out':
                    emp_row['total_present'] += 1
                    badge = '?'
                    css_class = 'badge-warning'
                    shift_title = 'Missing Out'
                else:
                    emp_row['total_present'] += 1
                    emp_row['total_day'] += 1
                    emp_row['total_hours'] += (rec.total_hours or 0.0)
                    emp_row['total_ot'] += (rec.overtime_hours or 0.0)
                    badge = 'D'
                    css_class = 'badge-day'
                    shift_title = rec.shift_type
                    
                emp_row['days'][day_num] = {
                    'badge': badge,
                    'css_class': css_class,
                    'shift_type': shift_title,
                    'in_time': rec.in_time.strftime('%H:%M') if rec.in_time else '-',
                    'out_time': rec.out_time.strftime('%H:%M') if rec.out_time else '-',
                    'hours': rec.total_hours or 0.0,
                    'ot': rec.overtime_hours or 0.0,
                    'status': f"Holiday Duty: {h_obj.name}" if is_govt_holiday else rec.status,
                    'is_holiday': is_govt_holiday,
                    'holiday_name': h_obj.name if is_govt_holiday else '',
                    'ai_analysis': rec.ai_analysis
                }
            else:
                if is_govt_holiday:
                    emp_row['total_holidays'] += 1
                    emp_row['days'][day_num] = {
                        'badge': 'H',
                        'css_class': 'badge-holiday',
                        'shift_type': f"Govt Holiday: {h_obj.name}",
                        'hours': 0.0,
                        'ot': 0.0,
                        'status': f"Govt Holiday: {h_obj.name}",
                        'is_holiday': True,
                        'holiday_name': h_obj.name
                    }
                elif is_weekend:
                    emp_row['total_weekends'] += 1
                    emp_row['days'][day_num] = {
                        'badge': 'W',
                        'css_class': 'badge-weekend',
                        'shift_type': 'Weekend Off',
                        'hours': 0.0,
                        'ot': 0.0,
                        'status': 'Weekend',
                        'is_holiday': False
                    }
                else:
                    emp_row['total_absent'] += 1
                    emp_row['days'][day_num] = {
                        'badge': 'A',
                        'css_class': 'badge-absent',
                        'shift_type': 'Absent',
                        'hours': 0.0,
                        'ot': 0.0,
                        'status': 'Absent',
                        'is_holiday': False
                    }
                    
        emp_row['total_hours'] = round(emp_row['total_hours'], 1)
        emp_row['total_ot'] = round(emp_row['total_ot'], 1)
        # Total payable days = Days Worked + Paid Govt Holidays + Weekends + Approved Leaves
        emp_row['total_payable_days'] = emp_row['total_present'] + emp_row['total_holidays'] + emp_row['total_weekends'] + emp_row.get('total_leaves', 0)
        matrix.append(emp_row)
        
    return {
        'year': year,
        'month': month,
        'days_in_month': days_in_month,
        'date_list': date_list,
        'holiday_map': holiday_map,
        'total_govt_holidays_in_month': len(set(h.id for h in holiday_map.values())),
        'matrix': matrix
    }


def generate_attendance_excel_export(matrix_data):
    """
    Generates a beautifully styled Excel workbook for the monthly attendance register,
    with distinct visual highlighting for Government Holidays and accurate summary statistics.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Attendance_{matrix_data['year']}_{matrix_data['month']:02d}"

    # Header style
    dark_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    dark_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    holiday_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    holiday_font = Font(name="Calibri", size=10, bold=True, color="92400E")

    holiday_work_fill = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
    holiday_work_font = Font(name="Calibri", size=10, bold=True, color="B45309")

    weekend_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    weekend_font = Font(name="Calibri", size=9, color="64748B")

    # Title block
    ws.merge_cells("A1:I1")
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"MONTHLY BIOMETRIC ATTENDANCE REGISTER - {matrix_data['year']}-{matrix_data['month']:02d}"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="0F172A")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    holiday_map = matrix_data.get('holiday_map', {})

    headers = ["Emp ID", "Employee Name", "Department", "Designation"]
    for d in matrix_data['date_list']:
        h_obj = holiday_map.get(d)
        if h_obj:
            headers.append(f"{d.day} (H)")
        else:
            headers.append(str(d.day))

    headers.extend([
        "Present (P)", "Holidays (H)", "Weekend (W)", "Absent (A)", "Payable Days",
        "Day", "Night", "Total Hrs", "OT Hrs", "Late"
    ])

    row_num = 3
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=row_num, column=col_idx)
        c.value = h
        c.fill = dark_fill
        c.font = dark_font
        c.alignment = center_align
        c.border = thin_border

    row_num = 4
    for item in matrix_data['matrix']:
        emp = item['employee']
        clean_eid = str(emp.employee_id).replace('EMP-', '') if str(emp.employee_id).startswith('EMP-') else str(emp.employee_id)
        
        row_vals = [clean_eid, emp.name, emp.department or '-', emp.designation or '-']
        for d in matrix_data['date_list']:
            day_info = item['days'].get(d.day, {})
            row_vals.append(day_info.get('badge', '-'))
            
        row_vals.extend([
            item['total_present'],
            item['total_holidays'],
            item['total_weekends'],
            item['total_absent'],
            item['total_payable_days'],
            item['total_day'],
            item['total_night'],
            item['total_hours'],
            item['total_ot'],
            item['late_days']
        ])
        
        ws.append(row_vals)
        for col_i in range(1, len(row_vals) + 1):
            c_cell = ws.cell(row=row_num, column=col_i)
            c_cell.border = thin_border
            if col_i > 4:
                c_cell.alignment = center_align
                val_str = str(c_cell.value or '')
                if val_str == 'H':
                    c_cell.fill = holiday_fill
                    c_cell.font = holiday_font
                elif val_str == 'H*':
                    c_cell.fill = holiday_work_fill
                    c_cell.font = holiday_work_font
                elif val_str == 'W':
                    c_cell.fill = weekend_fill
                    c_cell.font = weekend_font
        row_num += 1

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col[0].column <= 4:
            ws.column_dimensions[col_letter].width = 18
        elif col[0].column > len(headers) - 10:
            ws.column_dimensions[col_letter].width = 11
        else:
            ws.column_dimensions[col_letter].width = 4.8

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio
