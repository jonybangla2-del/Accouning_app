"""
Leave Manager & Policy Accrual Engine (50% Duty General PTO & Holiday PTO)
==========================================================================
Comprehensive Leave Management system providing:
1. 50% Duty General PTO Accrual (1.5 days accrued if duty days >= 50% of monthly working days).
2. Holiday PTO Compensatory Reward (+1.0 day accrued per government holiday duty worked).
3. Category-wise Leave Deductions (Sick, General/Casual, Parental, Marriage, Adjustments).
4. Continuous 12-Month Running Balance Cascading (Opening -> Earned -> Enjoyed -> Closing).
5. Year-end December Closing with 50% Carry Forward to next year's opening bank.
6. Custom Block-Format Multi-Employee Excel Importer & Template Exporter.
7. Executive Excel Statement & All-Staff Summary Exporters.
"""

import os
import io
import re
import calendar
import logging
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']


def get_working_days_in_month(year: int, month: int) -> int:
    """
    Calculates monthly working days excluding standard weekend days (Friday & Saturday in Bangladesh).
    Typically 20 to 23 working days per month.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    working_days = 0
    for day in range(1, days_in_month + 1):
        cur_date = date(year, month, day)
        # weekday(): Monday is 0, Sunday is 6. Friday is 4, Saturday is 5.
        if cur_date.weekday() not in (4, 5):
            working_days += 1
    return working_days if working_days > 0 else 22


def calculate_monthly_duty_and_pto(emp_id: int, year: int, month: int, db_session,
                                    EmployeeAttendance, AttendanceAdjustment, GovtHoliday, LeavePosting):
    """
    Calculates duty attendance days, 50% duty eligibility for 1.5 PTO,
    compensatory Govt Holiday PTO, and enjoyed leave categories for an employee in a given month.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    m_start = date(year, month, 1)
    m_end = date(year, month, days_in_month)
    working_days = get_working_days_in_month(year, month)

    # 1. Fetch attendance records for this month
    att_records = EmployeeAttendance.query.filter(
        EmployeeAttendance.employee_id == emp_id,
        EmployeeAttendance.attendance_date >= m_start,
        EmployeeAttendance.attendance_date <= m_end
    ).all()

    att_map = {rec.attendance_date: rec for rec in att_records}

    # Also fetch manual adjustments in this month
    adjustments = AttendanceAdjustment.query.filter(
        AttendanceAdjustment.employee_id == emp_id,
        AttendanceAdjustment.target_date >= m_start,
        AttendanceAdjustment.target_date <= m_end
    ).all()
    adj_map = {adj.target_date: adj for adj in adjustments}

    # Count actual duty days
    duty_dates = set()
    present_statuses = {
        'present', 'present (adjusted)', 'night shift', 'night duty',
        'day shift', 'evening shift', 'regular day shift', 'late'
    }

    for d_date, rec in att_map.items():
        st = (rec.status or '').strip().lower()
        shift = (rec.shift_type or '').strip().lower()
        hrs = float(rec.total_hours or 0.0)

        if st in present_statuses or 'present' in st or hrs > 0:
            # Exclude explicit pure leave or pure absent
            if 'leave' not in st and 'absent' not in st:
                duty_dates.add(d_date)

    for d_date, adj in adj_map.items():
        new_st = (adj.new_status or '').strip().lower()
        if 'present' in new_st or float(adj.adjusted_hours or 0.0) > 0:
            duty_dates.add(d_date)

    duty_days_count = len(duty_dates)
    duty_percentage = round((duty_days_count / float(working_days)) * 100.0, 1) if working_days > 0 else 0.0

    # Rule 1: 50% duty criteria -> 1.5 PTO if duty >= 50% of working days (or >= 11 days), else 0.0
    half_working_days = round(working_days * 0.5, 1)
    is_eligible_pto = (duty_days_count >= half_working_days) or (duty_percentage >= 50.0)
    general_pto = 1.5 if is_eligible_pto else 0.0

    # Rule 2: Govt Holiday Duty Reward (+1.0 day per holiday worked)
    holiday_pto = 0.0
    govt_holidays = GovtHoliday.query.filter(
        GovtHoliday.year == year
    ).all()

    holiday_dates = set()
    for gh in govt_holidays:
        h_start = gh.holiday_date
        h_end = gh.end_date or gh.holiday_date
        if h_start and h_end:
            cur_h = h_start
            while cur_h <= h_end:
                if m_start <= cur_h <= m_end:
                    holiday_dates.add(cur_h)
                cur_h += timedelta(days=1)

    for h_date in holiday_dates:
        if h_date in duty_dates:
            holiday_pto += 1.0

    # Rule 3: Calculate enjoyed leaves from LeavePosting overlapping this month
    leave_postings = LeavePosting.query.filter(
        LeavePosting.employee_id == emp_id,
        LeavePosting.status == 'Approved',
        LeavePosting.date_from <= m_end,
        LeavePosting.date_to >= m_start
    ).all()

    sick_leave = 0.0
    general_leave = 0.0
    parental_leave = 0.0
    marriage_leave = 0.0
    other_leave = 0.0

    for lp in leave_postings:
        cur_l = max(lp.date_from, m_start)
        end_l = min(lp.date_to, m_end)
        days = (end_l - cur_l).days + 1
        l_type = (lp.leave_type or '').strip().lower()

        if 'sick' in l_type:
            sick_leave += float(days)
        elif 'general' in l_type or 'casual' in l_type or 'earned' in l_type or 'pto' in l_type:
            general_leave += float(days)
        elif 'parental' in l_type or 'maternity' in l_type or 'paternity' in l_type:
            parental_leave += float(days)
        elif 'marriage' in l_type:
            marriage_leave += float(days)
        else:
            other_leave += float(days)

    return {
        'duty_days_count': duty_days_count,
        'working_days': working_days,
        'duty_percentage': duty_percentage,
        'is_eligible_pto': is_eligible_pto,
        'general_pto': general_pto,
        'holiday_pto': holiday_pto,
        'sick_leave': sick_leave,
        'general_leave': general_leave,
        'parental_leave': parental_leave,
        'marriage_leave': marriage_leave,
        'other_leave': other_leave
    }


def recalculate_employee_leave_ledger(emp_id: int, year: int, db_session, models, force_recalculate=False):
    """
    Recalculates an employee's complete 12-month leave ledger for a given year.
    Cascades opening balances through Jan to Dec, applies policy rules,
    computes closing balances, and calculates the 50% year-end carry forward.
    """
    LeaveBank = models['LeaveBank']
    LeaveMonthlyLedger = models['LeaveMonthlyLedger']
    Employee = models['Employee']
    EmployeeAttendance = models['EmployeeAttendance']
    AttendanceAdjustment = models['AttendanceAdjustment']
    GovtHoliday = models['GovtHoliday']
    LeavePosting = models['LeavePosting']

    emp = Employee.query.get(emp_id)
    if not emp:
        return None

    # Fetch or create LeaveBank
    bank = LeaveBank.query.filter_by(employee_id=emp_id, year=year).first()
    if not bank:
        bank = LeaveBank(
            company_id=1,
            employee_id=emp_id,
            emp_code=emp.employee_id,
            emp_name=emp.name,
            department=emp.department or 'General',
            designation=emp.designation or 'Staff',
            year=year,
            opening_balance=0.0
        )
        db_session.add(bank)
        db_session.flush()

    # Fetch or create 12 monthly ledger rows
    ledger_rows = LeaveMonthlyLedger.query.filter_by(
        employee_id=emp_id,
        year=year
    ).order_by(LeaveMonthlyLedger.month.asc()).all()

    month_to_row = {row.month: row for row in ledger_rows}

    # Ensure all 12 months exist
    for m in range(1, 13):
        if m not in month_to_row:
            new_row = LeaveMonthlyLedger(
                company_id=1,
                employee_id=emp_id,
                emp_code=emp.employee_id,
                emp_name=emp.name,
                department=emp.department or 'General',
                year=year,
                month=m,
                month_name=MONTH_NAMES[m - 1],
                opening_balance=0.0,
                closing_balance=0.0
            )
            db_session.add(new_row)
            db_session.flush()
            month_to_row[m] = new_row

    # Cascade running balances from Jan (1) to Dec (12)
    cur_opening = float(bank.opening_balance or 0.0)

    total_general_pto = 0.0
    total_holiday_pto = 0.0
    total_sick = 0.0
    total_general = 0.0
    total_parental = 0.0
    total_marriage = 0.0
    total_other = 0.0
    total_adj = 0.0

    today = date.today()
    latest_active_balance = cur_opening

    for m in range(1, 13):
        row = month_to_row[m]
        row.opening_balance = round(cur_opening, 2)

        if force_recalculate or not row.is_manually_locked:
            calc = calculate_monthly_duty_and_pto(
                emp_id=emp_id,
                year=year,
                month=m,
                db_session=db_session,
                EmployeeAttendance=EmployeeAttendance,
                AttendanceAdjustment=AttendanceAdjustment,
                GovtHoliday=GovtHoliday,
                LeavePosting=LeavePosting
            )
            row.duty_days_count = calc['duty_days_count']
            row.total_working_days = calc['working_days']
            row.duty_percentage = calc['duty_percentage']
            row.is_eligible_pto = calc['is_eligible_pto']
            row.general_pto = calc['general_pto']
            row.holiday_pto = calc['holiday_pto']
            row.sick_leave = calc['sick_leave']
            row.general_leave = calc['general_leave']
            row.parental_leave = calc['parental_leave']
            row.marriage_leave = calc['marriage_leave']
            row.other_leave = calc['other_leave']

        # Total deductions for month
        total_deductions = (
            float(row.sick_leave or 0.0) +
            float(row.general_leave or 0.0) +
            float(row.parental_leave or 0.0) +
            float(row.marriage_leave or 0.0) +
            float(row.other_leave or 0.0) +
            float(row.adjustment or 0.0)
        )

        row.closing_balance = round(
            row.opening_balance +
            float(row.general_pto or 0.0) +
            float(row.holiday_pto or 0.0) -
            total_deductions,
            2
        )

        # Accumulate annual totals
        total_general_pto += float(row.general_pto or 0.0)
        total_holiday_pto += float(row.holiday_pto or 0.0)
        total_sick += float(row.sick_leave or 0.0)
        total_general += float(row.general_leave or 0.0)
        total_parental += float(row.parental_leave or 0.0)
        total_marriage += float(row.marriage_leave or 0.0)
        total_other += float(row.other_leave or 0.0)
        total_adj += float(row.adjustment or 0.0)

        # Track active balance up to current month
        if year < today.year or (year == today.year and m <= today.month):
            latest_active_balance = row.closing_balance

        # Next month's opening is this month's closing
        cur_opening = row.closing_balance

    # Rule 5: Year Closing & 50% Carry Forward
    dec_closing = month_to_row[12].closing_balance
    carried_forward = round(max(0.0, dec_closing * 0.5), 2)

    bank.total_general_pto = round(total_general_pto, 2)
    bank.total_holiday_pto = round(total_holiday_pto, 2)
    bank.total_sick_leave = round(total_sick, 2)
    bank.total_general_leave = round(total_general, 2)
    bank.total_parental_leave = round(total_parental, 2)
    bank.total_marriage_leave = round(total_marriage, 2)
    bank.total_other_leave = round(total_other, 2)
    bank.total_adjustments = round(total_adj, 2)
    bank.current_balance = round(latest_active_balance, 2)
    bank.carried_forward_to_next_year = carried_forward
    bank.updated_at = datetime.utcnow()

    db_session.commit()
    return bank


def close_leave_year(emp_id: int, year: int, db_session, models):
    """
    Closes the leave year for an employee and rolls over 50% of the December
    closing balance into next year's (year + 1) Opening Leave Bank.
    """
    LeaveBank = models['LeaveBank']
    bank = recalculate_employee_leave_ledger(emp_id, year, db_session, models)
    if not bank:
        return None

    bank.is_year_closed = True
    carried_forward = bank.carried_forward_to_next_year

    # Roll over to next year
    next_bank = LeaveBank.query.filter_by(employee_id=emp_id, year=year + 1).first()
    if not next_bank:
        next_bank = LeaveBank(
            company_id=1,
            employee_id=emp_id,
            emp_code=bank.emp_code,
            emp_name=bank.emp_name,
            department=bank.department,
            designation=bank.designation,
            year=year + 1,
            opening_balance=carried_forward
        )
        db_session.add(next_bank)
    else:
        next_bank.opening_balance = carried_forward

    db_session.commit()
    recalculate_employee_leave_ledger(emp_id, year + 1, db_session, models)
    return bank


def parse_custom_block_leave_excel(file_stream_or_path, db_session, models, company_id=1):
    """
    Parses multi-employee custom block-format Excel spreadsheets:
    - Block Header: Name, ID, Designation, Organization, Date of Joining, Year, Opening Balance.
    - Monthly Grid Table (Jan to Dec): Month, Leave Bank, General PTO, Holiday PTO,
      Sick Leave, General Leave, Parental Leave, Marriage Leave, Adjustment, Balance, Remarks.
    Stores and cascades records into LeaveBank and LeaveMonthlyLedger.
    """
    LeaveBank = models['LeaveBank']
    LeaveMonthlyLedger = models['LeaveMonthlyLedger']
    Employee = models['Employee']

    wb = openpyxl.load_workbook(file_stream_or_path, data_only=True)
    imported_employees = []
    total_blocks_found = 0

    # Clean helper
    def clean_val(v):
        if v is None:
            return ""
        return str(v).strip()

    def to_float(v, default=0.0):
        if v is None:
            return default
        try:
            # Handle string numbers with comma
            cleaned = str(v).replace(',', '').strip()
            return float(cleaned)
        except (ValueError, TypeError):
            return default

    for sheet in wb.worksheets:
        max_r = sheet.max_row
        r = 1
        while r <= max_r:
            # Look for employee block indicator: row with "Name" or "Employee Name" or "ID"
            row_vals = [clean_val(sheet.cell(row=r, column=c).value) for c in range(1, 15)]
            row_str = " ".join(row_vals).lower()

            if ('name:' in row_str or 'name' in row_str) and ('id:' in row_str or 'id' in row_str or 'employee' in row_str):
                # We detected a candidate header!
                # Extract header fields by scanning this row and neighboring 3 rows
                block_info = {
                    'name': '',
                    'emp_code': '',
                    'designation': '',
                    'organization': '',
                    'joining_date': None,
                    'year': 2026,
                    'opening_balance': 0.0
                }

                for hr in range(r, min(r + 5, max_r + 1)):
                    for c in range(1, sheet.max_column + 1):
                        cell_txt = clean_val(sheet.cell(row=hr, column=c).value).lower()
                        cell_raw = sheet.cell(row=hr, column=c).value

                        # Name
                        if re.search(r'^(name|employee\s*name)\s*[:\-]?', cell_txt):
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            if next_val:
                                block_info['name'] = clean_val(next_val)
                            elif ':' in clean_val(cell_raw):
                                block_info['name'] = clean_val(cell_raw).split(':', 1)[1].strip()

                        # ID
                        if re.search(r'^(id|emp\s*id|employee\s*id)\s*[:\-]?', cell_txt):
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            if next_val:
                                block_info['emp_code'] = clean_val(next_val)
                            elif ':' in clean_val(cell_raw):
                                block_info['emp_code'] = clean_val(cell_raw).split(':', 1)[1].strip()

                        # Designation
                        if 'designation' in cell_txt:
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            if next_val:
                                block_info['designation'] = clean_val(next_val)
                            elif ':' in clean_val(cell_raw):
                                block_info['designation'] = clean_val(cell_raw).split(':', 1)[1].strip()

                        # Organization
                        if 'organization' in cell_txt or 'company' in cell_txt:
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            if next_val:
                                block_info['organization'] = clean_val(next_val)

                        # Year
                        if 'year' in cell_txt:
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            try:
                                if next_val:
                                    block_info['year'] = int(str(next_val).strip()[:4])
                            except (ValueError, TypeError):
                                pass

                        # Opening Balance / Remaining Leave Bank
                        if any(term in cell_txt for term in ['opening balance', 'remaining leave bank', 'leave bank', 'opening bank']):
                            next_val = sheet.cell(row=hr, column=c + 1).value
                            if next_val is not None:
                                block_info['opening_balance'] = to_float(next_val, 0.0)

                # Normalize emp_code
                emp_code = str(block_info['emp_code']).strip()
                emp_code_clean = re.sub(r'^(emp[\-_]?)', '', emp_code, flags=re.I).strip()
                emp_name = block_info['name'] or f"Employee {emp_code}"

                if not emp_code and not emp_name:
                    r += 1
                    continue

                # Find employee in DB or create stub
                emp = None
                if emp_code:
                    emp = Employee.query.filter(
                        (Employee.employee_id == emp_code) |
                        (Employee.employee_id == f"EMP-{emp_code_clean}") |
                        (Employee.employee_id == emp_code_clean)
                    ).first()
                if not emp and emp_name:
                    emp = Employee.query.filter(Employee.name.ilike(f"%{emp_name}%")).first()

                if not emp:
                    # Create new active employee record
                    new_code = emp_code if emp_code else f"EMP-{emp_code_clean or 'NEW'}"
                    emp = Employee(
                        employee_id=new_code,
                        name=emp_name,
                        designation=block_info['designation'] or 'Staff',
                        department='General',
                        status='Active'
                    )
                    db_session.add(emp)
                    db_session.flush()

                target_year = block_info['year'] or 2026

                # Search downward for table header row containing "month"
                table_header_row = None
                for tr in range(r, min(r + 10, max_r + 1)):
                    r_txts = [clean_val(sheet.cell(row=tr, column=c).value).lower() for c in range(1, 15)]
                    if any('month' in t or 'মাস' in t for t in r_txts):
                        table_header_row = tr
                        break

                if not table_header_row:
                    r += 1
                    continue

                # Map column indices from the table header row
                col_map = {}
                for c in range(1, sheet.max_column + 1):
                    c_txt = clean_val(sheet.cell(row=table_header_row, column=c).value).lower()
                    if 'month' in c_txt or 'মাস' in c_txt:
                        col_map['month'] = c
                    elif 'opening' in c_txt or 'leave bank' in c_txt or 'bank' in c_txt:
                        col_map.setdefault('opening_balance', c)
                    elif 'general pto' in c_txt or 'pto' in c_txt or 'gen pto' in c_txt:
                        col_map['general_pto'] = c
                    elif 'holiday' in c_txt or 'hol' in c_txt:
                        col_map['holiday_pto'] = c
                    elif 'sick' in c_txt:
                        col_map['sick_leave'] = c
                    elif 'general' in c_txt or 'casual' in c_txt:
                        col_map['general_leave'] = c
                    elif 'parental' in c_txt or 'maternity' in c_txt or 'paternity' in c_txt:
                        col_map['parental_leave'] = c
                    elif 'marriage' in c_txt:
                        col_map['marriage_leave'] = c
                    elif 'adjust' in c_txt or 'deduct' in c_txt:
                        col_map['adjustment'] = c
                    elif 'balance' in c_txt or 'closing' in c_txt:
                        col_map['balance'] = c
                    elif 'remark' in c_txt or 'note' in c_txt:
                        col_map['remarks'] = c

                # Create or fetch LeaveBank
                bank = LeaveBank.query.filter_by(employee_id=emp.id, year=target_year).first()
                if not bank:
                    bank = LeaveBank(
                        company_id=company_id,
                        employee_id=emp.id,
                        emp_code=emp.employee_id,
                        emp_name=emp.name,
                        department=emp.department or 'General',
                        designation=emp.designation or block_info['designation'] or 'Staff',
                        year=target_year,
                        opening_balance=block_info['opening_balance']
                    )
                    db_session.add(bank)
                    db_session.flush()
                else:
                    if block_info['opening_balance'] > 0:
                        bank.opening_balance = block_info['opening_balance']

                # Read 12 monthly rows
                month_cur = table_header_row + 1
                months_read = 0
                running_balance = float(bank.opening_balance or 0.0)

                while month_cur <= max_r and months_read < 12:
                    m_label = clean_val(sheet.cell(row=month_cur, column=col_map.get('month', 1)).value)
                    if not m_label or 'total' in m_label.lower() or 'sub' in m_label.lower() or 'carry' in m_label.lower():
                        break

                    # Match month number (1 to 12)
                    m_num = None
                    m_norm = m_label[:3].lower()
                    for idx, mn in enumerate(MONTH_NAMES, 1):
                        if mn.lower() == m_norm or m_label.lower().startswith(mn.lower()):
                            m_num = idx
                            break
                    if not m_num:
                        try:
                            m_num = int(m_label)
                        except ValueError:
                            m_num = months_read + 1

                    if not (1 <= m_num <= 12):
                        month_cur += 1
                        continue

                    # Fetch or create ledger row
                    ledger_row = LeaveMonthlyLedger.query.filter_by(
                        employee_id=emp.id,
                        year=target_year,
                        month=m_num
                    ).first()
                    if not ledger_row:
                        ledger_row = LeaveMonthlyLedger(
                            company_id=company_id,
                            employee_id=emp.id,
                            emp_code=emp.employee_id,
                            emp_name=emp.name,
                            department=emp.department or 'General',
                            year=target_year,
                            month=m_num,
                            month_name=MONTH_NAMES[m_num - 1]
                        )
                        db_session.add(ledger_row)

                    # Extract row values
                    if 'opening_balance' in col_map:
                        sheet_open = sheet.cell(row=month_cur, column=col_map['opening_balance']).value
                        if sheet_open is not None and str(sheet_open).strip() != '':
                            ledger_row.opening_balance = to_float(sheet_open, running_balance)
                        else:
                            ledger_row.opening_balance = running_balance
                    else:
                        ledger_row.opening_balance = running_balance

                    if 'general_pto' in col_map:
                        ledger_row.general_pto = to_float(sheet.cell(row=month_cur, column=col_map['general_pto']).value, 0.0)
                    if 'holiday_pto' in col_map:
                        ledger_row.holiday_pto = to_float(sheet.cell(row=month_cur, column=col_map['holiday_pto']).value, 0.0)
                    if 'sick_leave' in col_map:
                        ledger_row.sick_leave = to_float(sheet.cell(row=month_cur, column=col_map['sick_leave']).value, 0.0)
                    if 'general_leave' in col_map:
                        ledger_row.general_leave = to_float(sheet.cell(row=month_cur, column=col_map['general_leave']).value, 0.0)
                    if 'parental_leave' in col_map:
                        ledger_row.parental_leave = to_float(sheet.cell(row=month_cur, column=col_map['parental_leave']).value, 0.0)
                    if 'marriage_leave' in col_map:
                        ledger_row.marriage_leave = to_float(sheet.cell(row=month_cur, column=col_map['marriage_leave']).value, 0.0)
                    if 'adjustment' in col_map:
                        ledger_row.adjustment = to_float(sheet.cell(row=month_cur, column=col_map['adjustment']).value, 0.0)
                    if 'remarks' in col_map:
                        ledger_row.remarks = clean_val(sheet.cell(row=month_cur, column=col_map['remarks']).value)

                    # Compute closing balance
                    tot_deduct = (
                        float(ledger_row.sick_leave or 0.0) + float(ledger_row.general_leave or 0.0) +
                        float(ledger_row.parental_leave or 0.0) + float(ledger_row.marriage_leave or 0.0) +
                        float(ledger_row.other_leave or 0.0) + float(ledger_row.adjustment or 0.0)
                    )
                    ledger_row.closing_balance = round(
                        float(ledger_row.opening_balance or 0.0) +
                        float(ledger_row.general_pto or 0.0) +
                        float(ledger_row.holiday_pto or 0.0) - tot_deduct,
                        2
                    )

                    # Lock row to preserve imported Excel values
                    ledger_row.is_manually_locked = True
                    running_balance = ledger_row.closing_balance

                    months_read += 1
                    month_cur += 1

                # Recompute totals and carry forward
                recalculate_employee_leave_ledger(emp.id, target_year, db_session, models, force_recalculate=False)

                imported_employees.append({
                    'id': emp.id,
                    'emp_code': emp.employee_id,
                    'name': emp.name,
                    'year': target_year,
                    'opening_balance': bank.opening_balance,
                    'closing_balance': bank.current_balance,
                    'carried_forward': bank.carried_forward_to_next_year
                })
                total_blocks_found += 1
                r = month_cur
            else:
                r += 1

    return {
        'total_blocks_found': total_blocks_found,
        'employees': imported_employees
    }


def generate_sample_leave_excel_template():
    """
    Generates a beautifully styled sample Excel workbook with employee blocks,
    pre-configured formulas (General PTO 50% rule, Running Balance, 50% Carry Forward),
    and guide instructions.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leave Ledger Template"
    ws.views.sheetView[0].showGridLines = True

    # Styling Palettes
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    navy_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    teal_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    accent_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    closing_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    sample_employees = [
        {"name": "Md Faysal Islam", "id": "85", "designation": "Sr. Video Editor", "org": "TBN BD LTD", "joining": "2020-01-15", "opening": 8.0},
        {"name": "Uzzal Chandra Sarker", "id": "143", "designation": "Channel Manager", "org": "TBN BD LTD", "joining": "2021-06-01", "opening": 12.0}
    ]

    cur_row = 2

    # Instructions Banner
    ws.merge_cells(start_row=cur_row, start_column=2, end_row=cur_row, end_column=12)
    b_cell = ws.cell(row=cur_row, column=2, value="COMPANY LEAVE LEDGER IMPORT TEMPLATE (50% Duty PTO + Holiday Reward + 50% Year Closing)")
    b_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    b_cell.fill = navy_fill
    b_cell.alignment = center_align
    cur_row += 2

    for emp in sample_employees:
        # Header Info Card
        ws.cell(row=cur_row, column=2, value="Name:").font = bold_font
        ws.cell(row=cur_row, column=3, value=emp['name']).font = bold_font
        ws.cell(row=cur_row, column=5, value="ID:").font = bold_font
        ws.cell(row=cur_row, column=6, value=emp['id']).font = bold_font
        ws.cell(row=cur_row, column=8, value="Year:").font = bold_font
        ws.cell(row=cur_row, column=9, value=2026).font = bold_font

        cur_row += 1
        ws.cell(row=cur_row, column=2, value="Designation:").font = regular_font
        ws.cell(row=cur_row, column=3, value=emp['designation']).font = regular_font
        ws.cell(row=cur_row, column=5, value="Organization:").font = regular_font
        ws.cell(row=cur_row, column=6, value=emp['org']).font = regular_font
        ws.cell(row=cur_row, column=8, value="Remaining Leave Bank:").font = bold_font
        ws.cell(row=cur_row, column=9, value=emp['opening']).font = bold_font

        cur_row += 2

        # Monthly Table Headers
        headers = [
            "Month", "Leave Bank", "General PTO", "Holiday PTO",
            "Sick Leave", "General Leave", "Parental Leave", "Marriage Leave",
            "Adjustment", "Balance", "Remarks"
        ]

        table_start_col = 2
        for col_idx, h_text in enumerate(headers, start=table_start_col):
            cell = ws.cell(row=cur_row, column=col_idx, value=h_text)
            cell.font = header_font
            cell.fill = teal_fill
            cell.alignment = center_align
            cell.border = thin_border

        table_header_row = cur_row
        cur_row += 1

        # 12 Month Rows
        start_month_row = cur_row
        for m_idx, m_name in enumerate(MONTH_NAMES):
            row_num = cur_row
            ws.cell(row=row_num, column=2, value=m_name).alignment = center_align
            ws.cell(row=row_num, column=2).font = bold_font
            ws.cell(row=row_num, column=2).border = thin_border

            # Opening Bank
            c_bank = ws.cell(row=row_num, column=3)
            c_bank.alignment = right_align
            c_bank.border = thin_border
            if m_idx == 0:
                c_bank.value = emp['opening']
            else:
                prev_row = row_num - 1
                c_bank.value = f"=K{prev_row}"  # previous month's balance

            # General PTO (sample default 1.5)
            c_pto = ws.cell(row=row_num, column=4, value=1.5)
            c_pto.alignment = right_align
            c_pto.border = thin_border

            # Holiday PTO (sample default 0.0)
            c_hol = ws.cell(row=row_num, column=5, value=1.0 if m_idx in [1, 3] else 0.0)
            c_hol.alignment = right_align
            c_hol.border = thin_border

            # Sick Leave
            c_sick = ws.cell(row=row_num, column=6, value=1.0 if m_idx == 4 else 0.0)
            c_sick.alignment = right_align
            c_sick.border = thin_border

            # General Leave
            c_gen = ws.cell(row=row_num, column=7, value=1.0 if m_idx == 2 else 0.0)
            c_gen.alignment = right_align
            c_gen.border = thin_border

            # Parental Leave
            c_par = ws.cell(row=row_num, column=8, value=0.0)
            c_par.alignment = right_align
            c_par.border = thin_border

            # Marriage Leave
            c_mar = ws.cell(row=row_num, column=9, value=0.0)
            c_mar.alignment = right_align
            c_mar.border = thin_border

            # Adjustment
            c_adj = ws.cell(row=row_num, column=10, value=0.0)
            c_adj.alignment = right_align
            c_adj.border = thin_border

            # Net Balance Formula: Opening + Gen PTO + Holiday PTO - (Sick + Gen + Par + Mar) - Adj
            c_bal = ws.cell(row=row_num, column=11)
            c_bal.value = f"=C{row_num}+D{row_num}+E{row_num}-SUM(F{row_num}:I{row_num})-J{row_num}"
            c_bal.alignment = right_align
            c_bal.font = bold_font
            c_bal.border = thin_border

            # Remarks
            c_rem = ws.cell(row=row_num, column=12, value="Normal Accrual" if m_idx != 4 else "Sick Leave taken")
            c_rem.alignment = left_align
            c_rem.font = regular_font
            c_rem.border = thin_border

            cur_row += 1

        # Summary Row: Total & Carry Forward
        dec_row = cur_row - 1
        ws.cell(row=cur_row, column=2, value="Total Accrued").font = bold_font
        ws.cell(row=cur_row, column=2).alignment = center_align
        ws.cell(row=cur_row, column=2).border = thin_border

        for col_i in range(3, 11):
            col_letter = get_column_letter(col_i)
            cell = ws.cell(row=cur_row, column=col_i, value=f"=SUM({col_letter}{start_month_row}:{col_letter}{dec_row})")
            cell.font = bold_font
            cell.alignment = right_align
            cell.fill = accent_fill
            cell.border = thin_border

        ws.cell(row=cur_row, column=11, value=f"=K{dec_row}").font = bold_font
        ws.cell(row=cur_row, column=11).alignment = right_align
        ws.cell(row=cur_row, column=11).fill = accent_fill
        ws.cell(row=cur_row, column=11).border = thin_border
        cur_row += 1

        # 50% Carry Forward Card
        ws.merge_cells(start_row=cur_row, start_column=2, end_row=cur_row, end_column=8)
        cf_label = ws.cell(row=cur_row, column=2, value="Year Closing Carry Forward to Next Year (50% of Dec Closing):")
        cf_label.font = bold_font
        cf_label.alignment = right_align
        cf_label.fill = closing_fill

        ws.merge_cells(start_row=cur_row, start_column=9, end_row=cur_row, end_column=11)
        cf_val = ws.cell(row=cur_row, column=9, value=f"=ROUND(K{dec_row}*0.5, 2)")
        cf_val.font = Font(name="Calibri", size=11, bold=True, color="047857")
        cf_val.alignment = center_align
        cf_val.fill = closing_fill

        cur_row += 3  # Gap before next employee

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A', 'M', 'N']:
            ws.column_dimensions[col_letter].width = 4
        else:
            ws.column_dimensions[col_letter].width = 16

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def export_employee_leave_statement_excel(emp_id: int, year: int, db_session, models):
    """
    Exports a single employee's comprehensive 12-month Leave Statement into a formatted Excel sheet.
    """
    LeaveBank = models['LeaveBank']
    LeaveMonthlyLedger = models['LeaveMonthlyLedger']
    Employee = models['Employee']

    emp = Employee.query.get(emp_id)
    if not emp:
        return None

    bank = recalculate_employee_leave_ledger(emp_id, year, db_session, models)
    ledger_rows = LeaveMonthlyLedger.query.filter_by(
        employee_id=emp_id,
        year=year
    ).order_by(LeaveMonthlyLedger.month.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Leave_{emp.employee_id}_{year}"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    navy_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    teal_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    accent_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    closing_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # Title Banner
    ws.merge_cells("B2:L2")
    ws.cell(row=2, column=2, value="TBN BD LTD - INDIVIDUAL EMPLOYEE LEAVE STATEMENT").alignment = center_align
    ws.cell(row=2, column=2).font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws.cell(row=2, column=2).fill = navy_fill

    # Subtitle
    ws.merge_cells("B3:L3")
    ws.cell(row=3, column=2, value=f"Leave Accrual & Balance Statement for Year {year} (50% Duty General PTO & Holiday Compensation)").alignment = center_align
    ws.cell(row=3, column=2).font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    ws.cell(row=3, column=2).fill = navy_fill

    # Profile Block
    ws.cell(row=5, column=2, value="Employee Name:").font = bold_font
    ws.cell(row=5, column=3, value=emp.name).font = bold_font
    ws.cell(row=5, column=5, value="ID:").font = bold_font
    ws.cell(row=5, column=6, value=emp.employee_id).font = bold_font
    ws.cell(row=5, column=8, value="Year:").font = bold_font
    ws.cell(row=5, column=9, value=year).font = bold_font

    ws.cell(row=6, column=2, value="Designation:").font = regular_font
    ws.cell(row=6, column=3, value=emp.designation or 'Staff').font = regular_font
    ws.cell(row=6, column=5, value="Department:").font = regular_font
    ws.cell(row=6, column=6, value=emp.department or 'General').font = regular_font
    ws.cell(row=6, column=8, value="Opening Leave Bank:").font = bold_font
    ws.cell(row=6, column=9, value=float(bank.opening_balance or 0.0)).font = bold_font

    # KPI Summary Row
    ws.merge_cells("B8:D8")
    ws.cell(row=8, column=2, value=f"Opening Bank: {bank.opening_balance:.1f} days").font = bold_font
    ws.cell(row=8, column=2).fill = accent_fill
    ws.cell(row=8, column=2).alignment = center_align

    ws.merge_cells("E8:G8")
    total_accrued = (bank.total_general_pto or 0.0) + (bank.total_holiday_pto or 0.0)
    ws.cell(row=8, column=5, value=f"Total PTO Earned: {total_accrued:.1f} days").font = bold_font
    ws.cell(row=8, column=5).fill = accent_fill
    ws.cell(row=8, column=5).alignment = center_align

    ws.merge_cells("H8:I8")
    total_enjoyed = (bank.total_sick_leave or 0.0) + (bank.total_general_leave or 0.0) + (bank.total_parental_leave or 0.0) + (bank.total_marriage_leave or 0.0)
    ws.cell(row=8, column=8, value=f"Total Leaves Taken: {total_enjoyed:.1f} days").font = bold_font
    ws.cell(row=8, column=8).fill = accent_fill
    ws.cell(row=8, column=8).alignment = center_align

    ws.merge_cells("J8:L8")
    ws.cell(row=8, column=10, value=f"Current Net Balance: {bank.current_balance:.1f} days").font = Font(name="Calibri", size=11, bold=True, color="047857")
    ws.cell(row=8, column=10).fill = closing_fill
    ws.cell(row=8, column=10).alignment = center_align

    # Table Headers
    headers = [
        "Month", "Duty Days", "Leave Bank", "General PTO", "Holiday PTO",
        "Sick Leave", "General Leave", "Parental Leave", "Marriage Leave",
        "Adjustment", "Balance", "Remarks"
    ]
    cur_row = 10
    for idx, h in enumerate(headers, start=2):
        cell = ws.cell(row=cur_row, column=idx, value=h)
        cell.font = header_font
        cell.fill = teal_fill
        cell.alignment = center_align
        cell.border = thin_border

    cur_row += 1
    table_start = cur_row

    for row in ledger_rows:
        ws.cell(row=cur_row, column=2, value=row.month_name).alignment = center_align
        ws.cell(row=cur_row, column=2).font = bold_font
        ws.cell(row=cur_row, column=2).border = thin_border

        # Duty days
        duty_txt = f"{int(row.duty_days_count or 0)}/{row.total_working_days or 22} ({row.duty_percentage or 0:.0f}%)"
        ws.cell(row=cur_row, column=3, value=duty_txt).alignment = center_align
        ws.cell(row=cur_row, column=3).font = regular_font
        ws.cell(row=cur_row, column=3).border = thin_border

        # Opening Bank
        ws.cell(row=cur_row, column=4, value=float(row.opening_balance or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=4).font = regular_font
        ws.cell(row=cur_row, column=4).border = thin_border

        # General PTO
        ws.cell(row=cur_row, column=5, value=float(row.general_pto or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=5).font = regular_font
        ws.cell(row=cur_row, column=5).border = thin_border

        # Holiday PTO
        ws.cell(row=cur_row, column=6, value=float(row.holiday_pto or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=6).font = regular_font
        ws.cell(row=cur_row, column=6).border = thin_border

        # Sick Leave
        ws.cell(row=cur_row, column=7, value=float(row.sick_leave or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=7).font = regular_font
        ws.cell(row=cur_row, column=7).border = thin_border

        # General Leave
        ws.cell(row=cur_row, column=8, value=float(row.general_leave or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=8).font = regular_font
        ws.cell(row=cur_row, column=8).border = thin_border

        # Parental Leave
        ws.cell(row=cur_row, column=9, value=float(row.parental_leave or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=9).font = regular_font
        ws.cell(row=cur_row, column=9).border = thin_border

        # Marriage Leave
        ws.cell(row=cur_row, column=10, value=float(row.marriage_leave or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=10).font = regular_font
        ws.cell(row=cur_row, column=10).border = thin_border

        # Adjustment
        ws.cell(row=cur_row, column=11, value=float(row.adjustment or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=11).font = regular_font
        ws.cell(row=cur_row, column=11).border = thin_border

        # Closing Balance
        ws.cell(row=cur_row, column=12, value=float(row.closing_balance or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=12).font = bold_font
        ws.cell(row=cur_row, column=12).border = thin_border

        # Remarks
        ws.cell(row=cur_row, column=13, value=row.remarks or '').alignment = left_align
        ws.cell(row=cur_row, column=13).font = regular_font
        ws.cell(row=cur_row, column=13).border = thin_border

        cur_row += 1

    # Totals Row
    ws.cell(row=cur_row, column=2, value="Annual Total").font = bold_font
    ws.cell(row=cur_row, column=2).alignment = center_align
    ws.cell(row=cur_row, column=2).border = thin_border

    ws.cell(row=cur_row, column=3, value="-").alignment = center_align
    ws.cell(row=cur_row, column=3).border = thin_border
    ws.cell(row=cur_row, column=4, value="-").alignment = center_align
    ws.cell(row=cur_row, column=4).border = thin_border

    ws.cell(row=cur_row, column=5, value=float(bank.total_general_pto or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=5).alignment = right_align
    ws.cell(row=cur_row, column=5).border = thin_border

    ws.cell(row=cur_row, column=6, value=float(bank.total_holiday_pto or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=6).alignment = right_align
    ws.cell(row=cur_row, column=6).border = thin_border

    ws.cell(row=cur_row, column=7, value=float(bank.total_sick_leave or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=7).alignment = right_align
    ws.cell(row=cur_row, column=7).border = thin_border

    ws.cell(row=cur_row, column=8, value=float(bank.total_general_leave or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=8).alignment = right_align
    ws.cell(row=cur_row, column=8).border = thin_border

    ws.cell(row=cur_row, column=9, value=float(bank.total_parental_leave or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=9).alignment = right_align
    ws.cell(row=cur_row, column=9).border = thin_border

    ws.cell(row=cur_row, column=10, value=float(bank.total_marriage_leave or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=10).alignment = right_align
    ws.cell(row=cur_row, column=10).border = thin_border

    ws.cell(row=cur_row, column=11, value=float(bank.total_adjustments or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=11).alignment = right_align
    ws.cell(row=cur_row, column=11).border = thin_border

    ws.cell(row=cur_row, column=12, value=float(bank.current_balance or 0.0)).font = bold_font
    ws.cell(row=cur_row, column=12).alignment = right_align
    ws.cell(row=cur_row, column=12).border = thin_border

    ws.cell(row=cur_row, column=13, value="").border = thin_border
    cur_row += 2

    # December 50% Carry Forward
    ws.merge_cells(start_row=cur_row, start_column=2, end_row=cur_row, end_column=8)
    cf_label = ws.cell(row=cur_row, column=2, value=f"Year {year} Closing Carry Forward to {year + 1} (50% of Dec Closing):")
    cf_label.font = bold_font
    cf_label.alignment = right_align
    cf_label.fill = closing_fill

    ws.merge_cells(start_row=cur_row, start_column=9, end_row=cur_row, end_column=12)
    cf_val = ws.cell(row=cur_row, column=9, value=f"{float(bank.carried_forward_to_next_year or 0.0):.2f} days")
    cf_val.font = Font(name="Calibri", size=11, bold=True, color="047857")
    cf_val.alignment = center_align
    cf_val.fill = closing_fill

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A', 'N']:
            ws.column_dimensions[col_letter].width = 4
        elif col_letter in ['B', 'C']:
            ws.column_dimensions[col_letter].width = 14
        elif col_letter == 'M':
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = 14

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def export_all_staff_leave_summary_excel(year: int, db_session, models, company_id=1):
    """
    Exports all employees' Leave Bank summary for a given year into an executive Excel sheet.
    """
    LeaveBank = models['LeaveBank']
    Employee = models['Employee']

    employees = Employee.query.filter_by(status='Active').order_by(Employee.employee_id.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Leave_Summary_{year}"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    navy_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    teal_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    accent_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    closing_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # Title
    ws.merge_cells("A1:L1")
    ws.cell(row=1, column=1, value=f"ALL STAFF LEAVE BANK & PTO SUMMARY - YEAR {year}").font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).alignment = center_align
    ws.cell(row=1, column=1).fill = navy_fill

    headers = [
        "SL", "ID", "Employee Name", "Department", "Designation",
        "Opening Bank", "General PTO (50%)", "Holiday PTO", "Total Earned",
        "Leaves Taken", "Current Balance", "50% Carry Forward"
    ]

    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = teal_fill
        cell.alignment = center_align
        cell.border = thin_border

    cur_row = 4
    existing_banks = LeaveBank.query.filter_by(year=year).all()
    bank_map = {b.employee_id: b for b in existing_banks}

    for sl, emp in enumerate(employees, start=1):
        bank = bank_map.get(emp.id)
        if not bank:
            bank = recalculate_employee_leave_ledger(emp.id, year, db_session, models)
        if not bank:
            continue

        tot_earned = (bank.total_general_pto or 0.0) + (bank.total_holiday_pto or 0.0)
        tot_taken = (
            (bank.total_sick_leave or 0.0) +
            (bank.total_general_leave or 0.0) +
            (bank.total_parental_leave or 0.0) +
            (bank.total_marriage_leave or 0.0) +
            (bank.total_other_leave or 0.0)
        )

        ws.cell(row=cur_row, column=1, value=sl).alignment = center_align
        ws.cell(row=cur_row, column=2, value=emp.employee_id).alignment = center_align
        ws.cell(row=cur_row, column=3, value=emp.name).alignment = left_align
        ws.cell(row=cur_row, column=4, value=emp.department or 'General').alignment = left_align
        ws.cell(row=cur_row, column=5, value=emp.designation or 'Staff').alignment = left_align
        ws.cell(row=cur_row, column=6, value=float(bank.opening_balance or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=7, value=float(bank.total_general_pto or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=8, value=float(bank.total_holiday_pto or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=9, value=round(tot_earned, 2)).alignment = right_align
        ws.cell(row=cur_row, column=10, value=round(tot_taken, 2)).alignment = right_align
        ws.cell(row=cur_row, column=11, value=float(bank.current_balance or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=11).font = bold_font
        ws.cell(row=cur_row, column=12, value=float(bank.carried_forward_to_next_year or 0.0)).alignment = right_align
        ws.cell(row=cur_row, column=12).font = bold_font

        for c_i in range(1, 13):
            ws.cell(row=cur_row, column=c_i).border = thin_border

        cur_row += 1

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A']:
            ws.column_dimensions[col_letter].width = 6
        elif col_letter in ['B']:
            ws.column_dimensions[col_letter].width = 12
        elif col_letter in ['C']:
            ws.column_dimensions[col_letter].width = 26
        elif col_letter in ['D', 'E']:
            ws.column_dimensions[col_letter].width = 20
        else:
            ws.column_dimensions[col_letter].width = 16

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
