import os
from io import BytesIO
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import request, redirect, render_template, url_for, flash, jsonify, send_file


def init_payroll_salary_routes(app, db, models):
    Employee = models['Employee']
    MonthlySalarySheet = models['MonthlySalarySheet']
    MonthlySalaryItem = models['MonthlySalaryItem']
    Account = models['Account']
    JournalEntry = models['JournalEntry']
    TransactionLine = models['TransactionLine']
    num_to_words = models['num_to_words']

    # -------------------------------------------------------------------------
    # Helper: Recalculate Sheet Totals (Considers only Running employees)
    # -------------------------------------------------------------------------
    def recalculate_sheet_totals(sheet):
        all_items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).all()
        # Only active / running items contribute to the financial totals
        running_items = [i for i in all_items if (i.status or 'Running') != 'Inactive']

        sheet.total_basic = round(sum(i.basic or 0.0 for i in running_items), 2)
        sheet.total_hr_allowance = round(sum(i.hr_allowance or 0.0 for i in running_items), 2)
        sheet.total_medical_allowance = round(sum(i.medical_allowance or 0.0 for i in running_items), 2)
        sheet.total_conveyance = round(sum(i.conveyance_allowance or 0.0 for i in running_items), 2)
        sheet.total_gross = round(sum(i.gross_salary or 0.0 for i in running_items), 2)

        sheet.total_pf_employer = round(sum(i.pf_employer or 0.0 for i in running_items), 2)
        sheet.total_pf_employee = round(sum(i.pf_employee or 0.0 for i in running_items), 2)
        sheet.total_pf = round(sum(i.total_pf or 0.0 for i in running_items), 2)
        sheet.total_salary = round(sum(i.total_salary or 0.0 for i in running_items), 2)

        sheet.total_holiday_allowance = round(sum((i.holiday_extra_duty or 0.0) + (i.govt_holiday_duty or 0.0) for i in running_items), 2)
        sheet.total_loyalty_allowance = round(sum(i.loyalty_allowance or 0.0 for i in running_items), 2)
        sheet.total_additional_allowance = round(sum(i.additional_allowance or 0.0 for i in running_items), 2)
        sheet.total_eid_bonus = round(sum(i.eid_bonus or 0.0 for i in running_items), 2)
        sheet.total_other_allowance = round(sum(i.other_allowance or 0.0 for i in running_items), 2)

        sheet.total_absenteeism_deduction = round(sum(i.absenteeism_amount or 0.0 for i in running_items), 2)
        sheet.total_late_deduction = round(sum(i.late_leave_amount or 0.0 for i in running_items), 2)
        sheet.total_loan_deduction = round(sum(i.loan_deduction or 0.0 for i in running_items), 2)
        sheet.total_other_deduction = round(sum(i.other_deduction or 0.0 for i in running_items), 2)
        sheet.total_arrear = round(sum(i.arrear or 0.0 for i in running_items), 2)

        sheet.total_net_pay = round(sum(i.net_pay or 0.0 for i in running_items), 2)
        sheet.total_tds = round(sum(i.tds_deduction or 0.0 for i in running_items), 2)
        sheet.total_net_after_tds = round(sum(i.net_after_tds or 0.0 for i in running_items), 2)

        eft_items = [i for i in running_items if i.payment_method == 'EFT']
        beftn_items = [i for i in running_items if i.payment_method == 'BEFTN']
        cash_items = [i for i in running_items if i.payment_method == 'Cash']

        sheet.total_eft_amount = round(sum(i.net_after_tds or 0.0 for i in eft_items), 2)
        sheet.total_eft_count = len(eft_items)
        sheet.total_beftn_amount = round(sum(i.net_after_tds or 0.0 for i in beftn_items), 2)
        sheet.total_beftn_count = len(beftn_items)
        sheet.total_cash_amount = round(sum(i.net_after_tds or 0.0 for i in cash_items), 2)
        sheet.total_cash_count = len(cash_items)

    # -------------------------------------------------------------------------
    # Route: List of Monthly Salary Sheets
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheets')
    def monthly_salary_sheets():
        sheets = MonthlySalarySheet.query.order_by(MonthlySalarySheet.salary_month.desc()).all()
        return render_template('payroll/salary_sheets_index.html', sheets=sheets)

    # -------------------------------------------------------------------------
    # Route: Generate or Regenerate Monthly Salary Sheet
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheets/generate', methods=['POST'])
    def generate_monthly_salary_sheet():
        salary_month = request.form.get('salary_month', '').strip()
        title = request.form.get('title', '').strip() or f"Monthly Salary Sheet for {salary_month}"
        try:
            working_days = int(request.form.get('working_days', 30))
        except ValueError:
            working_days = 30
        notes = request.form.get('notes', '').strip()
        overwrite = request.form.get('overwrite') == '1'

        if not salary_month:
            flash("Please provide a salary month in YYYY-MM format.", "danger")
            return redirect(url_for('monthly_salary_sheets'))

        existing = MonthlySalarySheet.query.filter_by(salary_month=salary_month).first()
        if existing and not overwrite:
            flash(f"A salary sheet for {salary_month} already exists. Click below to view or choose regenerate to overwrite.", "info")
            return redirect(url_for('view_monthly_salary_sheet', sheet_id=existing.id))

        if existing and overwrite:
            MonthlySalaryItem.query.filter_by(sheet_id=existing.id).delete()
            sheet = existing
            sheet.title = title
            sheet.working_days = working_days
            sheet.notes = notes
            sheet.processed_date = date.today()
            sheet.status = 'Draft'
        else:
            sheet = MonthlySalarySheet(
                salary_month=salary_month,
                title=title,
                working_days=working_days,
                notes=notes,
                processed_date=date.today(),
                status='Draft'
            )
            db.session.add(sheet)
            db.session.flush()

        try:
            parts = salary_month.split('-')
            target_date = date(int(parts[0]), int(parts[1]), 1)
        except Exception:
            target_date = date.today()

        employees = Employee.query.filter_by(status='Active').order_by(Employee.department.asc(), Employee.name.asc()).all()

        for emp in employees:
            is_part_time = (emp.employment_type == 'Part-Time')
            duty_cat = emp.duty_category

            if is_part_time:
                # Default rates setup based on category
                eve_rate = 800.0 if duty_cat == 'News Presenter' else 0.0
                night_rate = 1400.0 if duty_cat == 'News Presenter' else (500.0 if duty_cat == 'Customer Care Agent' else 0.0)
                morn_rate = 1600.0 if (duty_cat == 'News Presenter' and 'Presenter' in (emp.designation or '') and 'Junior' not in (emp.designation or '')) else (1500.0 if duty_cat == 'News Presenter' else (500.0 if duty_cat == 'Customer Care Agent' else 0.0))
                aft_rate = 700.0 if duty_cat == 'Customer Care Agent' else 0.0

                item = MonthlySalaryItem(
                    sheet_id=sheet.id,
                    employee_id=emp.id,
                    emp_code=emp.employee_id,
                    name=emp.name,
                    department=emp.department or 'General',
                    designation=emp.designation or 'Staff',
                    joining_date=emp.joining_date,
                    bank_name=emp.bank_name or 'The Premier Bank PLC',
                    bank_branch=emp.bank_branch or 'Banani Branch',
                    routing_no=emp.routing_no or '085261198',
                    bank_account_no=emp.bank_account_no,
                    payment_method=emp.payment_method or 'EFT',
                    employment_type='Part-Time',
                    duty_category=duty_cat,
                    shift_eve_duty=0.0,
                    shift_eve_rate=eve_rate,
                    shift_night_duty=0.0,
                    shift_night_rate=night_rate,
                    shift_morn_duty=0.0,
                    shift_morn_rate=morn_rate,
                    shift_aft_duty=0.0,
                    shift_aft_rate=aft_rate,
                    total_duty=0.0,
                    basic=0.0,
                    hr_allowance=0.0,
                    medical_allowance=0.0,
                    conveyance_allowance=0.0,
                    gross_salary=0.0,
                    pf_status='N',
                    pf_employer=0.0,
                    pf_employee=0.0,
                    total_pf=0.0,
                    total_salary=0.0,
                    govt_holiday_duty=0.0,
                    holiday_extra_duty=0.0,
                    loyalty_allowance=0.0,
                    additional_allowance=0.0,
                    eid_bonus=0.0,
                    other_allowance=0.0,
                    working_days=working_days,
                    absent_days=0,
                    late_days=0,
                    absenteeism_amount=0.0,
                    late_leave_amount=0.0,
                    loan_deduction=0.0,
                    other_deduction=0.0,
                    arrear=0.0,
                    net_pay=0.0,
                    tds_deduction=round(emp.tds_amount or 0.0, 2),
                    net_after_tds=round(0.0 - (emp.tds_amount or 0.0), 2)
                )
                db.session.add(item)
                continue

            basic = round(emp.basic_salary or 0.0, 2)
            hr = round(emp.house_rent if (emp.house_rent and emp.house_rent > 0) else basic * 0.50, 2)
            med = round(emp.medical_allowance if (emp.medical_allowance and emp.medical_allowance > 0) else 1480.0, 2)
            conv = round(emp.conveyance if (emp.conveyance and emp.conveyance > 0) else 2000.0, 2)
            gross = round(basic + hr + med + conv, 2)

            pf_st = emp.pf_status or 'P'
            if pf_st == 'P':
                pf_employer = round(basic * 0.025, 2)
                pf_employee = round(basic * 0.10, 2)
            else:
                pf_employer = 0.0
                pf_employee = 0.0
            tot_pf = round(pf_employer + pf_employee, 2)
            tot_sal = round(gross - pf_employee, 2)

            loyalty = 0.0
            if emp.joining_date:
                tenure_years = (target_date - emp.joining_date).days / 365.25
                if tenure_years >= 10.0:
                    loyalty = round(basic * 0.15, 2)
                elif tenure_years >= 5.0:
                    loyalty = round(basic * 0.075, 2)

            net_pay = round(tot_sal + loyalty, 2)
            tds = round(emp.tds_amount or 0.0, 2)
            net_after_tds = round(net_pay - tds, 2)

            item = MonthlySalaryItem(
                sheet_id=sheet.id,
                employee_id=emp.id,
                emp_code=emp.employee_id,
                name=emp.name,
                department=emp.department or 'General',
                designation=emp.designation or 'Staff',
                joining_date=emp.joining_date,
                bank_name=emp.bank_name or 'The Premier Bank PLC',
                bank_branch=emp.bank_branch or 'Banani Branch',
                routing_no=emp.routing_no or '085261198',
                bank_account_no=emp.bank_account_no,
                payment_method=emp.payment_method or 'EFT',
                basic=basic,
                hr_allowance=hr,
                medical_allowance=med,
                conveyance_allowance=conv,
                gross_salary=gross,
                pf_status=pf_st,
                pf_employer=pf_employer,
                pf_employee=pf_employee,
                total_pf=tot_pf,
                total_salary=tot_sal,
                govt_holiday_duty=0.0,
                holiday_extra_duty=0.0,
                loyalty_allowance=loyalty,
                additional_allowance=0.0,
                eid_bonus=0.0,
                other_allowance=0.0,
                working_days=working_days,
                absent_days=0,
                late_days=0,
                absenteeism_amount=0.0,
                late_leave_amount=0.0,
                loan_deduction=0.0,
                other_deduction=0.0,
                arrear=0.0,
                net_pay=net_pay,
                tds_deduction=tds,
                net_after_tds=net_after_tds
            )
            db.session.add(item)

        db.session.flush()
        recalculate_sheet_totals(sheet)
        db.session.commit()
        flash(f"Monthly Salary Sheet for {salary_month} generated successfully for {len(employees)} employees!", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: View Detailed 32-Column Monthly Salary Sheet
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>')
    def view_monthly_salary_sheet(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        all_items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).order_by(MonthlySalaryItem.department.asc(), MonthlySalaryItem.name.asc()).all()

        running_items = [i for i in all_items if (i.status or 'Running') != 'Inactive']
        inactive_items = [i for i in all_items if (i.status or 'Running') == 'Inactive']

        # Dedicated Part-Time Duty Lists
        news_presenter_items = [
            i for i in running_items 
            if i.duty_category == 'News Presenter' or (i.employment_type == 'Part-Time' and (i.department == 'News' or 'News' in (i.designation or '')))
        ]
        news_order_map = {'630': 1, '162': 2, 'IS': 3, '706': 4, '629': 5, 'N1': 6, '713': 7, '704': 8, '465': 9, '717': 10, 'Nab': 11, '716': 12}
        news_presenter_items.sort(key=lambda x: news_order_map.get(x.emp_code, 99))

        customer_care_items = [
            i for i in running_items 
            if i.duty_category == 'Customer Care Agent' or (i.employment_type == 'Part-Time' and ('Customer' in (i.department or '') or 'Agent' in (i.designation or '')))
        ]
        cc_order_map = {'457': 1, '647': 2, 'C2': 3, 'MM': 4, 'T4': 5, 'Mom': 6}
        customer_care_items.sort(key=lambda x: cc_order_map.get(x.emp_code, 99))

        news_totals = {
            'eve_duty': sum(i.shift_eve_duty or 0.0 for i in news_presenter_items),
            'night_duty': sum(i.shift_night_duty or 0.0 for i in news_presenter_items),
            'morn_duty': sum(i.shift_morn_duty or 0.0 for i in news_presenter_items),
            'total_duty': sum(i.total_duty or 0.0 for i in news_presenter_items),
            'salary': sum(i.gross_salary or 0.0 for i in news_presenter_items),
            'loyalty': sum(i.loyalty_allowance or 0.0 for i in news_presenter_items),
            'additional': sum(i.additional_allowance or 0.0 for i in news_presenter_items),
            'arrear': sum(i.arrear or 0.0 for i in news_presenter_items),
            'net_pay': sum(i.net_pay or 0.0 for i in news_presenter_items),
            'tds': sum(i.tds_deduction or 0.0 for i in news_presenter_items),
            'net_after_tds': sum(i.net_after_tds or 0.0 for i in news_presenter_items),
        }

        cc_totals = {
            'morn_duty': sum(i.shift_morn_duty or 0.0 for i in customer_care_items),
            'aft_duty': sum(i.shift_aft_duty or 0.0 for i in customer_care_items),
            'night_duty': sum(i.shift_night_duty or 0.0 for i in customer_care_items),
            'total_duty': sum(i.total_duty or 0.0 for i in customer_care_items),
            'salary': sum(i.gross_salary or 0.0 for i in customer_care_items),
            'arrear': sum(i.arrear or 0.0 for i in customer_care_items),
            'deductions': sum((i.absenteeism_amount or 0.0) + (i.late_leave_amount or 0.0) + (i.other_deduction or 0.0) for i in customer_care_items),
            'net_pay': sum(i.net_pay or 0.0 for i in customer_care_items),
            'tds': sum(i.tds_deduction or 0.0 for i in customer_care_items),
            'net_after_tds': sum(i.net_after_tds or 0.0 for i in customer_care_items),
        }

        dept_map = {}
        for it in running_items:
            dept = it.department or 'General'
            if dept not in dept_map:
                dept_map[dept] = {
                    'items': [],
                    'tot': {
                        'basic': 0.0, 'hr': 0.0, 'med': 0.0, 'conv': 0.0, 'gross': 0.0,
                        'pf_comp': 0.0, 'pf_emp': 0.0, 'total_pf': 0.0, 'total_sal': 0.0,
                        'govt_hol': 0.0, 'hol_extra': 0.0, 'loyalty': 0.0, 'addl': 0.0,
                        'bonus': 0.0, 'other': 0.0, 'absenteeism': 0.0, 'late_leave': 0.0,
                        'loan': 0.0, 'other_ded': 0.0, 'arrear': 0.0, 'net_pay': 0.0,
                        'tds': 0.0, 'net_after_tds': 0.0
                    }
                }
            g = dept_map[dept]
            g['items'].append(it)
            t = g['tot']
            t['basic'] = round(t['basic'] + (it.basic or 0.0), 2)
            t['hr'] = round(t['hr'] + (it.hr_allowance or 0.0), 2)
            t['med'] = round(t['med'] + (it.medical_allowance or 0.0), 2)
            t['conv'] = round(t['conv'] + (it.conveyance_allowance or 0.0), 2)
            t['gross'] = round(t['gross'] + (it.gross_salary or 0.0), 2)
            t['pf_comp'] = round(t['pf_comp'] + (it.pf_employer or 0.0), 2)
            t['pf_emp'] = round(t['pf_emp'] + (it.pf_employee or 0.0), 2)
            t['total_pf'] = round(t['total_pf'] + (it.total_pf or 0.0), 2)
            t['total_sal'] = round(t['total_sal'] + (it.total_salary or 0.0), 2)
            t['govt_hol'] = round(t['govt_hol'] + (it.govt_holiday_duty or 0.0), 2)
            t['hol_extra'] = round(t['hol_extra'] + (it.holiday_extra_duty or 0.0), 2)
            t['loyalty'] = round(t['loyalty'] + (it.loyalty_allowance or 0.0), 2)
            t['addl'] = round(t['addl'] + (it.additional_allowance or 0.0), 2)
            t['bonus'] = round(t['bonus'] + (it.eid_bonus or 0.0), 2)
            t['other'] = round(t['other'] + (it.other_allowance or 0.0), 2)
            t['absenteeism'] = round(t['absenteeism'] + (it.absenteeism_amount or 0.0), 2)
            t['late_leave'] = round(t['late_leave'] + (it.late_leave_amount or 0.0), 2)
            t['loan'] = round(t['loan'] + (it.loan_deduction or 0.0), 2)
            t['other_ded'] = round(t['other_ded'] + (it.other_deduction or 0.0), 2)
            t['arrear'] = round(t['arrear'] + (it.arrear or 0.0), 2)
            t['net_pay'] = round(t['net_pay'] + (it.net_pay or 0.0), 2)
            t['tds'] = round(t['tds'] + (it.tds_deduction or 0.0), 2)
            t['net_after_tds'] = round(t['net_after_tds'] + (it.net_after_tds or 0.0), 2)

        dept_groups = [(d_name, d_val['items'], d_val['tot']) for d_name, d_val in dept_map.items()]

        eft_items = [i for i in running_items if i.payment_method == 'EFT']
        beftn_items = [i for i in running_items if i.payment_method == 'BEFTN']
        cash_items = [i for i in running_items if i.payment_method == 'Cash']

        return render_template(
            'payroll/monthly_salary_sheet.html',
            sheet=sheet,
            items=running_items,
            running_items=running_items,
            inactive_items=inactive_items,
            all_items=all_items,
            news_presenter_items=news_presenter_items,
            customer_care_items=customer_care_items,
            news_totals=news_totals,
            cc_totals=cc_totals,
            eft_items=eft_items,
            beftn_items=beftn_items,
            cash_items=cash_items,
            dept_groups=dept_groups,
            running_count=len(running_items),
            inactive_count=len(inactive_items)
        )

    # -------------------------------------------------------------------------
    # Route: Update Single Salary Sheet Item
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/update-item/<int:item_id>', methods=['POST'])
    def update_salary_sheet_item(sheet_id, item_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        item = MonthlySalaryItem.query.filter_by(id=item_id, sheet_id=sheet.id).first_or_404()

        def parse_flt(key, default=0.0):
            try:
                return float(request.form.get(key, default) or 0.0)
            except (ValueError, TypeError):
                return default

        def parse_int(key, default=0):
            try:
                return int(request.form.get(key, default) or 0)
            except (ValueError, TypeError):
                return default

        item.basic = parse_flt('basic', item.basic)
        item.hr_allowance = parse_flt('hr_allowance', item.hr_allowance)
        item.medical_allowance = parse_flt('medical_allowance', item.medical_allowance)
        item.conveyance_allowance = parse_flt('conveyance_allowance', item.conveyance_allowance)
        item.gross_salary = round(item.basic + item.hr_allowance + item.medical_allowance + item.conveyance_allowance, 2)

        item.pf_status = request.form.get('pf_status', item.pf_status or 'P').strip()
        if item.pf_status == 'P':
            item.pf_employer = round(item.basic * 0.025, 2)
            item.pf_employee = round(item.basic * 0.10, 2)
        else:
            item.pf_employer = 0.0
            item.pf_employee = 0.0
        item.total_pf = round(item.pf_employer + item.pf_employee, 2)
        item.total_salary = round(item.gross_salary - item.pf_employee, 2)

        item.govt_holiday_duty = parse_flt('govt_holiday_duty', item.govt_holiday_duty)
        item.holiday_extra_duty = parse_flt('holiday_extra_duty', item.holiday_extra_duty)
        item.loyalty_allowance = parse_flt('loyalty_allowance', item.loyalty_allowance)
        item.additional_allowance = parse_flt('additional_allowance', item.additional_allowance)
        item.eid_bonus = parse_flt('eid_bonus', item.eid_bonus)
        item.other_allowance = parse_flt('other_allowance', item.other_allowance)

        item.working_days = parse_int('working_days', item.working_days)
        item.absent_days = parse_int('absent_days', item.absent_days)
        item.late_days = parse_int('late_days', item.late_days)
        item.absenteeism_amount = parse_flt('absenteeism_amount', item.absenteeism_amount)
        item.late_leave_amount = parse_flt('late_leave_amount', item.late_leave_amount)
        item.loan_deduction = parse_flt('loan_deduction', item.loan_deduction)
        item.other_deduction = parse_flt('other_deduction', item.other_deduction)
        item.arrear = parse_flt('arrear', item.arrear)

        tot_allowances = (
            item.govt_holiday_duty + item.holiday_extra_duty + item.loyalty_allowance +
            item.additional_allowance + item.eid_bonus + item.other_allowance + item.arrear
        )
        tot_deductions = (
            item.absenteeism_amount + item.late_leave_amount + item.loan_deduction + item.other_deduction
        )
        item.net_pay = round(item.total_salary + tot_allowances - tot_deductions, 2)
        item.tds_deduction = parse_flt('tds_deduction', item.tds_deduction)
        item.net_after_tds = round(item.net_pay - item.tds_deduction, 2)

        item.bank_name = request.form.get('bank_name', item.bank_name).strip()
        item.bank_branch = request.form.get('bank_branch', item.bank_branch).strip()
        item.routing_no = request.form.get('routing_no', item.routing_no).strip()
        item.bank_account_no = request.form.get('bank_account_no', item.bank_account_no).strip()
        item.payment_method = request.form.get('payment_method', item.payment_method)
        item.remarks = request.form.get('remarks', item.remarks or '').strip()

        recalculate_sheet_totals(sheet)
        db.session.commit()
        flash(f"Updated salary line for {item.name} ({item.emp_code}).", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Bulk Update Salary Sheet Items (Inline Spreadsheet Mode)
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/bulk-update', methods=['POST'])
    def bulk_update_salary_sheet_items(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        payload = request.get_json(silent=True)

        if payload and 'items' in payload:
            updates = payload['items']
            for u in updates:
                item_id = u.get('id')
                item = MonthlySalaryItem.query.filter_by(id=item_id, sheet_id=sheet.id).first()
                if not item:
                    continue
                for fld in ['basic', 'hr_allowance', 'medical_allowance', 'conveyance_allowance',
                            'govt_holiday_duty', 'holiday_extra_duty', 'loyalty_allowance',
                            'additional_allowance', 'eid_bonus', 'other_allowance', 'absenteeism_amount',
                            'late_leave_amount', 'loan_deduction', 'other_deduction', 'arrear', 'tds_deduction']:
                    if fld in u:
                        try:
                            setattr(item, fld, float(u[fld]))
                        except (ValueError, TypeError):
                            pass

                for fld in ['absent_days', 'late_days', 'working_days']:
                    if fld in u:
                        try:
                            setattr(item, fld, int(u[fld]))
                        except (ValueError, TypeError):
                            pass

                for fld in ['bank_name', 'bank_branch', 'routing_no', 'bank_account_no', 'payment_method', 'pf_status', 'remarks']:
                    if fld in u and u[fld] is not None:
                        setattr(item, fld, str(u[fld]).strip())

                item.gross_salary = round((item.basic or 0.0) + (item.hr_allowance or 0.0) + (item.medical_allowance or 0.0) + (item.conveyance_allowance or 0.0), 2)
                if item.pf_status == 'P':
                    item.pf_employer = round((item.basic or 0.0) * 0.025, 2)
                    item.pf_employee = round((item.basic or 0.0) * 0.10, 2)
                else:
                    item.pf_employer = 0.0
                    item.pf_employee = 0.0
                item.total_pf = round(item.pf_employer + item.pf_employee, 2)
                item.total_salary = round(item.gross_salary - item.pf_employee, 2)

                tot_allowances = (
                    (item.govt_holiday_duty or 0.0) + (item.holiday_extra_duty or 0.0) +
                    (item.loyalty_allowance or 0.0) + (item.additional_allowance or 0.0) +
                    (item.eid_bonus or 0.0) + (item.other_allowance or 0.0) + (item.arrear or 0.0)
                )
                tot_deductions = (
                    (item.absenteeism_amount or 0.0) + (item.late_leave_amount or 0.0) +
                    (item.loan_deduction or 0.0) + (item.other_deduction or 0.0)
                )
                item.net_pay = round(item.total_salary + tot_allowances - tot_deductions, 2)
                item.net_after_tds = round(item.net_pay - (item.tds_deduction or 0.0), 2)

            recalculate_sheet_totals(sheet)
            db.session.commit()
            return jsonify({'success': True, 'message': 'Salary sheet bulk updates applied successfully.'})

        # Fallback to form post
        flash("Bulk changes saved.", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Toggle Individual Employee Status (Running <-> Inactive)
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/toggle-status/<int:item_id>', methods=['POST'])
    def toggle_salary_sheet_item_status(sheet_id, item_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        item = MonthlySalaryItem.query.filter_by(id=item_id, sheet_id=sheet.id).first_or_404()

        requested_status = request.form.get('status')
        if not requested_status and request.is_json:
            data = request.get_json(silent=True) or {}
            requested_status = data.get('status')

        if requested_status in ['Running', 'Inactive']:
            item.status = requested_status
        else:
            item.status = 'Inactive' if (item.status or 'Running') == 'Running' else 'Running'

        # Also keep master employee database status aligned
        if item.employee:
            item.employee.status = 'Active' if item.status == 'Running' else 'Inactive'

        recalculate_sheet_totals(sheet)
        db.session.commit()

        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
            all_sheet_items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).all()
            run_items = [i for i in all_sheet_items if (i.status or 'Running') != 'Inactive']
            inact_items = [i for i in all_sheet_items if (i.status or 'Running') == 'Inactive']
            return jsonify({
                'success': True,
                'item_id': item.id,
                'status': item.status,
                'emp_name': item.name,
                'emp_code': item.emp_code,
                'running_count': len(run_items),
                'inactive_count': len(inact_items),
                'total_gross': f"{sheet.total_gross:,.0f}",
                'total_net_after_tds': f"{sheet.total_net_after_tds:,.2f}",
                'total_eft_amount': f"{sheet.total_eft_amount:,.2f}",
                'total_eft_count': sheet.total_eft_count,
                'total_beftn_amount': f"{sheet.total_beftn_amount:,.2f}",
                'total_beftn_count': sheet.total_beftn_count,
                'total_cash_amount': f"{sheet.total_cash_amount:,.2f}",
                'total_cash_count': sheet.total_cash_count,
                'message': f"{item.name} ({item.emp_code}) is now marked as {item.status}."
            })

        flash(f"{item.name} ({item.emp_code}) marked as {item.status}.", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Bulk Employee Status Update
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/bulk-status', methods=['POST'])
    def bulk_salary_sheet_status(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        target_status = request.form.get('target_status', 'Inactive')
        item_ids = request.form.getlist('item_ids')
        if not item_ids and request.is_json:
            data = request.get_json(silent=True) or {}
            target_status = data.get('target_status', 'Inactive')
            item_ids = data.get('item_ids', [])

        count = 0
        for i_id in item_ids:
            try:
                iid = int(i_id)
                item = MonthlySalaryItem.query.filter_by(id=iid, sheet_id=sheet.id).first()
                if item:
                    item.status = target_status
                    if item.employee:
                        item.employee.status = 'Active' if target_status == 'Running' else 'Inactive'
                    count += 1
            except (ValueError, TypeError):
                pass

        recalculate_sheet_totals(sheet)
        db.session.commit()

        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
            all_sheet_items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).all()
            run_items = [i for i in all_sheet_items if (i.status or 'Running') != 'Inactive']
            inact_items = [i for i in all_sheet_items if (i.status or 'Running') == 'Inactive']
            return jsonify({
                'success': True,
                'count': count,
                'target_status': target_status,
                'running_count': len(run_items),
                'inactive_count': len(inact_items),
                'total_gross': f"{sheet.total_gross:,.0f}",
                'total_net_after_tds': f"{sheet.total_net_after_tds:,.2f}"
            })

        flash(f"{count} employee(s) moved to {target_status}.", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Update Single Part-Time Per-Duty Salary Item
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/update-part-time-item/<int:item_id>', methods=['POST'])
    def update_part_time_salary_item(sheet_id, item_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        item = MonthlySalaryItem.query.filter_by(id=item_id, sheet_id=sheet.id).first_or_404()

        def parse_flt(key, default=0.0):
            try:
                data = request.get_json(silent=True) if request.is_json else request.form
                if data and key in data and data[key] is not None and str(data[key]).strip() != '':
                    return float(data[key])
                return default
            except (ValueError, TypeError, AttributeError):
                return default

        req_cat = None
        if request.is_json:
            req_data = request.get_json(silent=True) or {}
            req_cat = req_data.get('duty_category')
        else:
            req_cat = request.form.get('duty_category')
        cat = item.duty_category or req_cat
        
        item.shift_eve_duty = parse_flt('shift_eve_duty', item.shift_eve_duty or 0.0)
        item.shift_eve_rate = parse_flt('shift_eve_rate', item.shift_eve_rate or 800.0)
        item.shift_night_duty = parse_flt('shift_night_duty', item.shift_night_duty or 0.0)
        item.shift_night_rate = parse_flt('shift_night_rate', item.shift_night_rate or (1400.0 if cat == 'News Presenter' else 500.0))
        item.shift_morn_duty = parse_flt('shift_morn_duty', item.shift_morn_duty or 0.0)
        item.shift_morn_rate = parse_flt('shift_morn_rate', item.shift_morn_rate or (1500.0 if cat == 'News Presenter' else 500.0))
        item.shift_aft_duty = parse_flt('shift_aft_duty', item.shift_aft_duty or 0.0)
        item.shift_aft_rate = parse_flt('shift_aft_rate', item.shift_aft_rate or 700.0)
        
        item.loyalty_allowance = parse_flt('loyalty_allowance', item.loyalty_allowance or 0.0)
        item.additional_allowance = parse_flt('additional_allowance', item.additional_allowance or 0.0)
        item.arrear = parse_flt('arrear', item.arrear or 0.0)
        item.absenteeism_amount = parse_flt('absenteeism_amount', item.absenteeism_amount or 0.0)
        item.late_leave_amount = parse_flt('late_leave_amount', item.late_leave_amount or 0.0)
        item.other_deduction = parse_flt('other_deduction', item.other_deduction or 0.0)
        item.tds_deduction = parse_flt('tds_deduction', item.tds_deduction or 0.0)

        # Duty and salary calculations
        if cat == 'News Presenter' or (item.shift_aft_duty == 0 and item.shift_eve_duty > 0):
            item.total_duty = round(item.shift_eve_duty + item.shift_night_duty + item.shift_morn_duty, 2)
            duty_sal = round((item.shift_eve_duty * item.shift_eve_rate) + (item.shift_night_duty * item.shift_night_rate) + (item.shift_morn_duty * item.shift_morn_rate), 2)
            item.gross_salary = duty_sal
            item.total_salary = duty_sal
            item.net_pay = round(duty_sal + item.loyalty_allowance + item.additional_allowance + item.arrear, 2)
            item.net_after_tds = round(item.net_pay - item.tds_deduction, 2)
        else: # Customer Care Agent
            item.total_duty = round(item.shift_morn_duty + item.shift_aft_duty + item.shift_night_duty, 2)
            duty_sal = round((item.shift_morn_duty * item.shift_morn_rate) + (item.shift_aft_duty * item.shift_aft_rate) + (item.shift_night_duty * item.shift_night_rate), 2)
            item.gross_salary = duty_sal
            item.total_salary = duty_sal
            deductions = item.absenteeism_amount + item.late_leave_amount + item.other_deduction
            item.net_pay = round(duty_sal + item.arrear - deductions, 2)
            item.net_after_tds = round(item.net_pay - item.tds_deduction, 2)

        recalculate_sheet_totals(sheet)
        db.session.commit()

        if request.is_json or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({
                'success': True,
                'item_id': item.id,
                'total_duty': item.total_duty,
                'gross_salary': item.gross_salary,
                'net_pay': item.net_pay,
                'net_after_tds': item.net_after_tds,
                'message': f"Updated {item.name}: {item.total_duty} duties -> ৳{item.net_after_tds:,.2f}"
            })
        flash(f"Updated duty calculations for {item.name}.", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Bulk Update Part-Time Per-Duty Shifts
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/bulk-update-part-time', methods=['POST'])
    def bulk_update_part_time(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        category = request.form.get('category') # 'News Presenter' or 'Customer Care Agent'
        
        item_ids = request.form.getlist('item_id[]')
        for iid in item_ids:
            try:
                item = MonthlySalaryItem.query.filter_by(id=int(iid), sheet_id=sheet.id).first()
                if not item:
                    continue
                
                cat = item.duty_category or category
                if cat == 'News Presenter':
                    item.shift_eve_duty = float(request.form.get(f'eve_{iid}', item.shift_eve_duty or 0.0) or 0.0)
                    item.shift_night_duty = float(request.form.get(f'night_{iid}', item.shift_night_duty or 0.0) or 0.0)
                    item.shift_morn_duty = float(request.form.get(f'morn_{iid}', item.shift_morn_duty or 0.0) or 0.0)
                    item.additional_allowance = float(request.form.get(f'addl_{iid}', item.additional_allowance or 0.0) or 0.0)
                    item.arrear = float(request.form.get(f'arrear_{iid}', item.arrear or 0.0) or 0.0)
                    
                    item.total_duty = round(item.shift_eve_duty + item.shift_night_duty + item.shift_morn_duty, 2)
                    duty_sal = round((item.shift_eve_duty * (item.shift_eve_rate or 800.0)) + 
                                     (item.shift_night_duty * (item.shift_night_rate or 1400.0)) + 
                                     (item.shift_morn_duty * (item.shift_morn_rate or 1500.0)), 2)
                    item.gross_salary = duty_sal
                    item.total_salary = duty_sal
                    item.net_pay = round(duty_sal + (item.loyalty_allowance or 0.0) + item.additional_allowance + item.arrear, 2)
                    item.net_after_tds = round(item.net_pay - (item.tds_deduction or 0.0), 2)
                else: # Customer Care Agent
                    item.shift_morn_duty = float(request.form.get(f'cc_morn_{iid}', item.shift_morn_duty or 0.0) or 0.0)
                    item.shift_aft_duty = float(request.form.get(f'cc_aft_{iid}', item.shift_aft_duty or 0.0) or 0.0)
                    item.shift_night_duty = float(request.form.get(f'cc_night_{iid}', item.shift_night_duty or 0.0) or 0.0)
                    item.arrear = float(request.form.get(f'cc_arrear_{iid}', item.arrear or 0.0) or 0.0)
                    
                    item.total_duty = round(item.shift_morn_duty + item.shift_aft_duty + item.shift_night_duty, 2)
                    duty_sal = round((item.shift_morn_duty * (item.shift_morn_rate or 500.0)) + 
                                     (item.shift_aft_duty * (item.shift_aft_rate or 700.0)) + 
                                     (item.shift_night_duty * (item.shift_night_rate or 500.0)), 2)
                    item.gross_salary = duty_sal
                    item.total_salary = duty_sal
                    deductions = (item.absenteeism_amount or 0.0) + (item.late_leave_amount or 0.0) + (item.other_deduction or 0.0)
                    item.net_pay = round(duty_sal + item.arrear - deductions, 2)
                    item.net_after_tds = round(item.net_pay - (item.tds_deduction or 0.0), 2)
            except Exception as e:
                print(f"Error updating item {iid}: {e}")

        recalculate_sheet_totals(sheet)
        db.session.commit()
        flash(f"Successfully saved all duty calculations for {category or 'Part-Time staff'}!", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Export Part-Time (News Presenter & Customer Care) Excel Workbook
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/export-part-time-excel')
    def export_part_time_excel(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        
        all_items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).all()
        running_items = [i for i in all_items if (i.status or 'Running') != 'Inactive']
        
        news_items = [
            i for i in running_items 
            if i.duty_category == 'News Presenter' or (i.employment_type == 'Part-Time' and (i.department == 'News' or 'News' in (i.designation or '')))
        ]
        news_order_map = {'630': 1, '162': 2, 'IS': 3, '706': 4, '629': 5, 'N1': 6, '713': 7, '704': 8, '465': 9, '717': 10, 'Nab': 11, '716': 12}
        news_items.sort(key=lambda x: news_order_map.get(x.emp_code, 99))

        cc_items = [
            i for i in running_items 
            if i.duty_category == 'Customer Care Agent' or (i.employment_type == 'Part-Time' and ('Customer' in (i.department or '') or 'Agent' in (i.designation or '')))
        ]
        cc_order_map = {'457': 1, '647': 2, 'C2': 3, 'MM': 4, 'T4': 5, 'Mom': 6}
        cc_items.sort(key=lambda x: cc_order_map.get(x.emp_code, 99))

        wb = openpyxl.Workbook()
        
        # --- TAB 1: News Presenters ---
        ws1 = wb.active
        ws1.title = "News Presenters"
        ws1.views.sheetView[0].showGridLines = True

        news_hdr_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        news_hdr_font = Font(name="Calibri", size=9, bold=True, color="000000")
        news_tot_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        bold_font = Font(name="Calibri", size=9, bold=True)
        norm_font = Font(name="Calibri", size=9)

        ws1.cell(row=1, column=1, value="ZR-ERP NEWS PRESENTERS DUTY & SALARY STATEMENT").font = Font(name="Calibri", size=13, bold=True, color="C2410C")
        ws1.cell(row=2, column=1, value=f"Salary Month: {sheet.salary_month} | Confidential Duty Breakdown").font = Font(name="Calibri", size=10, italic=True)

        headers1 = [
            "SL No", "Name", "Department", "Designation", "Date of Joining",
            "Evening Shift (6pm to 11pm) @800", "Night Shift (10 Pm -6am) @ 1200-1400",
            "Morning Shift (6-11.30) @ 1300-1500", "Total Duty", "Salary",
            "Loyalty Allowance 7.5%", "5Yrs Additional allowance as per office",
            "Update", "Arrear", "Net Pay", "TDS Deduction", "Net pay after TDS",
            "Payment Method", "Bank Name", "Account No"
        ]
        
        hdr_row1 = 4
        ws1.row_dimensions[hdr_row1].height = 28
        for col_idx, h in enumerate(headers1, 1):
            c = ws1.cell(row=hdr_row1, column=col_idx, value=h)
            c.font = news_hdr_font
            c.fill = news_hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border

        r1 = hdr_row1 + 1
        for it in news_items:
            ws1.row_dimensions[r1].height = 20
            row_vals = [
                it.emp_code,
                it.name,
                it.department,
                it.designation,
                it.joining_date.strftime('%d/%m/%Y') if it.joining_date else '',
                it.shift_eve_duty or 0,
                it.shift_night_duty or 0,
                it.shift_morn_duty or 0,
                it.total_duty or 0,
                it.gross_salary or 0.0,
                it.loyalty_allowance or 0.0,
                it.additional_allowance or 0.0,
                0,
                it.arrear or 0.0,
                it.net_pay or 0.0,
                it.tds_deduction or 0.0,
                it.net_after_tds or 0.0,
                it.payment_method,
                it.bank_name,
                it.bank_account_no
            ]
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws1.cell(row=r1, column=col_idx, value=val)
                cell.font = norm_font
                cell.border = thin_border
                if col_idx in [6, 7, 8, 9, 13]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if isinstance(val, (int, float)) and val > 0:
                        cell.number_format = '#,##0'
                elif col_idx in [10, 11, 12, 14, 15, 16, 17]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r1 += 1

        ws1.row_dimensions[r1].height = 22
        ws1.cell(row=r1, column=1, value="Total").font = bold_font
        for c in range(1, len(headers1) + 1):
            cell = ws1.cell(row=r1, column=c)
            cell.fill = news_tot_fill
            cell.border = thin_border
            cell.font = bold_font
            if c in [6, 7, 8, 9]:
                cl = get_column_letter(c)
                cell.value = f"=SUM({cl}{hdr_row1 + 1}:{cl}{r1 - 1})"
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [10, 11, 12, 14, 15, 16, 17]:
                cl = get_column_letter(c)
                cell.value = f"=SUM({cl}{hdr_row1 + 1}:{cl}{r1 - 1})"
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws1.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws1.column_dimensions[col_letter].width = max(max_len + 3, 11)

        # --- TAB 2: Customer Care Agents ---
        ws2 = wb.create_sheet(title="Customer Care")
        ws2.views.sheetView[0].showGridLines = True

        cc_hdr_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        cc_hdr_font = Font(name="Calibri", size=9, bold=True, color="000000")
        cc_tot_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

        ws2.cell(row=1, column=1, value="ZR-ERP CUSTOMER CARE AGENTS DUTY & SALARY STATEMENT").font = Font(name="Calibri", size=13, bold=True, color="15803D")
        ws2.cell(row=2, column=1, value=f"Salary Month: {sheet.salary_month} | Confidential Duty Breakdown").font = Font(name="Calibri", size=10, italic=True)

        headers2 = [
            "SL No / Part time", "Name", "Department", "Designation",
            "Morning (500 BDT)", "Afternoon (700 BDT)", "Night (500 BDT)",
            "Total Duty", "Gross Salary / Total", "Arrear",
            "Absenteeism Amount", "Late/Leave Amount", "Other Deduction",
            "Net Pay", "TDS Deduction", "Net pay after TDS",
            "Payment Method", "Bank Name", "Account No"
        ]

        hdr_row2 = 4
        ws2.row_dimensions[hdr_row2].height = 28
        for col_idx, h in enumerate(headers2, 1):
            c = ws2.cell(row=hdr_row2, column=col_idx, value=h)
            c.font = cc_hdr_font
            c.fill = cc_hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border

        r2 = hdr_row2 + 1
        for it in cc_items:
            ws2.row_dimensions[r2].height = 20
            row_vals = [
                it.emp_code,
                it.name,
                it.department,
                it.designation,
                it.shift_morn_duty or 0,
                it.shift_aft_duty or 0,
                it.shift_night_duty or 0,
                it.total_duty or 0,
                it.gross_salary or 0.0,
                it.arrear or 0.0,
                it.absenteeism_amount or 0.0,
                it.late_leave_amount or 0.0,
                it.other_deduction or 0.0,
                it.net_pay or 0.0,
                it.tds_deduction or 0.0,
                it.net_after_tds or 0.0,
                it.payment_method,
                it.bank_name,
                it.bank_account_no
            ]
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws2.cell(row=r2, column=col_idx, value=val)
                cell.font = norm_font
                cell.border = thin_border
                if col_idx in [5, 6, 7, 8]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if isinstance(val, (int, float)) and val > 0:
                        cell.number_format = '#,##0'
                elif col_idx in [9, 10, 11, 12, 13, 14, 15, 16]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r2 += 1

        ws2.row_dimensions[r2].height = 22
        ws2.cell(row=r2, column=1, value="Total").font = bold_font
        for c in range(1, len(headers2) + 1):
            cell = ws2.cell(row=r2, column=c)
            cell.fill = cc_tot_fill
            cell.border = thin_border
            cell.font = bold_font
            if c in [5, 6, 7, 8]:
                cl = get_column_letter(c)
                cell.value = f"=SUM({cl}{hdr_row2 + 1}:{cl}{r2 - 1})"
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [9, 10, 11, 12, 13, 14, 15, 16]:
                cl = get_column_letter(c)
                cell.value = f"=SUM({cl}{hdr_row2 + 1}:{cl}{r2 - 1})"
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws2.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws2.column_dimensions[col_letter].width = max(max_len + 3, 11)

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"Part_Time_Duty_Salary_Sheet_{sheet.salary_month}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    @app.route('/payroll/salary-sheet/<int:sheet_id>/delete', methods=['POST'])
    def delete_salary_sheet(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        if sheet.status == 'Disbursed':
            flash("Cannot delete a disbursed salary sheet. Please reverse the General Ledger voucher first.", "danger")
            return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

        month_label = sheet.salary_month
        db.session.delete(sheet)
        db.session.commit()
        flash(f"Monthly Salary Sheet for {month_label} has been deleted.", "info")
        return redirect(url_for('monthly_salary_sheets'))

    # -------------------------------------------------------------------------
    # Route: Disburse Salary to General Ledger (Payment Voucher)
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/disburse-gl', methods=['POST'])
    def disburse_salary_sheet_gl(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        if sheet.status == 'Disbursed':
            flash("This salary sheet has already been disbursed to General Ledger.", "warning")
            return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

        # Debit Account: Salary Expense (IS-000)
        salary_expense_acc = Account.query.filter(Account.code == 'IS-000').first() or Account.query.filter(Account.name.ilike('%salary%')).first()
        # Credit Account: Premier Bank (BS-002)
        bank_acc = Account.query.filter(Account.code == 'BS-002').first() or Account.query.filter(Account.name.ilike('%premier%')).first()

        if not salary_expense_acc or not bank_acc:
            flash("Error: Missing GL accounts. Ensure 'IS-000 Salary' and 'BS-002 Premier Bank' exist in Chart of Accounts.", "danger")
            return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

        voucher_no = f"PV-SAL-{sheet.salary_month}"
        # Check if voucher already exists
        existing_v = JournalEntry.query.filter_by(voucher_no=voucher_no).first()
        if existing_v:
            voucher_no = f"PV-SAL-{sheet.salary_month}-{datetime.now().strftime('%H%M%S')}"

        total_disbursed = sheet.total_net_after_tds
        narration = f"Disbursement of Monthly Salary & Allowances for {sheet.salary_month} ({sheet.total_eft_count} EFT, {sheet.total_beftn_count} BEFTN, {sheet.total_cash_count} Cash)"

        je = JournalEntry(
            voucher_no=voucher_no,
            voucher_type='Payment',
            date=date.today(),
            reference=f"SAL-SHEET-{sheet.salary_month}",
            narration=narration,
            total_amount=total_disbursed
        )
        db.session.add(je)
        db.session.flush()

        # DR: Salary Expense
        dr_line = TransactionLine(
            journal_entry_id=je.id,
            account_id=salary_expense_acc.id,
            debit=total_disbursed,
            credit=0.0,
            particulars=f"Salary & Allowances for {sheet.salary_month}"
        )
        # CR: Bank Account
        cr_line = TransactionLine(
            journal_entry_id=je.id,
            account_id=bank_acc.id,
            debit=0.0,
            credit=total_disbursed,
            particulars=f"Corporate Bank Transfer (Premier Bank A/C 02071110000026)"
        )
        db.session.add_all([dr_line, cr_line])

        # Update balances
        salary_expense_acc.balance = (salary_expense_acc.balance or 0.0) + total_disbursed
        bank_acc.balance = (bank_acc.balance or 0.0) - total_disbursed

        sheet.status = 'Disbursed'
        sheet.voucher_id = je.id
        sheet.payment_account_id = bank_acc.id
        db.session.commit()

        flash(f"Successfully posted Payment Voucher {voucher_no} to General Ledger for BDT {total_disbursed:,.2f}!", "success")
        return redirect(url_for('view_monthly_salary_sheet', sheet_id=sheet.id))

    # -------------------------------------------------------------------------
    # Route: Printable Bank Advice Forwarding Letter
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/bank-advice-letter/<channel>')
    def view_salary_bank_advice_letter(sheet_id, channel='eft'):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        channel = channel.lower()

        items_query = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).filter(MonthlySalaryItem.status != 'Inactive')
        if channel == 'eft':
            items = items_query.filter_by(payment_method='EFT').order_by(MonthlySalaryItem.department.asc(), MonthlySalaryItem.name.asc()).all()
            letter_title = "The Premier Bank PLC - Internal EFT Salary Advice Letter"
        elif channel == 'beftn':
            items = items_query.filter_by(payment_method='BEFTN').order_by(MonthlySalaryItem.department.asc(), MonthlySalaryItem.name.asc()).all()
            letter_title = "Bangladesh Bank BEFTN Clearing Salary Advice Letter"
        else:
            items = items_query.order_by(MonthlySalaryItem.department.asc(), MonthlySalaryItem.name.asc()).all()
            letter_title = "All Channels Salary Disbursement Advice Letter"

        total_net = sum(i.net_after_tds or 0.0 for i in items)
        amount_words = num_to_words(total_net)

        return render_template(
            'payroll/salary_bank_advice_letter.html',
            sheet=sheet,
            items=items,
            channel=channel,
            letter_title=letter_title,
            total_net=total_net,
            amount_in_words=amount_words,
            today_date=date.today()
        )

    # -------------------------------------------------------------------------
    # Route: Export Full Master Salary Sheet to Excel (32 Columns)
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/export-master-excel')
    def export_master_salary_sheet_excel(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id).filter(MonthlySalaryItem.status != 'Inactive').order_by(MonthlySalaryItem.department.asc(), MonthlySalaryItem.name.asc()).all()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Salary Sheet - {sheet.salary_month}"
        ws.views.sheetView[0].showGridLines = True

        title_font = Font(name="Calibri", size=15, bold=True, color="1E3A8A")
        sub_font = Font(name="Calibri", size=10, bold=False, color="475569")
        dept_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
        dept_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        subtot_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        grand_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        bold_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        norm_font = Font(name="Calibri", size=9, bold=False, color="000000")
        hdr_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        hdr_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")

        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        headers = [
            "ID No.", "Name", "Department", "Designation", "Date of Joining",
            "Basic", "HR Allowance", "Medical Allowance", "Conveyance Allowance", "Gross Salary",
            "PF", "PF Employer", "PF Employee", "Total PF", "Total Salary",
            "Govt. Holiday", "Holiday & Extra Duty", "Loyalty Allowance 7.5% & 15%", "Additional Allowance",
            "Eid Bonus", "Other", "Working Day", "Absent", "Late",
            "Absenteeism Amount", "Late/Leave Amount", "Loan Deduction", "Other Deduction",
            "Arrear", "Net Pay", "TDS Deduction", "Net Pay after TDS",
            "Bank Name", "Branch", "Routing No", "Account No", "Payment Method"
        ]

        ws.cell(row=1, column=1, value="ZR-ERP ENTERPRISE LIMITED").font = title_font
        ws.cell(row=2, column=1, value=f"{sheet.title} (Month: {sheet.salary_month}) | Total Staff: {len(items)}").font = sub_font
        ws.cell(row=3, column=1, value=f"Generated on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')} | Confidential Payroll Document").font = sub_font

        header_row = 5
        ws.row_dimensions[header_row].height = 28
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        cur_row = header_row + 1
        current_dept = None
        dept_start_row = None

        numeric_cols = [6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 25, 26, 27, 28, 29, 30, 31, 32]

        for item in items:
            dept = item.department or 'General'
            if dept != current_dept:
                if current_dept is not None and dept_start_row is not None:
                    # Write Sub-Total for previous department
                    ws.row_dimensions[cur_row].height = 20
                    ws.cell(row=cur_row, column=1, value=f"Sub-Total ({current_dept}):").font = bold_font
                    for c in range(1, len(headers) + 1):
                        cell = ws.cell(row=cur_row, column=c)
                        cell.fill = subtot_fill
                        cell.border = thin_border
                        if c in numeric_cols:
                            cl = get_column_letter(c)
                            cell.value = f"=SUM({cl}{dept_start_row}:{cl}{cur_row - 1})"
                            cell.number_format = '#,##0.00'
                            cell.font = bold_font
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                    cur_row += 1

                current_dept = dept
                dept_start_row = cur_row + 1
                # Department banner
                ws.row_dimensions[cur_row].height = 22
                dcell = ws.cell(row=cur_row, column=1, value=f"Department: {current_dept}")
                dcell.font = dept_font
                for c in range(1, len(headers) + 1):
                    ws.cell(row=cur_row, column=c).fill = dept_fill
                    ws.cell(row=cur_row, column=c).border = thin_border
                cur_row += 1

            ws.row_dimensions[cur_row].height = 19
            row_data = [
                item.emp_code,
                item.name,
                item.department,
                item.designation,
                item.joining_date.strftime('%d/%m/%Y') if item.joining_date else '',
                item.basic,
                item.hr_allowance,
                item.medical_allowance,
                item.conveyance_allowance,
                item.gross_salary,
                item.pf_status,
                item.pf_employer,
                item.pf_employee,
                item.total_pf,
                item.total_salary,
                item.govt_holiday_duty,
                item.holiday_extra_duty,
                item.loyalty_allowance,
                item.additional_allowance,
                item.eid_bonus,
                item.other_allowance,
                item.working_days,
                item.absent_days,
                item.late_days,
                item.absenteeism_amount,
                item.late_leave_amount,
                item.loan_deduction,
                item.other_deduction,
                item.arrear,
                item.net_pay,
                item.tds_deduction,
                item.net_after_tds,
                item.bank_name,
                item.bank_branch or '',
                item.routing_no or '',
                item.bank_account_no or '',
                item.payment_method
            ]

            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=cur_row, column=c_idx, value=val)
                cell.font = norm_font
                cell.border = thin_border
                if c_idx in numeric_cols:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c_idx in [1, 5, 11, 22, 23, 24, 35, 36, 37]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            cur_row += 1

        # Last dept subtotal
        if current_dept is not None and dept_start_row is not None:
            ws.row_dimensions[cur_row].height = 20
            ws.cell(row=cur_row, column=1, value=f"Sub-Total ({current_dept}):").font = bold_font
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=cur_row, column=c)
                cell.fill = subtot_fill
                cell.border = thin_border
                if c in numeric_cols:
                    cl = get_column_letter(c)
                    cell.value = f"=SUM({cl}{dept_start_row}:{cl}{cur_row - 1})"
                    cell.number_format = '#,##0.00'
                    cell.font = bold_font
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            cur_row += 1

        # Grand Total Row
        grand_row = cur_row
        ws.row_dimensions[grand_row].height = 24
        ws.cell(row=grand_row, column=1, value="GRAND TOTAL:").font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
        ws.cell(row=grand_row, column=1).alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=grand_row, column=c)
            cell.fill = grand_fill
            cell.border = thin_border
            if c in numeric_cols:
                cl = get_column_letter(c)
                cell.value = f"=SUMIF(A{header_row+1}:A{grand_row-1}, \"Sub-Total*\", {cl}{header_row+1}:{cl}{grand_row-1})"
                cell.number_format = '#,##0.00'
                cell.font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
                cell.alignment = Alignment(horizontal="right", vertical="center")

        # Auto-adjust column widths
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(len(str(col[header_row - 1].value or '')) + 4, 12)

        # Signature Block
        sig_start = grand_row + 4
        ws.cell(row=sig_start, column=2, value="Prepared By").font = bold_font
        ws.cell(row=sig_start, column=8, value="Checked By").font = bold_font
        ws.cell(row=sig_start, column=15, value="Manager - Accounts & Finance").font = bold_font
        ws.cell(row=sig_start, column=26, value="Director / CEO").font = bold_font

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"Master_Salary_Sheet_{sheet.salary_month}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # -------------------------------------------------------------------------
    # Route: Export Premier Bank EFT Excel Sheet
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/export-eft-excel')
    def export_premier_bank_eft_excel(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id, payment_method='EFT').filter(MonthlySalaryItem.status != 'Inactive').order_by(MonthlySalaryItem.name.asc()).all()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Premier Bank EFT"

        title_font = Font(name="Calibri", size=13, bold=True, color="1E3A8A")
        hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        bold_font = Font(name="Calibri", size=10, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        ws.cell(row=1, column=1, value="THE PREMIER BANK PLC - INTERNAL EFT SALARY DISBURSEMENT SCHEDULE").font = title_font
        ws.cell(row=2, column=1, value=f"Debit Account: 02071110000026 (ZR-ERP ENTERPRISE) | Banani Branch | Month: {sheet.salary_month}")
        ws.cell(row=3, column=1, value=f"Total Beneficiaries: {len(items)} | Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}")

        headers = [
            "SL", "Debit Account No", "Company Title", "Beneficiary ID",
            "Beneficiary Name", "Beneficiary Bank", "Beneficiary Branch",
            "Beneficiary Account No", "Amount (BDT)", "Narration"
        ]

        hrow = 5
        ws.row_dimensions[hrow].height = 25
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=hrow, column=c_idx, value=h)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        cur = hrow + 1
        for idx, it in enumerate(items, 1):
            ws.row_dimensions[cur].height = 19
            row_data = [
                idx,
                "02071110000026",
                "ZR-ERP ENTERPRISE",
                it.emp_code,
                it.name,
                it.bank_name,
                it.bank_branch or 'Banani Branch',
                it.bank_account_no or '',
                it.net_after_tds,
                f"Salary {sheet.salary_month}"
            ]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=cur, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx == 9:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.font = bold_font
                elif c_idx in [1, 2, 4, 8]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            cur += 1

        tot_row = cur
        ws.cell(row=tot_row, column=1, value="TOTAL EFT DISBURSEMENT (BDT):").font = bold_font
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=tot_row, column=c)
            cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
            cell.border = thin_border
            if c == 9:
                cell.value = f"=SUM(I{hrow+1}:I{tot_row-1})"
                cell.number_format = '#,##0.00'
                cell.font = Font(name="Calibri", size=11, bold=True, color="047857")
                cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(len(str(col[hrow - 1].value or '')) + 4, 14)

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"Premier_Bank_EFT_{sheet.salary_month}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Alias for advice letter link
    app.add_url_rule('/payroll/salary-sheet/<int:sheet_id>/export-salary-sheet-eft', 'export_salary_sheet_eft_excel', export_premier_bank_eft_excel)

    # -------------------------------------------------------------------------
    # Route: Export BEFTN Clearing Excel Sheet (Bangladesh Bank Clearing)
    # -------------------------------------------------------------------------
    @app.route('/payroll/salary-sheet/<int:sheet_id>/export-beftn-excel')
    def export_beftn_clearing_excel(sheet_id):
        sheet = MonthlySalarySheet.query.get_or_404(sheet_id)
        items = MonthlySalaryItem.query.filter_by(sheet_id=sheet.id, payment_method='BEFTN').filter(MonthlySalaryItem.status != 'Inactive').order_by(MonthlySalaryItem.bank_name.asc(), MonthlySalaryItem.name.asc()).all()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BEFTN Clearing"

        title_font = Font(name="Calibri", size=13, bold=True, color="7C3AED")
        hdr_fill = PatternFill(start_color="4C1D95", end_color="4C1D95", fill_type="solid")
        hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        bold_font = Font(name="Calibri", size=10, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        ws.cell(row=1, column=1, value="BANGLADESH ELECTRONIC FUNDS TRANSFER NETWORK (BEFTN) - CLEARING SCHEDULE").font = title_font
        ws.cell(row=2, column=1, value=f"Originating Account: 02071110000026 (The Premier Bank PLC) | Originator: ZR-ERP ENTERPRISE | Month: {sheet.salary_month}")
        ws.cell(row=3, column=1, value=f"Total Clearing Beneficiaries: {len(items)} | Clearing Format: BB-BEFTN-2026")

        headers = [
            "SL", "Originator Account", "Originator Name", "Beneficiary ID",
            "Beneficiary Name", "Receiving Bank", "Receiving Branch",
            "9-Digit Routing No", "Beneficiary Account No", "Amount (BDT)", "Payment Reference"
        ]

        hrow = 5
        ws.row_dimensions[hrow].height = 25
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=hrow, column=c_idx, value=h)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        cur = hrow + 1
        for idx, it in enumerate(items, 1):
            ws.row_dimensions[cur].height = 19
            row_data = [
                idx,
                "02071110000026",
                "ZR-ERP ENTERPRISE",
                it.emp_code,
                it.name,
                it.bank_name,
                it.bank_branch or '',
                it.routing_no or '',
                it.bank_account_no or '',
                it.net_after_tds,
                f"Salary {sheet.salary_month}"
            ]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=cur, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx == 10:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.font = bold_font
                elif c_idx in [1, 2, 4, 8, 9]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            cur += 1

        tot_row = cur
        ws.cell(row=tot_row, column=1, value="TOTAL BEFTN CLEARING (BDT):").font = bold_font
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=tot_row, column=c)
            cell.fill = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")
            cell.border = thin_border
            if c == 10:
                cell.value = f"=SUM(J{hrow+1}:J{tot_row-1})"
                cell.number_format = '#,##0.00'
                cell.font = Font(name="Calibri", size=11, bold=True, color="6D28D9")
                cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(len(str(col[hrow - 1].value or '')) + 4, 14)

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"BEFTN_Clearing_{sheet.salary_month}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Alias for advice letter link
    app.add_url_rule('/payroll/salary-sheet/<int:sheet_id>/export-salary-sheet-beftn', 'export_salary_sheet_beftn_excel', export_beftn_clearing_excel)

    # -------------------------------------------------------------------------
    # Route: Employee Bank Info Register View
    # -------------------------------------------------------------------------
    @app.route('/payroll/bank-info')
    def employee_bank_info():
        channel = request.args.get('channel', 'all').lower()
        department = request.args.get('department', '').strip()
        q = request.args.get('q', '').strip()

        query = Employee.query
        if channel == 'eft':
            query = query.filter(Employee.payment_method == 'EFT')
        elif channel == 'beftn':
            query = query.filter(Employee.payment_method == 'BEFTN')
        elif channel == 'cash':
            query = query.filter(Employee.payment_method == 'Cash')

        if department:
            query = query.filter(Employee.department == department)

        if q:
            term = f"%{q}%"
            query = query.filter(
                (Employee.name.ilike(term)) |
                (Employee.employee_id.ilike(term)) |
                (Employee.bank_account_no.ilike(term)) |
                (Employee.bank_name.ilike(term))
            )

        employees = query.order_by(Employee.department.asc(), Employee.name.asc()).all()

        total_count = Employee.query.count()
        eft_count = Employee.query.filter_by(payment_method='EFT').count()
        beftn_count = Employee.query.filter_by(payment_method='BEFTN').count()
        cash_count = Employee.query.filter_by(payment_method='Cash').count()
        missing_acct_count = Employee.query.filter(
            (Employee.bank_account_no == None) | (Employee.bank_account_no == '') | (Employee.bank_account_no == '-')
        ).count()

        departments = [d[0] for d in db.session.query(Employee.department).distinct().order_by(Employee.department).all() if d[0]]

        return render_template(
            'payroll/employee_bank_info.html',
            employees=employees,
            current_channel=channel,
            current_dept=department,
            search_query=q,
            total_count=total_count,
            eft_count=eft_count,
            beftn_count=beftn_count,
            cash_count=cash_count,
            missing_acct_count=missing_acct_count,
            departments=departments
        )

    # -------------------------------------------------------------------------
    # Route: Edit Single Employee Bank Details
    # -------------------------------------------------------------------------
    @app.route('/payroll/employees/<int:emp_id>/edit-bank', methods=['POST'])
    def employee_edit_bank(emp_id):
        emp = Employee.query.get_or_404(emp_id)
        emp.bank_name = request.form.get('bank_name', emp.bank_name).strip()
        emp.bank_branch = request.form.get('bank_branch', emp.bank_branch or '').strip()
        emp.routing_no = request.form.get('routing_no', emp.routing_no or '').strip()
        emp.bank_account_no = request.form.get('bank_account_no', emp.bank_account_no or '').strip()
        emp.payment_method = request.form.get('payment_method', emp.payment_method)
        emp.pf_status = request.form.get('pf_status', emp.pf_status or 'P')
        try:
            emp.tds_amount = float(request.form.get('tds_amount', emp.tds_amount or 0.0) or 0.0)
        except ValueError:
            pass

        db.session.commit()
        flash(f"Updated bank info for {emp.name} ({emp.employee_id}).", "success")
        return redirect(request.referrer or url_for('employee_bank_info'))

    # -------------------------------------------------------------------------
    # Route: Bulk Update Employee Bank Information (Inline Spreadsheet Mode)
    # -------------------------------------------------------------------------
    @app.route('/payroll/bank-info/bulk-update', methods=['POST'])
    def bulk_update_employee_bank_info():
        payload = request.get_json(silent=True)
        if not payload or 'employees' not in payload:
            return jsonify({'success': False, 'message': 'No employee payload received'}), 400

        updated_count = 0
        for item in payload['employees']:
            emp_id = item.get('id')
            emp = Employee.query.get(emp_id)
            if not emp:
                continue

            if 'bank_name' in item:
                emp.bank_name = str(item['bank_name'] or '').strip()
            if 'bank_branch' in item:
                emp.bank_branch = str(item['bank_branch'] or '').strip()
            if 'routing_no' in item:
                emp.routing_no = str(item['routing_no'] or '').strip()
            if 'bank_account_no' in item:
                emp.bank_account_no = str(item['bank_account_no'] or '').strip()
            if 'payment_method' in item:
                emp.payment_method = str(item['payment_method'] or 'EFT').strip()
            if 'pf_status' in item:
                emp.pf_status = str(item['pf_status'] or 'P').strip()
            if 'tds_amount' in item:
                try:
                    emp.tds_amount = float(item['tds_amount'] or 0.0)
                except (ValueError, TypeError):
                    pass
            updated_count += 1

        db.session.commit()
        return jsonify({'success': True, 'message': f'Successfully updated bank details for {updated_count} employees.'})

    # -------------------------------------------------------------------------
    # Route: Export Employee Bank Register to Excel
    # -------------------------------------------------------------------------
    @app.route('/payroll/bank-info/export-excel')
    def export_employee_bank_info_excel():
        employees = Employee.query.order_by(Employee.department.asc(), Employee.name.asc()).all()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Employee Bank Info"

        title_font = Font(name="Calibri", size=13, bold=True, color="1E3A8A")
        hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        ws.cell(row=1, column=1, value="ZR-ERP ENTERPRISE - EMPLOYEE BANK & DISBURSEMENT INFORMATION REGISTER").font = title_font
        ws.cell(row=2, column=1, value=f"Total Registered Employees: {len(employees)} | Export Date: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}")

        headers = [
            "SL", "Employee ID", "Employee Name", "Department", "Designation",
            "Payment Method", "Bank Name", "Branch Name", "9-Digit Routing No",
            "Account Number", "PF Status", "Monthly TDS (৳)", "Basic Salary (৳)"
        ]

        hrow = 4
        ws.row_dimensions[hrow].height = 25
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=hrow, column=c_idx, value=h)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        cur = hrow + 1
        for idx, emp in enumerate(employees, 1):
            ws.row_dimensions[cur].height = 19
            row_data = [
                idx,
                emp.employee_id,
                emp.name,
                emp.department,
                emp.designation,
                emp.payment_method,
                emp.bank_name,
                emp.bank_branch or '',
                emp.routing_no or '',
                emp.bank_account_no or '',
                emp.pf_status or 'P',
                emp.tds_amount or 0.0,
                emp.basic_salary or 0.0
            ]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=cur, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx in [12, 13]:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c_idx in [1, 2, 6, 9, 10, 11]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            cur += 1

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(len(str(col[hrow - 1].value or '')) + 4, 14)

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"Employee_Bank_Info_Register_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # -------------------------------------------------------------------------
    # Route: Upload / Import Employee Bank Info Excel
    # -------------------------------------------------------------------------
    @app.route('/payroll/bank-info/upload-excel', methods=['POST'])
    def upload_employee_bank_info_excel():
        if 'excel_file' not in request.files:
            flash("No file part uploaded.", "danger")
            return redirect(url_for('employee_bank_info'))

        file = request.files['excel_file']
        if not file.filename:
            flash("Please choose an Excel file to upload.", "warning")
            return redirect(url_for('employee_bank_info'))

        try:
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active

            # Find header row
            header_row_idx = 1
            col_map = {}
            for r in range(1, 10):
                row_vals = [str(ws.cell(row=r, column=c).value or '').strip().lower() for c in range(1, 20)]
                if any('employee id' in v or 'id no' in v or 'emp id' in v for v in row_vals):
                    header_row_idx = r
                    for c in range(1, 20):
                        v = str(ws.cell(row=r, column=c).value or '').strip().lower()
                        if 'employee id' in v or 'id no' in v or 'emp id' in v or v == 'id':
                            col_map['emp_id'] = c
                        elif 'bank name' in v:
                            col_map['bank_name'] = c
                        elif 'branch' in v:
                            col_map['branch'] = c
                        elif 'routing' in v:
                            col_map['routing'] = c
                        elif 'account' in v:
                            col_map['account'] = c
                        elif 'method' in v or 'channel' in v or 'transfer' in v:
                            col_map['payment_method'] = c
                        elif 'pf' in v:
                            col_map['pf_status'] = c
                        elif 'tds' in v:
                            col_map['tds'] = c
                    break

            if 'emp_id' not in col_map:
                flash("Could not detect 'Employee ID' header in the uploaded Excel. Please ensure column is labeled 'Employee ID' or 'ID No.'.", "danger")
                return redirect(url_for('employee_bank_info'))

            updated_count = 0
            for r in range(header_row_idx + 1, ws.max_row + 1):
                raw_id = ws.cell(row=r, column=col_map['emp_id']).value
                if not raw_id:
                    continue
                emp_id_str = str(raw_id).strip()

                # Match employee by exact id, or stripped, or numeric id
                emp = Employee.query.filter((Employee.employee_id == emp_id_str) | (Employee.employee_id == f"EMP-{emp_id_str.zfill(3)}")).first()
                if not emp and emp_id_str.isdigit():
                    emp = Employee.query.filter_by(id=int(emp_id_str)).first()

                if not emp:
                    continue

                if 'bank_name' in col_map:
                    val = ws.cell(row=r, column=col_map['bank_name']).value
                    if val:
                        emp.bank_name = str(val).strip()
                if 'branch' in col_map:
                    val = ws.cell(row=r, column=col_map['branch']).value
                    if val:
                        emp.bank_branch = str(val).strip()
                if 'routing' in col_map:
                    val = ws.cell(row=r, column=col_map['routing']).value
                    if val:
                        emp.routing_no = str(val).strip()
                if 'account' in col_map:
                    val = ws.cell(row=r, column=col_map['account']).value
                    if val:
                        emp.bank_account_no = str(val).strip()
                if 'payment_method' in col_map:
                    val = ws.cell(row=r, column=col_map['payment_method']).value
                    if val:
                        m = str(val).strip().upper()
                        if 'EFT' in m and 'BEFTN' not in m:
                            emp.payment_method = 'EFT'
                        elif 'BEFTN' in m:
                            emp.payment_method = 'BEFTN'
                        elif 'CASH' in m:
                            emp.payment_method = 'Cash'
                if 'pf_status' in col_map:
                    val = ws.cell(row=r, column=col_map['pf_status']).value
                    if val:
                        emp.pf_status = str(val).strip()
                if 'tds' in col_map:
                    val = ws.cell(row=r, column=col_map['tds']).value
                    if val is not None:
                        try:
                            emp.tds_amount = float(val)
                        except (ValueError, TypeError):
                            pass

                updated_count += 1

            db.session.commit()
            flash(f"Successfully processed Excel upload! Updated {updated_count} employee bank records.", "success")
        except Exception as e:
            db.session.rollback()
            flash(f"Error parsing Excel file: {str(e)}", "danger")

        return redirect(url_for('employee_bank_info'))

    print(" * Payroll Salary & Bank Transfer Routes Registered Successfully.")
