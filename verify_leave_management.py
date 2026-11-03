"""
Comprehensive Verification Test Suite for Enterprise Leave Management & Attendance Adjustment System
"""

import os
import sys
import io
import unittest
from datetime import date, datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8')

from app import (
    app, db, Employee, EmployeeAttendance, GovtHoliday,
    LeavePosting, RosterSwap, AttendanceAdjustment,
    LeaveBank, LeaveMonthlyLedger, LeaveTransaction, DutySwap,
    ensure_attendance_modules_tables
)
import leave_manager

class TestLeaveManagementSystem(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        app.config['TESTING'] = True
        cls.client = app.test_client()
        with app.app_context():
            ensure_attendance_modules_tables()

            # Ensure a test employee exists
            test_emp = Employee.query.filter_by(employee_id='TEST-999').first()
            if not test_emp:
                test_emp = Employee(
                    employee_id='TEST-999',
                    name='Test Leave Officer',
                    department='Engineering',
                    designation='Senior QA Lead',
                    status='Active',
                    joining_date=date(2022, 1, 1)
                )
                db.session.add(test_emp)
                db.session.commit()
            cls.test_emp_id = test_emp.id

    def setUp(self):
        self.app_ctx = app.app_context()
        self.app_ctx.push()

    def tearDown(self):
        self.app_ctx.pop()

    def test_01_models_and_aliases(self):
        print("\n--- Test 1: Models & Aliases ---")
        self.assertEqual(LeaveBank.__tablename__, 'leave_banks')
        self.assertEqual(LeaveMonthlyLedger.__tablename__, 'leave_monthly_ledgers')
        self.assertIs(LeaveTransaction, LeavePosting)
        self.assertIs(DutySwap, RosterSwap)
        print("✔ Models LeaveBank, LeaveMonthlyLedger, LeaveTransaction, DutySwap verified.")

    def test_02_working_days_calculation(self):
        print("\n--- Test 2: Working Days Calculation ---")
        wd_oct = leave_manager.get_working_days_in_month(2026, 10)
        wd_feb = leave_manager.get_working_days_in_month(2026, 2)
        print(f"Working days: Oct 2026 = {wd_oct}, Feb 2026 = {wd_feb}")
        self.assertTrue(18 <= wd_feb <= 22)
        self.assertTrue(20 <= wd_oct <= 24)
        print("✔ Working days calculation excluding Fridays & Saturdays verified.")

    def test_03_monthly_duty_pto_accrual_policy(self):
        print("\n--- Test 3: 50% Monthly Duty General PTO Rule & Holiday Reward ---")
        models = {
            'LeaveBank': LeaveBank,
            'LeaveMonthlyLedger': LeaveMonthlyLedger,
            'Employee': Employee,
            'EmployeeAttendance': EmployeeAttendance,
            'AttendanceAdjustment': AttendanceAdjustment,
            'GovtHoliday': GovtHoliday,
            'LeavePosting': LeavePosting
        }

        # Clear attendance for test employee in Nov 2026
        EmployeeAttendance.query.filter(
            EmployeeAttendance.employee_id == self.test_emp_id,
            EmployeeAttendance.attendance_date >= date(2026, 11, 1),
            EmployeeAttendance.attendance_date <= date(2026, 11, 30)
        ).delete()
        db.session.commit()

        # Case A: <50% duty attendance (e.g. 5 days present)
        for d in range(1, 6):
            att = EmployeeAttendance(
                employee_id=self.test_emp_id,
                attendance_date=date(2026, 11, d),
                shift_type='Day Shift',
                status='Present',
                total_hours=8.0
            )
            db.session.add(att)
        db.session.commit()

        calc_a = leave_manager.calculate_monthly_duty_and_pto(
            self.test_emp_id, 2026, 11, db.session,
            EmployeeAttendance, AttendanceAdjustment, GovtHoliday, LeavePosting
        )
        print(f"Case A (5 days duty): Duty={calc_a['duty_days_count']}/{calc_a['working_days']} ({calc_a['duty_percentage']}%) -> Eligible={calc_a['is_eligible_pto']}, General PTO={calc_a['general_pto']}")
        self.assertFalse(calc_a['is_eligible_pto'])
        self.assertEqual(calc_a['general_pto'], 0.0)

        # Case B: >=50% duty attendance (add 10 more days -> 15 days total >= 11 days)
        for d in range(6, 16):
            att = EmployeeAttendance(
                employee_id=self.test_emp_id,
                attendance_date=date(2026, 11, d),
                shift_type='Day Shift',
                status='Present',
                total_hours=8.0
            )
            db.session.add(att)

        # Add a Govt Holiday on Nov 10 and verify Holiday PTO reward (+1.0)
        h = GovtHoliday.query.filter_by(holiday_date=date(2026, 11, 10)).first()
        if not h:
            h = GovtHoliday(
                name="Test Special Holiday",
                holiday_date=date(2026, 11, 10),
                year=2026,
                days_count=1
            )
            db.session.add(h)
        db.session.commit()

        calc_b = leave_manager.calculate_monthly_duty_and_pto(
            self.test_emp_id, 2026, 11, db.session,
            EmployeeAttendance, AttendanceAdjustment, GovtHoliday, LeavePosting
        )
        print(f"Case B (15 days duty + Holiday): Duty={calc_b['duty_days_count']} -> Eligible={calc_b['is_eligible_pto']}, General PTO={calc_b['general_pto']}, Holiday PTO={calc_b['holiday_pto']}")
        self.assertTrue(calc_b['is_eligible_pto'])
        self.assertEqual(calc_b['general_pto'], 1.5)
        self.assertEqual(calc_b['holiday_pto'], 1.0)
        print("✔ 50% duty condition and Govt Holiday PTO reward (+1.0) verified.")

    def test_04_12_month_cascading_and_50pct_carry_forward(self):
        print("\n--- Test 4: Continuous 12-Month Running Balance & 50% Carry Forward ---")
        models = {
            'LeaveBank': LeaveBank,
            'LeaveMonthlyLedger': LeaveMonthlyLedger,
            'Employee': Employee,
            'EmployeeAttendance': EmployeeAttendance,
            'AttendanceAdjustment': AttendanceAdjustment,
            'GovtHoliday': GovtHoliday,
            'LeavePosting': LeavePosting
        }

        # Initialize bank with opening balance = 10.0
        bank = LeaveBank.query.filter_by(employee_id=self.test_emp_id, year=2026).first()
        if not bank:
            bank = LeaveBank(company_id=1, employee_id=self.test_emp_id, year=2026, opening_balance=10.0)
            db.session.add(bank)
        else:
            bank.opening_balance = 10.0
        db.session.commit()

        bank = leave_manager.recalculate_employee_leave_ledger(self.test_emp_id, 2026, db.session, models, force_recalculate=True)
        rows = LeaveMonthlyLedger.query.filter_by(employee_id=self.test_emp_id, year=2026).order_by(LeaveMonthlyLedger.month.asc()).all()

        self.assertEqual(len(rows), 12)
        self.assertEqual(rows[0].opening_balance, 10.0)

        # Verify cascading: for each month m > 1, opening balance == previous month closing balance
        for m in range(1, 12):
            self.assertAlmostEqual(rows[m].opening_balance, rows[m-1].closing_balance, places=2)

        # Verify 50% carry forward formula: Dec closing * 0.5
        dec_closing = rows[11].closing_balance
        expected_cf = round(max(0.0, dec_closing * 0.5), 2)
        print(f"Dec Closing Balance = {dec_closing}, Carried Forward to 2027 (50%) = {bank.carried_forward_to_next_year} (expected {expected_cf})")
        self.assertEqual(bank.carried_forward_to_next_year, expected_cf)
        print("✔ 12-month balance cascading and 50% carry forward rollover verified.")

    def test_05_sample_template_and_block_excel_parser(self):
        print("\n--- Test 5: Sample Template Generation & Custom Block Excel Parser ---")
        models = {
            'LeaveBank': LeaveBank,
            'LeaveMonthlyLedger': LeaveMonthlyLedger,
            'Employee': Employee,
            'EmployeeAttendance': EmployeeAttendance,
            'AttendanceAdjustment': AttendanceAdjustment,
            'GovtHoliday': GovtHoliday,
            'LeavePosting': LeavePosting
        }

        # Generate sample template
        excel_buf = leave_manager.generate_sample_leave_excel_template()
        self.assertIsNotNone(excel_buf)
        excel_bytes = excel_buf.getvalue()
        self.assertTrue(len(excel_bytes) > 2000)
        print(f"Generated sample template: {len(excel_bytes)} bytes.")

        # Parse generated template back
        parse_res = leave_manager.parse_custom_block_leave_excel(io.BytesIO(excel_bytes), db.session, models)
        print(f"Parsed blocks found: {parse_res['total_blocks_found']}, Employees: {len(parse_res['employees'])}")
        self.assertTrue(parse_res['total_blocks_found'] >= 2)
        print("✔ Custom block Excel parser successfully parsed sample template.")

    def test_06_excel_exporters(self):
        print("\n--- Test 6: Single Statement & All Staff Summary Excel Exporters ---")
        models = {
            'LeaveBank': LeaveBank,
            'LeaveMonthlyLedger': LeaveMonthlyLedger,
            'Employee': Employee,
            'EmployeeAttendance': EmployeeAttendance,
            'AttendanceAdjustment': AttendanceAdjustment,
            'GovtHoliday': GovtHoliday,
            'LeavePosting': LeavePosting
        }

        # Statement export
        stmt_buf = leave_manager.export_employee_leave_statement_excel(self.test_emp_id, 2026, db.session, models)
        self.assertIsNotNone(stmt_buf)
        self.assertTrue(len(stmt_buf.getvalue()) > 2000)

        # Summary export
        sum_buf = leave_manager.export_all_staff_leave_summary_excel(2026, db.session, models)
        self.assertIsNotNone(sum_buf)
        self.assertTrue(len(sum_buf.getvalue()) > 2000)
        print("✔ Excel exporters verified.")

    def test_07_api_and_web_routes(self):
        print("\n--- Test 7: API Endpoints & Web Route Responses ---")
        # Test client session simulation
        with self.client.session_transaction() as sess:
            sess['user_id'] = 1
            sess['username'] = 'Admin'
            sess['active_company_id'] = 1

        # 1. API: Live Leave Balance
        res_bal = self.client.get('/attendance/api/leave-balance/TEST-999?year=2026')
        print("API Leave Balance status:", res_bal.status_code)
        self.assertEqual(res_bal.status_code, 200)
        data = res_bal.get_json()
        self.assertTrue(data['success'])
        self.assertIn('available_balance', data)
        print(f"Live balance for TEST-999: {data['available_balance']} days (Carried: {data['carried_forward_to_next_year']})")

        # 2. View Ledger Page
        res_ledger = self.client.get(f'/attendance/leave/ledger/{self.test_emp_id}?year=2026')
        print("View Ledger status:", res_ledger.status_code)
        self.assertEqual(res_ledger.status_code, 200)
        self.assertIn(b'Individual Leave Statement', res_ledger.data)

        # 3. View Summary Page
        res_sum = self.client.get('/attendance/leave/summary?year=2026')
        print("View Summary status:", res_sum.status_code)
        self.assertEqual(res_sum.status_code, 200)
        self.assertIn(b'All Staff Leave Bank', res_sum.data)

        # 4. Template download
        res_tmpl = self.client.get('/attendance/leave/template/download')
        print("Template download status:", res_tmpl.status_code)
        self.assertEqual(res_tmpl.status_code, 200)

        # 5. Export Statement
        res_exp = self.client.get(f'/attendance/leave/ledger/{self.test_emp_id}/export-excel?year=2026')
        print("Export Statement status:", res_exp.status_code)
        self.assertEqual(res_exp.status_code, 200)

        print("✔ All API endpoints and web routes verified successfully.")

if __name__ == '__main__':
    unittest.main()
