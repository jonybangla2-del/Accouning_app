"""
Verification test for Leave Excel & CSV File Importer with user-submitted data.
"""

import os
import sys
import io
import unittest
from datetime import date

sys.stdout.reconfigure(encoding='utf-8')

from app import (
    app, db, Employee, LeaveBank, LeaveMonthlyLedger,
    EmployeeAttendance, AttendanceAdjustment, GovtHoliday, LeavePosting,
    ensure_attendance_modules_tables
)
import leave_manager

class TestUserLeaveFileImport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        app.config['TESTING'] = True
        cls.client = app.test_client()
        with app.app_context():
            ensure_attendance_modules_tables()

    def setUp(self):
        self.app_ctx = app.app_context()
        self.app_ctx.push()

    def tearDown(self):
        self.app_ctx.pop()

    def test_01_parse_csv_user_data(self):
        print("\n--- Test 1: Direct CSV Parsing of User Leave Data ---")
        models = {
            'LeaveBank': LeaveBank,
            'LeaveMonthlyLedger': LeaveMonthlyLedger,
            'Employee': Employee,
            'EmployeeAttendance': EmployeeAttendance,
            'AttendanceAdjustment': AttendanceAdjustment,
            'GovtHoliday': GovtHoliday,
            'LeavePosting': LeavePosting
        }

        with open('test_sample_full.csv', 'rb') as f:
            csv_bytes = f.read()

        res = leave_manager.parse_custom_block_leave_excel(io.BytesIO(csv_bytes), db.session, models)
        print(f"Total blocks imported: {res['total_blocks_found']}")
        self.assertTrue(res['total_blocks_found'] >= 3)

        # Check Zillur Rahman (ID: 198)
        emp_198 = Employee.query.filter_by(employee_id='198').first()
        self.assertIsNotNone(emp_198)
        bank_198 = LeaveBank.query.filter_by(employee_id=emp_198.id, year=2026).first()
        self.assertIsNotNone(bank_198)
        print(f"Employee 198 (Zillur Rahman): Opening={bank_198.opening_balance}, Current Balance={bank_198.current_balance}")
        self.assertEqual(bank_198.opening_balance, 12.5625)

        # Check Md. Rahat Pathan (ID: 127)
        emp_127 = Employee.query.filter_by(employee_id='127').first()
        self.assertIsNotNone(emp_127)
        bank_127 = LeaveBank.query.filter_by(employee_id=emp_127.id, year=2026).first()
        self.assertIsNotNone(bank_127)
        print(f"Employee 127 (Md. Rahat Pathan): Opening={bank_127.opening_balance}, Current Balance={bank_127.current_balance}")
        self.assertEqual(bank_127.opening_balance, 1.21875)

        # Check Din Islam (ID: 433)
        emp_433 = Employee.query.filter_by(employee_id='433').first()
        self.assertIsNotNone(emp_433)
        bank_433 = LeaveBank.query.filter_by(employee_id=emp_433.id, year=2026).first()
        self.assertIsNotNone(bank_433)
        print(f"Employee 433 (Din Islam): Opening={bank_433.opening_balance}, Current Balance={bank_433.current_balance}")
        self.assertEqual(bank_433.opening_balance, 13.0)

        print("✔ CSV user data parsed and verified in DB.")

    def test_02_web_route_file_upload(self):
        print("\n--- Test 2: Web Route File Upload via /attendance/leave/import-excel ---")
        with self.client.session_transaction() as sess:
            sess['user_id'] = 1
            sess['username'] = 'Admin'
            sess['active_company_id'] = 1

        with open('test_sample_full.csv', 'rb') as f:
            data = {
                'excel_file': (io.BytesIO(f.read()), 'company_leave_history.csv')
            }
            res = self.client.post('/attendance/leave/import-excel', data=data, content_type='multipart/form-data', follow_redirects=True)
            print("Upload Response Status:", res.status_code)
            self.assertEqual(res.status_code, 200)
            self.assertIn(b'Successfully imported', res.data)
            print("✔ Web route accepted file and returned 200 with success flash message.")

if __name__ == '__main__':
    unittest.main()
