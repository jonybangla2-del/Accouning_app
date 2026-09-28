import os
import io
from datetime import datetime, date, timedelta
from functools import wraps
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from werkzeug.security import generate_password_hash, check_password_hash
import re
import csv
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pypdf

app = Flask(__name__)

# Configure SQLite database
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SECRET_KEY'] = 'zr-erp-enterprise-2026-secret'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'accounting.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'connect_args': {
        'timeout': 60,
        'check_same_thread': False
    }
}

db = SQLAlchemy(app)

from collections.abc import Mapping
from company_manager import CompanyManager, CompanyProfile
import email_service
from email_service import ADMIN_EMAIL, send_otp_email, generate_otp, get_smtp_config, save_smtp_config

class DynamicEngineMap(Mapping):
    def __init__(self, default_engine):
        self.default_engine = default_engine

    def __getitem__(self, key):
        from flask import has_request_context, session
        if has_request_context():
            cid = session.get('active_company_id', 1)
            return CompanyManager.get_engine(cid, metadata=db.metadata)
        return self.default_engine

    def __iter__(self):
        yield None

    def __len__(self):
        return 1

    def __contains__(self, key):
        return key is None

default_eng = db._app_engines[app][None]
db._app_engines[app] = DynamicEngineMap(default_eng)

from sqlalchemy import event
from sqlalchemy.engine import Engine
import sqlite3

@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    if isinstance(dbapi_connection, sqlite3.Connection):
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA busy_timeout=60000")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.close()


# ==============================================================================
# DATABASE MODELS
# ==============================================================================

class User(db.Model):
    __tablename__ = 'user'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(100), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='General User') # 'Admin' or 'General User'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __init__(self, username="", password="", full_name="", role="General User", **kwargs):
        super().__init__(**kwargs)
        self.username = username
        self.full_name = full_name
        self.role = role
        if password:
            self.set_password(password)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.username} ({self.role})>'

class Account(db.Model):
    __tablename__ = 'account'
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), unique=True, nullable=False) # e.g. 1010, 1030-01
    name = db.Column(db.String(100), unique=True, nullable=False)
    type = db.Column(db.String(50), nullable=False) # Asset, Liability, Equity, Revenue, Expense
    category = db.Column(db.String(50), nullable=False) # Current Asset, Non-Current Asset, etc.
    parent_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=True) # Subledger link
    is_cash_or_bank = db.Column(db.Boolean, default=False)
    opening_balance = db.Column(db.Float, default=0.00)
    balance = db.Column(db.Float, default=0.00)

    # Self-referencing relationship for subledgers
    parent = db.relationship('Account', remote_side=[id], backref=db.backref('subledgers', lazy=True, cascade='all, delete-orphan'))

    def __init__(self, code="", name="", type="", category="", parent_id=None, is_cash_or_bank=False, opening_balance=0.00, balance=0.00, **kwargs):
        super().__init__(**kwargs)
        self.code = code
        self.name = name
        self.type = type
        self.category = category
        self.parent_id = parent_id
        self.is_cash_or_bank = is_cash_or_bank
        self.opening_balance = opening_balance
        self.balance = balance

    def __repr__(self):
        return f'<Account {self.code} - {self.name}>'

class JournalEntry(db.Model):
    __tablename__ = 'journal_entry'
    id = db.Column(db.Integer, primary_key=True)
    voucher_no = db.Column(db.String(50), unique=True, nullable=False, index=True)
    voucher_type = db.Column(db.String(20), nullable=False) # Payment, Receipt, Journal, Contra
    date = db.Column(db.Date, nullable=False, default=date.today)
    reference = db.Column(db.String(100), nullable=True) # Cheque #, Inv #
    narration = db.Column(db.String(500), nullable=False)
    total_amount = db.Column(db.Float, default=0.00)
    attachment = db.Column(db.String(255), nullable=True) # Supporting document filename
    original_filename = db.Column(db.String(255), nullable=True) # Original uploaded filename
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    lines = db.relationship('TransactionLine', backref='journal_entry', cascade="all, delete-orphan", lazy=True)

    def __init__(self, voucher_no="", voucher_type="", date=None, reference=None, narration="", total_amount=0.00, attachment=None, original_filename=None, **kwargs):
        super().__init__(**kwargs)
        self.voucher_no = voucher_no
        self.voucher_type = voucher_type
        self.date = date if date is not None else datetime.utcnow().date()
        self.reference = reference
        self.narration = narration
        self.total_amount = total_amount
        self.attachment = attachment
        self.original_filename = original_filename

    def __repr__(self):
        return f'<JournalEntry {self.voucher_no} ({self.voucher_type})>'

class TransactionLine(db.Model):
    __tablename__ = 'transaction_line'
    id = db.Column(db.Integer, primary_key=True)
    journal_entry_id = db.Column(db.Integer, db.ForeignKey('journal_entry.id'), nullable=False)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    debit = db.Column(db.Float, default=0.00)
    credit = db.Column(db.Float, default=0.00)
    particulars = db.Column(db.String(255), nullable=True)

    account = db.relationship('Account', backref=db.backref('transactions', lazy=True))

    def __init__(self, journal_entry_id=None, account_id=None, debit=0.00, credit=0.00, particulars=None, **kwargs):
        super().__init__(**kwargs)
        self.journal_entry_id = journal_entry_id
        self.account_id = account_id
        self.debit = debit
        self.credit = credit
        self.particulars = particulars

    def __repr__(self):
        return f'<Line Acc:{self.account_id} Dr:{self.debit} Cr:{self.credit}>'

class Employee(db.Model):
    __tablename__ = 'employee'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(20), unique=True, nullable=False, index=True) # e.g. EMP-001
    name = db.Column(db.String(100), nullable=False)
    designation = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), nullable=False)
    basic_salary = db.Column(db.Float, default=0.00)
    allowances = db.Column(db.Float, default=0.00)
    deductions = db.Column(db.Float, default=0.00)
    bank_account_no = db.Column(db.String(50), nullable=True)
    status = db.Column(db.String(20), default='Active') # Active, Inactive
    joining_date = db.Column(db.Date, nullable=True)
    release_date = db.Column(db.Date, nullable=True)
    nid_passport = db.Column(db.String(50), nullable=True)
    email = db.Column(db.String(100), nullable=True)
    phone = db.Column(db.String(30), nullable=True)
    present_address = db.Column(db.String(255), nullable=True)
    gender = db.Column(db.String(10), default='Male')
    house_rent = db.Column(db.Float, default=0.00)
    medical_allowance = db.Column(db.Float, default=0.00)
    conveyance = db.Column(db.Float, default=0.00)
    provident_fund = db.Column(db.Float, default=0.00)
    bank_name = db.Column(db.String(100), default='The Premier Bank PLC')
    bank_branch = db.Column(db.String(100), default='Banani Branch')
    routing_no = db.Column(db.String(20), default='085261198')
    payment_method = db.Column(db.String(20), default='EFT') # 'EFT', 'BEFTN', 'Cash'
    pf_status = db.Column(db.String(10), default='P')
    tds_amount = db.Column(db.Float, default=0.00)
    employment_type = db.Column(db.String(20), default='Full-Time') # 'Full-Time', 'Part-Time'
    duty_category = db.Column(db.String(50), nullable=True) # 'News Presenter', 'Customer Care Agent'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __init__(self, employee_id="", name="", designation="", department="", basic_salary=0.00, allowances=0.00, deductions=0.00, bank_account_no=None, status="Active", joining_date=None, release_date=None, nid_passport=None, email=None, phone=None, present_address=None, gender="Male", house_rent=0.0, medical_allowance=0.0, conveyance=0.0, provident_fund=0.0, bank_name='The Premier Bank PLC', bank_branch='Banani Branch', routing_no='085261198', payment_method='EFT', pf_status='P', tds_amount=0.0, employment_type='Full-Time', duty_category=None, **kwargs):
        super().__init__(**kwargs)
        self.employee_id = employee_id
        self.name = name
        self.designation = designation
        self.department = department
        self.basic_salary = basic_salary
        self.allowances = allowances
        self.deductions = deductions
        self.bank_account_no = bank_account_no
        self.status = status
        self.joining_date = joining_date
        self.release_date = release_date
        self.nid_passport = nid_passport
        self.email = email
        self.phone = phone
        self.present_address = present_address
        self.gender = gender
        self.house_rent = house_rent
        self.medical_allowance = medical_allowance
        self.conveyance = conveyance
        self.provident_fund = provident_fund
        self.bank_name = bank_name
        self.bank_branch = bank_branch
        self.routing_no = routing_no
        self.payment_method = payment_method
        self.pf_status = pf_status
        self.tds_amount = tds_amount
        self.employment_type = employment_type
        self.duty_category = duty_category

    def __repr__(self):
        return f'<Employee {self.employee_id} - {self.name}>'

class PayrollRecord(db.Model):
    __tablename__ = 'payroll_record'
    id = db.Column(db.Integer, primary_key=True)
    payroll_month = db.Column(db.String(20), nullable=False) # e.g. "2026-09"
    processed_date = db.Column(db.Date, nullable=False, default=date.today)
    total_basic = db.Column(db.Float, default=0.00)
    total_allowances = db.Column(db.Float, default=0.00)
    total_deductions = db.Column(db.Float, default=0.00)
    total_net = db.Column(db.Float, default=0.00)
    payment_account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    voucher_id = db.Column(db.Integer, db.ForeignKey('journal_entry.id'), nullable=True) # Linked PV
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    payment_account = db.relationship('Account')
    voucher = db.relationship('JournalEntry')
    items = db.relationship('PayrollItem', backref='payroll_record', cascade="all, delete-orphan", lazy=True)

    def __init__(self, payroll_month="", processed_date=None, total_basic=0.00, total_allowances=0.00, total_deductions=0.00, total_net=0.00, payment_account_id=None, voucher_id=None, **kwargs):
        super().__init__(**kwargs)
        self.payroll_month = payroll_month
        self.processed_date = processed_date if processed_date is not None else datetime.utcnow().date()
        self.total_basic = total_basic
        self.total_allowances = total_allowances
        self.total_deductions = total_deductions
        self.total_net = total_net
        self.payment_account_id = payment_account_id
        self.voucher_id = voucher_id

    def __repr__(self):
        return f'<PayrollRecord {self.payroll_month} - ${self.total_net}>'

class PayrollItem(db.Model):
    __tablename__ = 'payroll_item'
    id = db.Column(db.Integer, primary_key=True)
    payroll_record_id = db.Column(db.Integer, db.ForeignKey('payroll_record.id'), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    basic_salary = db.Column(db.Float, default=0.00)
    allowances = db.Column(db.Float, default=0.00)
    deductions = db.Column(db.Float, default=0.00)
    net_salary = db.Column(db.Float, default=0.00)

    employee = db.relationship('Employee')

    def __init__(self, payroll_record_id=None, employee_id=None, basic_salary=0.00, allowances=0.00, deductions=0.00, net_salary=0.00, **kwargs):
        super().__init__(**kwargs)
        self.payroll_record_id = payroll_record_id
        self.employee_id = employee_id
        self.basic_salary = basic_salary
        self.allowances = allowances
        self.deductions = deductions
        self.net_salary = net_salary

    def __repr__(self):
        return f'<PayrollItem Emp:{self.employee_id} Net:${self.net_salary}>'


# ============================================================================
# MONTHLY CUSTOMIZED SALARY SHEET & BANK TRANSFER MODELS
# ============================================================================
class MonthlySalarySheet(db.Model):
    __tablename__ = 'monthly_salary_sheet'
    id = db.Column(db.Integer, primary_key=True)
    salary_month = db.Column(db.String(20), nullable=False, unique=True, index=True) # e.g. "2026-09"
    title = db.Column(db.String(150), nullable=False) # e.g. "Monthly Salary Sheet for September 2026"
    processed_date = db.Column(db.Date, nullable=False, default=date.today)
    working_days = db.Column(db.Integer, default=30)

    # Core Salary Aggregates
    total_basic = db.Column(db.Float, default=0.0)
    total_hr_allowance = db.Column(db.Float, default=0.0)
    total_medical_allowance = db.Column(db.Float, default=0.0)
    total_conveyance = db.Column(db.Float, default=0.0)
    total_gross = db.Column(db.Float, default=0.0)

    # Provident Fund Aggregates
    total_pf_employer = db.Column(db.Float, default=0.0)
    total_pf_employee = db.Column(db.Float, default=0.0)
    total_pf = db.Column(db.Float, default=0.0)
    total_salary = db.Column(db.Float, default=0.0) # Gross - PF Employee

    # Additional Allowances
    total_holiday_allowance = db.Column(db.Float, default=0.0)
    total_loyalty_allowance = db.Column(db.Float, default=0.0)
    total_additional_allowance = db.Column(db.Float, default=0.0)
    total_eid_bonus = db.Column(db.Float, default=0.0)
    total_other_allowance = db.Column(db.Float, default=0.0)

    # Attendance & Deductions
    total_absenteeism_deduction = db.Column(db.Float, default=0.0)
    total_late_deduction = db.Column(db.Float, default=0.0)
    total_loan_deduction = db.Column(db.Float, default=0.0)
    total_other_deduction = db.Column(db.Float, default=0.0)
    total_arrear = db.Column(db.Float, default=0.0)

    # Net Disbursal Aggregates
    total_net_pay = db.Column(db.Float, default=0.0)
    total_tds = db.Column(db.Float, default=0.0)
    total_net_after_tds = db.Column(db.Float, default=0.0)

    # Payment Channels Breakdown
    total_eft_amount = db.Column(db.Float, default=0.0) # Premier Bank Internal EFT
    total_eft_count = db.Column(db.Integer, default=0)
    total_beftn_amount = db.Column(db.Float, default=0.0) # Other Banks BEFTN
    total_beftn_count = db.Column(db.Integer, default=0)
    total_cash_amount = db.Column(db.Float, default=0.0)
    total_cash_count = db.Column(db.Integer, default=0)

    status = db.Column(db.String(20), default='Draft') # 'Draft', 'Finalized', 'Disbursed'
    payment_account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=True) # e.g. Premier Bank
    voucher_id = db.Column(db.Integer, db.ForeignKey('journal_entry.id'), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    payment_account = db.relationship('Account')
    voucher = db.relationship('JournalEntry')
    items = db.relationship('MonthlySalaryItem', backref='salary_sheet', cascade="all, delete-orphan", lazy=True)


class MonthlySalaryItem(db.Model):
    __tablename__ = 'monthly_salary_item'
    id = db.Column(db.Integer, primary_key=True)
    sheet_id = db.Column(db.Integer, db.ForeignKey('monthly_salary_sheet.id'), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    emp_code = db.Column(db.String(20), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), nullable=False)
    sub_department = db.Column(db.String(100), nullable=True)
    designation = db.Column(db.String(100), nullable=False)
    joining_date = db.Column(db.Date, nullable=True)

    # Bank Information & Payment Method
    bank_name = db.Column(db.String(100), default='The Premier Bank PLC')
    bank_branch = db.Column(db.String(100), nullable=True)
    routing_no = db.Column(db.String(20), nullable=True)
    bank_account_no = db.Column(db.String(50), nullable=True)
    payment_method = db.Column(db.String(20), default='EFT') # 'EFT', 'BEFTN', 'Cash'

    # Core Salary
    basic = db.Column(db.Float, default=0.0)
    hr_allowance = db.Column(db.Float, default=0.0)
    medical_allowance = db.Column(db.Float, default=0.0)
    conveyance_allowance = db.Column(db.Float, default=0.0)
    gross_salary = db.Column(db.Float, default=0.0)

    # Provident Fund
    pf_status = db.Column(db.String(10), default='P')
    pf_employer = db.Column(db.Float, default=0.0)
    pf_employee = db.Column(db.Float, default=0.0)
    total_pf = db.Column(db.Float, default=0.0)
    total_salary = db.Column(db.Float, default=0.0)

    # Allowances
    govt_holiday_duty = db.Column(db.Float, default=0.0)
    holiday_extra_duty = db.Column(db.Float, default=0.0)
    loyalty_allowance = db.Column(db.Float, default=0.0)
    additional_allowance = db.Column(db.Float, default=0.0)
    eid_bonus = db.Column(db.Float, default=0.0)
    other_allowance = db.Column(db.Float, default=0.0)

    # Attendance & Deductions
    working_days = db.Column(db.Integer, default=30)
    absent_days = db.Column(db.Integer, default=0)
    late_days = db.Column(db.Integer, default=0)
    absenteeism_amount = db.Column(db.Float, default=0.0)
    late_leave_amount = db.Column(db.Float, default=0.0)
    loan_deduction = db.Column(db.Float, default=0.0)
    other_deduction = db.Column(db.Float, default=0.0)
    arrear = db.Column(db.Float, default=0.0)

    # Net Disbursal
    net_pay = db.Column(db.Float, default=0.0)
    tds_deduction = db.Column(db.Float, default=0.0)
    net_after_tds = db.Column(db.Float, default=0.0)
    remarks = db.Column(db.String(255), nullable=True)
    status = db.Column(db.String(20), default='Running') # 'Running', 'Inactive'

    # Part-Time / Per-Duty Specific Columns
    employment_type = db.Column(db.String(20), default='Full-Time') # 'Full-Time', 'Part-Time'
    duty_category = db.Column(db.String(50), nullable=True) # 'News Presenter', 'Customer Care Agent'
    shift_eve_duty = db.Column(db.Float, default=0.0)
    shift_eve_rate = db.Column(db.Float, default=800.0)
    shift_night_duty = db.Column(db.Float, default=0.0)
    shift_night_rate = db.Column(db.Float, default=1400.0)
    shift_morn_duty = db.Column(db.Float, default=0.0)
    shift_morn_rate = db.Column(db.Float, default=1500.0)
    shift_aft_duty = db.Column(db.Float, default=0.0)
    shift_aft_rate = db.Column(db.Float, default=700.0)
    total_duty = db.Column(db.Float, default=0.0)

    employee = db.relationship('Employee')



class CashBookRow(db.Model):
    __tablename__ = 'cash_book_row'
    id = db.Column(db.Integer, primary_key=True)
    row_order = db.Column(db.Integer, default=0, index=True)

    # Receipts Columns
    receipt_code = db.Column(db.String(50), nullable=True)
    receipt_date = db.Column(db.Date, nullable=True)
    receipt_voucher = db.Column(db.String(50), nullable=True)
    receipt_desc = db.Column(db.String(255), nullable=True)
    receipt_ref = db.Column(db.String(100), nullable=True)
    receipt_cash = db.Column(db.Float, default=0.0)
    receipt_premier = db.Column(db.Float, default=0.0)
    receipt_bank = db.Column(db.Float, default=0.0)
    receipt_total = db.Column(db.Float, default=0.0)

    # Payments Columns
    payment_date = db.Column(db.Date, nullable=True)
    payment_vno = db.Column(db.String(50), nullable=True)
    payment_code = db.Column(db.String(50), nullable=True)
    payment_desc = db.Column(db.String(255), nullable=True)
    payment_ref = db.Column(db.String(100), nullable=True)
    payment_cash = db.Column(db.Float, default=0.0)
    payment_bank = db.Column(db.Float, default=0.0)
    payment_premier = db.Column(db.Float, default=0.0)
    payment_total = db.Column(db.Float, default=0.0)

    # Balance Columns
    balance_cash = db.Column(db.Float, default=0.0)
    balance_bank = db.Column(db.Float, default=0.0)
    balance_premier = db.Column(db.Float, default=0.0)
    remarks = db.Column(db.String(500), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def __repr__(self):
        return f'<CashBookRow {self.row_order} - R:{self.receipt_total} P:{self.payment_total}>'


class BankStatement(db.Model):
    __tablename__ = 'bank_statement'
    id = db.Column(db.Integer, primary_key=True)
    bank_account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    statement_title = db.Column(db.String(150), nullable=False)
    filename = db.Column(db.String(255), nullable=True)
    from_date = db.Column(db.Date, nullable=False)
    to_date = db.Column(db.Date, nullable=False)
    opening_balance = db.Column(db.Float, default=0.00)
    closing_balance = db.Column(db.Float, default=0.00)
    total_withdrawals = db.Column(db.Float, default=0.00)
    total_deposits = db.Column(db.Float, default=0.00)
    notes = db.Column(db.String(500), nullable=True, default='')
    imported_at = db.Column(db.DateTime, default=datetime.utcnow)

    bank_account = db.relationship('Account')
    lines = db.relationship('BankStatementLine', backref='statement', cascade="all, delete-orphan", lazy=True)

    def __init__(self, bank_account_id=None, statement_title="", filename=None, from_date=None, to_date=None, opening_balance=0.0, closing_balance=0.0, total_withdrawals=0.0, total_deposits=0.0, notes="", **kwargs):
        super().__init__(**kwargs)
        self.bank_account_id = bank_account_id
        self.statement_title = statement_title
        self.filename = filename
        self.from_date = from_date
        self.to_date = to_date
        self.opening_balance = opening_balance
        self.closing_balance = closing_balance
        self.total_withdrawals = total_withdrawals
        self.total_deposits = total_deposits
        self.notes = notes

    def __repr__(self):
        return f'<BankStatement {self.statement_title}>'


class BankStatementLine(db.Model):
    __tablename__ = 'bank_statement_line'
    id = db.Column(db.Integer, primary_key=True)
    statement_id = db.Column(db.Integer, db.ForeignKey('bank_statement.id'), nullable=True)
    bank_account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    txn_date = db.Column(db.Date, nullable=False)
    value_date = db.Column(db.Date, nullable=True)
    particulars = db.Column(db.String(500), nullable=False)
    cheque_ref = db.Column(db.String(100), nullable=True)
    withdrawal_dr = db.Column(db.Float, default=0.00)
    deposit_cr = db.Column(db.Float, default=0.00)
    balance = db.Column(db.Float, default=0.00)
    notes = db.Column(db.String(500), nullable=True, default='')
    is_reconciled = db.Column(db.Boolean, default=False)
    reconciled_date = db.Column(db.Date, nullable=True)
    matched_voucher_id = db.Column(db.Integer, db.ForeignKey('journal_entry.id'), nullable=True)
    matched_voucher_no = db.Column(db.String(50), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    bank_account = db.relationship('Account')
    voucher = db.relationship('JournalEntry')

    def __init__(self, statement_id=None, bank_account_id=None, txn_date=None, value_date=None, particulars="", cheque_ref=None, withdrawal_dr=0.0, deposit_cr=0.0, balance=0.0, notes="", is_reconciled=False, reconciled_date=None, matched_voucher_id=None, matched_voucher_no=None, **kwargs):
        super().__init__(**kwargs)
        self.statement_id = statement_id
        self.bank_account_id = bank_account_id
        self.txn_date = txn_date
        self.value_date = value_date
        self.particulars = particulars
        self.cheque_ref = cheque_ref
        self.withdrawal_dr = withdrawal_dr
        self.deposit_cr = deposit_cr
        self.balance = balance
        self.notes = notes
        self.is_reconciled = is_reconciled
        self.reconciled_date = reconciled_date
        self.matched_voucher_id = matched_voucher_id
        self.matched_voucher_no = matched_voucher_no

    def __repr__(self):
        return f'<BankLine {self.txn_date} Dr:{self.withdrawal_dr} Cr:{self.deposit_cr}>'


class BankReconciliation(db.Model):
    __tablename__ = 'bank_reconciliation'
    id = db.Column(db.Integer, primary_key=True)
    bank_account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    recon_month = db.Column(db.String(20), nullable=False)
    statement_date = db.Column(db.Date, nullable=False)
    statement_balance = db.Column(db.Float, default=0.00)
    book_balance = db.Column(db.Float, default=0.00)
    uncredited_deposits = db.Column(db.Float, default=0.00)
    unpresented_cheques = db.Column(db.Float, default=0.00)
    bank_charges = db.Column(db.Float, default=0.00)
    direct_credits = db.Column(db.Float, default=0.00)
    adjusted_bank_balance = db.Column(db.Float, default=0.00)
    adjusted_book_balance = db.Column(db.Float, default=0.00)
    variance = db.Column(db.Float, default=0.00)
    status = db.Column(db.String(20), default='Reconciled')
    notes = db.Column(db.Text, nullable=True)
    reconciled_by = db.Column(db.String(100), default='Administrator')
    reconciled_at = db.Column(db.DateTime, default=datetime.utcnow)

    bank_account = db.relationship('Account')

    def __init__(self, bank_account_id=None, recon_month="", statement_date=None, statement_balance=0.0, book_balance=0.0, uncredited_deposits=0.0, unpresented_cheques=0.0, bank_charges=0.0, direct_credits=0.0, adjusted_bank_balance=0.0, adjusted_book_balance=0.0, variance=0.0, status="Reconciled", notes=None, reconciled_by="Administrator", **kwargs):
        super().__init__(**kwargs)
        self.bank_account_id = bank_account_id
        self.recon_month = recon_month
        self.statement_date = statement_date
        self.statement_balance = statement_balance
        self.book_balance = book_balance
        self.uncredited_deposits = uncredited_deposits
        self.unpresented_cheques = unpresented_cheques
        self.bank_charges = bank_charges
        self.direct_credits = direct_credits
        self.adjusted_bank_balance = adjusted_bank_balance
        self.adjusted_book_balance = adjusted_book_balance
        self.variance = variance
        self.status = status
        self.notes = notes
        self.reconciled_by = reconciled_by

    def __repr__(self):
        return f'<BankReconciliation {self.recon_month} Acc:{self.bank_account_id}>'


class EmployeeFundContribution(db.Model):
    __tablename__ = 'employee_fund_contribution'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    contribution_month = db.Column(db.String(20), nullable=False)
    contribution_date = db.Column(db.Date, nullable=False, default=date.today)
    employee_share = db.Column(db.Float, default=0.00) # PF Employee Share
    employer_share = db.Column(db.Float, default=0.00) # PF Employer Share
    welfare_fund = db.Column(db.Float, default=0.00)   # Welfare Fund Contribution
    profit_interest = db.Column(db.Float, default=0.00)
    total_monthly = db.Column(db.Float, default=0.00)
    cumulative_balance = db.Column(db.Float, default=0.00)
    remarks = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship('Employee', backref=db.backref('fund_contributions', lazy=True, cascade='all, delete-orphan'))

    def __init__(self, employee_id=None, contribution_month="", contribution_date=None, employee_share=0.0, employer_share=0.0, welfare_fund=0.0, profit_interest=0.0, total_monthly=0.0, cumulative_balance=0.0, remarks=None, **kwargs):
        super().__init__(**kwargs)
        self.employee_id = employee_id
        self.contribution_month = contribution_month
        self.contribution_date = contribution_date
        self.employee_share = employee_share
        self.employer_share = employer_share
        self.welfare_fund = welfare_fund
        self.profit_interest = profit_interest
        self.total_monthly = total_monthly
        self.cumulative_balance = cumulative_balance
        self.remarks = remarks

    def __repr__(self):
        return f'<FundContribution Emp:{self.employee_id} {self.contribution_month} ${self.total_monthly}>'


class EmployeeSalaryHistory(db.Model):
    __tablename__ = 'employee_salary_history'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    salary_month = db.Column(db.String(20), nullable=False)
    payment_date = db.Column(db.Date, nullable=False, default=date.today)
    department = db.Column(db.String(100), nullable=True)
    designation = db.Column(db.String(100), nullable=True)
    joining_date = db.Column(db.Date, nullable=True)

    # 1. Base Compensation
    basic_salary = db.Column(db.Float, default=0.00)
    house_rent = db.Column(db.Float, default=0.00)
    medical = db.Column(db.Float, default=0.00)
    conveyance = db.Column(db.Float, default=0.00)
    gross_salary = db.Column(db.Float, default=0.00)

    # 2. Provident Fund
    pf_rate = db.Column(db.Float, default=0.00)
    pf_employer = db.Column(db.Float, default=0.00)
    pf_employee = db.Column(db.Float, default=0.00)
    pf_total = db.Column(db.Float, default=0.00)

    # 3. Special Allowances & Duty
    salary_payable = db.Column(db.Float, default=0.00)
    govt_holiday = db.Column(db.Float, default=0.00)
    holiday_extra_duty = db.Column(db.Float, default=0.00)
    loyalty_allowance = db.Column(db.Float, default=0.00)
    additional_duty = db.Column(db.Float, default=0.00)
    office_order_allowance = db.Column(db.Float, default=0.00)
    eid_bonus = db.Column(db.Float, default=0.00)
    other_addition = db.Column(db.Float, default=0.00)

    # 4. Attendance & Deductions
    working_days = db.Column(db.Float, default=0.00)
    absent_days = db.Column(db.Float, default=0.00)
    late_days = db.Column(db.Float, default=0.00)
    absenteeism_amount = db.Column(db.Float, default=0.00)
    lateness_amount = db.Column(db.Float, default=0.00)
    other_deduction = db.Column(db.Float, default=0.00)
    total_deductions = db.Column(db.Float, default=0.00)

    # 5. Settlement & Company Total Paid
    arrear = db.Column(db.Float, default=0.00)
    net_salary = db.Column(db.Float, default=0.00) # Net Pay
    tds = db.Column(db.Float, default=0.00) # TDS
    net_pay_after_tds = db.Column(db.Float, default=0.00) # Net pay after TDS
    company_total_paid = db.Column(db.Float, default=0.00) # Total Company Cost / Paid

    # Compatibility / Legacy fields
    pf_deduction = db.Column(db.Float, default=0.00)
    tax_deduction = db.Column(db.Float, default=0.00)
    other_allowance = db.Column(db.Float, default=0.00)
    payment_method = db.Column(db.String(50), default='Bank Transfer')
    remarks = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship('Employee', backref=db.backref('salary_histories', lazy=True, cascade='all, delete-orphan'))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    @property
    def display_net(self):
        if self.net_pay_after_tds is not None and self.net_pay_after_tds > 0:
            return self.net_pay_after_tds
        return self.net_salary or 0.0

    @property
    def display_company_paid(self):
        if self.company_total_paid is not None and self.company_total_paid > 0:
            return self.company_total_paid
        empr = self.pf_employer or self.pf_deduction or 0.0
        return round(((self.net_pay_after_tds or self.net_salary or 0.0) + (self.tds or self.tax_deduction or 0.0) + empr), 2)

    def __repr__(self):
        return f'<SalaryHistory Emp:{self.employee_id} {self.salary_month} ${self.net_salary}>'


class MonthlyBudget(db.Model):
    __tablename__ = 'monthly_budget'
    id = db.Column(db.Integer, primary_key=True)
    month_year = db.Column(db.String(7), nullable=False, index=True) # e.g. '2026-09'
    title = db.Column(db.String(150), nullable=False)
    fiscal_year = db.Column(db.String(20), nullable=True) # e.g. '2026-2027'
    status = db.Column(db.String(20), default='Issued') # 'Draft', 'Issued', 'Closed'
    total_budgeted = db.Column(db.Float, default=0.00)
    notes = db.Column(db.Text, nullable=True)
    issued_by = db.Column(db.String(100), default='Admin')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Multi-division corporate budget columns (TBN, TIS, ISP)
    tbn_budget = db.Column(db.Float, default=0.00)
    tis_budget = db.Column(db.Float, default=0.00)
    isp_budget = db.Column(db.Float, default=0.00)
    tbn_headcount = db.Column(db.Integer, default=104)
    tis_headcount = db.Column(db.Integer, default=42)
    isp_headcount = db.Column(db.Integer, default=6)
    additional_expenses = db.Column(db.Float, default=0.00)
    additional_expenses_note = db.Column(db.String(255), default='Incentive for Sale (July)')
    gross_total = db.Column(db.Float, default=0.00)
    bill_adjustment = db.Column(db.Float, default=0.00)
    bill_adjustment_note = db.Column(db.String(255), default='Adjusted with ISP Bill Received')
    net_budget = db.Column(db.Float, default=0.00)
    usd_rate = db.Column(db.Float, default=121.75)
    usd_amount = db.Column(db.Float, default=0.00)
    budget_data_json = db.Column(db.Text, nullable=True)

    items = db.relationship('BudgetItem', backref='budget', cascade='all, delete-orphan', lazy=True)

    def __init__(self, month_year="", title="", fiscal_year=None, status="Issued", total_budgeted=0.00, notes="", issued_by="Admin", tbn_budget=0.0, tis_budget=0.0, isp_budget=0.0, tbn_headcount=104, tis_headcount=42, isp_headcount=6, additional_expenses=0.0, additional_expenses_note="", gross_total=0.0, bill_adjustment=0.0, bill_adjustment_note="", net_budget=0.0, usd_rate=121.75, usd_amount=0.0, budget_data_json=None, **kwargs):
        super().__init__(**kwargs)
        self.month_year = month_year
        self.title = title
        self.fiscal_year = fiscal_year
        self.status = status
        self.total_budgeted = total_budgeted
        self.notes = notes
        self.issued_by = issued_by
        self.tbn_budget = tbn_budget
        self.tis_budget = tis_budget
        self.isp_budget = isp_budget
        self.tbn_headcount = tbn_headcount
        self.tis_headcount = tis_headcount
        self.isp_headcount = isp_headcount
        self.additional_expenses = additional_expenses
        self.additional_expenses_note = additional_expenses_note
        self.gross_total = gross_total
        self.bill_adjustment = bill_adjustment
        self.bill_adjustment_note = bill_adjustment_note
        self.net_budget = net_budget
        self.usd_rate = usd_rate
        self.usd_amount = usd_amount
        self.budget_data_json = budget_data_json

    @property
    def month_display(self):
        try:
            parts = self.month_year.split('-')
            dt = datetime(int(parts[0]), int(parts[1]), 1)
            return dt.strftime('%B %Y')
        except Exception:
            return self.month_year

    def __repr__(self):
        return f'<MonthlyBudget {self.month_year} - {self.title}>'


class BudgetItem(db.Model):
    __tablename__ = 'budget_item'
    id = db.Column(db.Integer, primary_key=True)
    budget_id = db.Column(db.Integer, db.ForeignKey('monthly_budget.id'), nullable=False)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=True)
    particulars = db.Column(db.String(150), nullable=True)
    section = db.Column(db.String(50), default='direct') # 'direct', 'proportionate', 'additional'
    category_group = db.Column(db.String(100), nullable=True)
    tbn_amount = db.Column(db.Float, default=0.00)
    tis_amount = db.Column(db.Float, default=0.00)
    isp_amount = db.Column(db.Float, default=0.00)
    allocated_amount = db.Column(db.Float, default=0.00)
    is_proportionate = db.Column(db.Boolean, default=False)
    notes = db.Column(db.String(255), nullable=True)

    account = db.relationship('Account', backref=db.backref('budget_items', lazy=True))

    def __init__(self, budget_id=None, account_id=None, particulars="", section="direct", category_group="", tbn_amount=0.0, tis_amount=0.0, isp_amount=0.0, allocated_amount=0.0, is_proportionate=False, notes="", **kwargs):
        super().__init__(**kwargs)
        self.budget_id = budget_id
        self.account_id = account_id
        self.particulars = particulars
        self.section = section
        self.category_group = category_group
        self.tbn_amount = tbn_amount
        self.tis_amount = tis_amount
        self.isp_amount = isp_amount
        self.allocated_amount = allocated_amount
        self.is_proportionate = is_proportionate
        self.notes = notes

    def __repr__(self):
        return f'<BudgetItem Budget:{self.budget_id} {self.particulars} Total:৳{self.allocated_amount}>'


class RemittanceRecord(db.Model):
    __tablename__ = 'remittance_record'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    particulars = db.Column(db.String(100), nullable=False) # 'Budget Issue', 'Remittance Received', 'Remittance Received-local'
    budget_amount = db.Column(db.Float, default=0.00)
    received_bdt = db.Column(db.Float, default=0.00)
    running_balance = db.Column(db.Float, default=0.00)
    deal_no = db.Column(db.String(50), nullable=True) # or 'Cash'
    invoice_no = db.Column(db.String(50), nullable=True)
    usd_rate = db.Column(db.Float, default=0.00)
    remittance_usd = db.Column(db.Float, default=0.00)
    sent_by = db.Column(db.String(100), nullable=True)
    remittance_date = db.Column(db.Date, nullable=True)
    incentive_10_bdt = db.Column(db.Float, default=0.00)
    incentive_10_audited_bdt = db.Column(db.Float, default=0.00)
    
    # 10 Compliance Checkpoints
    tt_certificate = db.Column(db.Boolean, default=False)
    encashment_certificate = db.Column(db.Boolean, default=False)
    authorized_dealer_cert = db.Column(db.Boolean, default=False)
    cash_incentive_cert = db.Column(db.Boolean, default=False)
    send_to_bank = db.Column(db.Boolean, default=False)
    date_send_to_bank = db.Column(db.Date, nullable=True)
    send_to_head_office = db.Column(db.Boolean, default=False)
    send_to_audit_office = db.Column(db.Boolean, default=False)
    audit_certificate_date = db.Column(db.Date, nullable=True)
    send_to_intl_division = db.Column(db.Boolean, default=False)
    send_to_bangladesh_bank = db.Column(db.Boolean, default=False)
    
    # Benefit Realization
    benefit_receivable = db.Column(db.Float, default=0.00)
    benefit_received = db.Column(db.Float, default=0.00)
    benefit_received_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(100), default='Pending') # 'Clear & Expense in Dec-2023', 'Declined', 'Pending'
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __init__(self, date=None, particulars="", budget_amount=0.0, received_bdt=0.0, running_balance=0.0, deal_no=None, invoice_no=None, usd_rate=0.0, remittance_usd=0.0, sent_by=None, remittance_date=None, incentive_10_bdt=0.0, incentive_10_audited_bdt=0.0, tt_certificate=False, encashment_certificate=False, authorized_dealer_cert=False, cash_incentive_cert=False, send_to_bank=False, date_send_to_bank=None, send_to_head_office=False, send_to_audit_office=False, audit_certificate_date=None, send_to_intl_division=False, send_to_bangladesh_bank=False, benefit_receivable=0.0, benefit_received=0.0, benefit_received_date=None, status="Pending", notes="", **kwargs):
        super().__init__(**kwargs)
        self.date = date if date is not None else datetime.utcnow().date()
        self.particulars = particulars
        self.budget_amount = budget_amount
        self.received_bdt = received_bdt
        self.running_balance = running_balance
        self.deal_no = deal_no
        self.invoice_no = invoice_no
        self.usd_rate = usd_rate
        self.remittance_usd = remittance_usd
        self.sent_by = sent_by
        self.remittance_date = remittance_date
        self.incentive_10_bdt = incentive_10_bdt
        self.incentive_10_audited_bdt = incentive_10_audited_bdt
        self.tt_certificate = tt_certificate
        self.encashment_certificate = encashment_certificate
        self.authorized_dealer_cert = authorized_dealer_cert
        self.cash_incentive_cert = cash_incentive_cert
        self.send_to_bank = send_to_bank
        self.date_send_to_bank = date_send_to_bank
        self.send_to_head_office = send_to_head_office
        self.send_to_audit_office = send_to_audit_office
        self.audit_certificate_date = audit_certificate_date
        self.send_to_intl_division = send_to_intl_division
        self.send_to_bangladesh_bank = send_to_bangladesh_bank
        self.benefit_receivable = benefit_receivable
        self.benefit_received = benefit_received
        self.benefit_received_date = benefit_received_date
        self.status = status
        self.notes = notes

    def __repr__(self):
        return f'<RemittanceRecord {self.date} {self.particulars} Inv:{self.invoice_no}>'


# ==============================================================================
# CONTEXT PROCESSORS & AUTHENTICATION HELPERS
# ==============================================================================

def num_to_words(n):
    units = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
             'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

    def convert_below_thousand(val):
        res = ''
        if val >= 100:
            res += units[val // 100] + ' Hundred '
            val %= 100
        if val >= 20:
            res += tens[val // 10] + ' '
            val %= 10
        if val > 0:
            res += units[val] + ' '
        return res.strip()

    try:
        n = float(n)
    except (ValueError, TypeError):
        return 'Zero Taka Only'

    taka = int(abs(n))
    paisa = int(round((abs(n) - taka) * 100))
    if taka == 0:
        words = 'Zero'
    else:
        parts = []
        crore = taka // 10000000
        taka %= 10000000
        lakh = taka // 100000
        taka %= 100000
        thousand = taka // 1000
        taka %= 1000
        hundred_plus = taka

        if crore > 0:
            parts.append(convert_below_thousand(crore) + ' Crore')
        if lakh > 0:
            parts.append(convert_below_thousand(lakh) + ' Lakh')
        if thousand > 0:
            parts.append(convert_below_thousand(thousand) + ' Thousand')
        if hundred_plus > 0:
            parts.append(convert_below_thousand(hundred_plus))
        words = ' '.join(parts)

    paisa_str = f' and {convert_below_thousand(paisa)} Paisa' if paisa > 0 else ''
    return f'{words} Taka{paisa_str} Only'


@app.context_processor
def inject_globals():
    user = None
    if 'user_id' in session:
        user = User.query.get(session['user_id'])
        if not user:
            try:
                eng1 = CompanyManager.get_engine(1, metadata=db.metadata)
                from sqlalchemy.orm import Session as SASession
                with SASession(bind=eng1) as s1:
                    user = s1.query(User).get(session['user_id'])
            except Exception:
                pass

    active_cid = session.get('active_company_id', 1)
    active_company = CompanyManager.get_by_id(active_cid)
    if not active_company:
        active_company = CompanyManager.get_by_id(1)
        session['active_company_id'] = 1
    all_companies = CompanyManager.get_all()

    return {
        'now': datetime.utcnow(),
        'current_user': user,
        'currency': '৳',
        'num_to_words': num_to_words,
        'active_company': active_company,
        'all_companies': all_companies
    }

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            # Auto-assign admin in dev if exists, or redirect to login
            admin = User.query.filter_by(role='Admin').first()
            if admin and 'dev_auto_auth' in session:
                session['user_id'] = admin.id
            else:
                return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login', next=request.url))
        if session.get('role') == 'Admin':
            return f(*args, **kwargs)
        user = User.query.get(session['user_id'])
        if not user or user.role != 'Admin':
            flash('Access restricted: Administrator privileges required.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# ==============================================================================
# MONTHLY CUSTOMIZED SALARY SHEET & BANK TRANSFER SYSTEM INITIALIZATION
# ==============================================================================
from payroll_salary_manager import init_payroll_salary_routes
init_payroll_salary_routes(app, db, {
    'Employee': Employee,
    'MonthlySalarySheet': MonthlySalarySheet,
    'MonthlySalaryItem': MonthlySalaryItem,
    'Account': Account,
    'JournalEntry': JournalEntry,
    'TransactionLine': TransactionLine,
    'num_to_words': num_to_words
})


# ==============================================================================
# ACCOUNTING CORE LOGIC & HELPERS
# ==============================================================================

def recalculate_account_balances():
    """Recalculates ledger balances for all accounts based on double-entry posted lines."""
    accounts = Account.query.all()
    for acc in accounts:
        dr_sum = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .filter(TransactionLine.account_id == acc.id).scalar()
        cr_sum = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .filter(TransactionLine.account_id == acc.id).scalar()

        if acc.type in ['Asset', 'Expense', 'Contra']:
            acc.balance = round(acc.opening_balance + dr_sum - cr_sum, 2)
        else: # Liability, Equity, Revenue
            acc.balance = round(acc.opening_balance + cr_sum - dr_sum, 2)
    db.session.commit()

def recalculate_cash_book_running_balances():
    """Recalculates sequential running balances in CashBookRow from opening balances to the latest transaction."""
    cash_acc = Account.query.filter_by(code='BS-001').first()
    prem_acc = Account.query.filter_by(code='BS-002').first()
    start_cash = cash_acc.opening_balance if cash_acc else 1676799.00
    start_prem = prem_acc.opening_balance if prem_acc else -279306.00
    start_bank = 0.0

    rows = CashBookRow.query.order_by(
        func.coalesce(CashBookRow.payment_date, CashBookRow.receipt_date).asc(),
        CashBookRow.row_order.asc(),
        CashBookRow.id.asc()
    ).all()

    cur_cash = start_cash
    cur_prem = start_prem
    cur_bank = start_bank

    for idx, r in enumerate(rows):
        r.row_order = idx + 1
        desc = (r.receipt_desc or '').lower()
        if 'opening' in desc or 'opeining' in desc:
            r.balance_cash = round(start_cash, 2)
            r.balance_premier = round(start_prem, 2)
            r.balance_bank = round(start_bank, 2)
            continue

        cur_cash += (r.receipt_cash or 0.0) - (r.payment_cash or 0.0)
        cur_prem += (r.receipt_premier or 0.0) - (r.payment_premier or 0.0)
        cur_bank += (r.receipt_bank or 0.0) - (r.payment_bank or 0.0)

        r.balance_cash = round(cur_cash, 2)
        r.balance_premier = round(cur_prem, 2)
        r.balance_bank = round(cur_bank, 2)

    db.session.flush()

def sync_voucher_to_cash_book(voucher):
    """
    Bi-directional synchronization: Whenever a voucher is created or edited in the General Ledger,
    automatically reflects and updates the Cash & Bank Book.
    """
    if not voucher:
        return

    # Check which lines involve Cash or Bank accounts
    cash_bank_lines = [l for l in voucher.lines if l.account and l.account.is_cash_or_bank]
    if not cash_bank_lines:
        # Non-cash/bank journal voucher (e.g. depreciation) -> ensure no phantom cash book row exists
        existing = CashBookRow.query.filter(
            db.or_(CashBookRow.payment_vno == voucher.voucher_no, CashBookRow.receipt_voucher == voucher.voucher_no)
        ).first()
        if existing:
            db.session.delete(existing)
            recalculate_cash_book_running_balances()
            db.session.commit()
        return

    # Inflows into Cash/Bank
    cash_dr = sum(l.debit for l in voucher.lines if l.account and l.account.code == 'BS-001')
    prem_dr = sum(l.debit for l in voucher.lines if l.account and l.account.code == 'BS-002')
    other_bank_dr = sum(l.debit for l in voucher.lines if l.account and l.account.is_cash_or_bank and l.account.code not in ['BS-001', 'BS-002'])
    total_inflow = cash_dr + prem_dr + other_bank_dr

    # Outflows from Cash/Bank
    cash_cr = sum(l.credit for l in voucher.lines if l.account and l.account.code == 'BS-001')
    prem_cr = sum(l.credit for l in voucher.lines if l.account and l.account.code == 'BS-002')
    other_bank_cr = sum(l.credit for l in voucher.lines if l.account and l.account.is_cash_or_bank and l.account.code not in ['BS-001', 'BS-002'])
    total_outflow = cash_cr + prem_cr + other_bank_cr

    # Counterparty lines (Expense, Revenue, Customer, Vendor, etc.)
    cp_dr_lines = [l for l in voucher.lines if l.account and not l.account.is_cash_or_bank and l.debit > 0]
    cp_cr_lines = [l for l in voucher.lines if l.account and not l.account.is_cash_or_bank and l.credit > 0]

    # Find or initialize CashBookRow
    cb_row = CashBookRow.query.filter(
        db.or_(CashBookRow.payment_vno == voucher.voucher_no, CashBookRow.receipt_voucher == voucher.voucher_no)
    ).first()

    if not cb_row:
        max_order = db.session.query(func.coalesce(func.max(CashBookRow.row_order), 0)).scalar() or 0
        cb_row = CashBookRow(row_order=max_order + 1)
        db.session.add(cb_row)

    # CASE 1: Contra Transaction (Cash transfer between bank and cash)
    if total_inflow > 0 and total_outflow > 0:
        cb_row.receipt_code = "BS-003"
        cb_row.receipt_date = voucher.date
        cb_row.receipt_voucher = voucher.voucher_no
        cb_row.receipt_desc = "Contra Transfer"
        cb_row.receipt_ref = voucher.reference or ""
        cb_row.receipt_cash = round(cash_dr, 2)
        cb_row.receipt_premier = round(prem_dr, 2)
        cb_row.receipt_bank = round(other_bank_dr, 2)
        cb_row.receipt_total = round(total_inflow, 2)

        cb_row.payment_date = voucher.date
        cb_row.payment_vno = voucher.voucher_no
        cb_row.payment_code = "BS-003"
        cb_row.payment_desc = "Contra Transfer"
        cb_row.payment_ref = voucher.reference or ""
        cb_row.payment_cash = round(cash_cr, 2)
        cb_row.payment_premier = round(prem_cr, 2)
        cb_row.payment_bank = round(other_bank_cr, 2)
        cb_row.payment_total = round(total_outflow, 2)
        cb_row.remarks = voucher.narration or ""

    # CASE 2: Payment Voucher (Cash/Bank Outflow)
    elif total_outflow > 0:
        primary_dr = cp_dr_lines[0] if cp_dr_lines else None
        p_code = primary_dr.account.code if primary_dr else ""
        p_desc = primary_dr.particulars or (primary_dr.account.name if primary_dr else "Payment")

        cb_row.payment_date = voucher.date
        cb_row.payment_vno = voucher.voucher_no
        cb_row.payment_code = p_code
        cb_row.payment_desc = p_desc
        cb_row.payment_ref = voucher.reference or ""
        cb_row.payment_cash = round(cash_cr, 2)
        cb_row.payment_premier = round(prem_cr, 2)
        cb_row.payment_bank = round(other_bank_cr, 2)
        cb_row.payment_total = round(total_outflow, 2)

        # Clear receipt side
        cb_row.receipt_code = ""
        cb_row.receipt_voucher = ""
        cb_row.receipt_desc = ""
        cb_row.receipt_ref = ""
        cb_row.receipt_cash = 0.0
        cb_row.receipt_premier = 0.0
        cb_row.receipt_bank = 0.0
        cb_row.receipt_total = 0.0
        cb_row.remarks = voucher.narration or ""

    # CASE 3: Receipt Voucher (Cash/Bank Inflow)
    elif total_inflow > 0:
        primary_cr = cp_cr_lines[0] if cp_cr_lines else None
        r_code = primary_cr.account.code if primary_cr else ""
        r_desc = primary_cr.particulars or (primary_cr.account.name if primary_cr else "Receipt")

        cb_row.receipt_date = voucher.date
        cb_row.receipt_voucher = voucher.voucher_no
        cb_row.receipt_code = r_code
        cb_row.receipt_desc = r_desc
        cb_row.receipt_ref = voucher.reference or ""
        cb_row.receipt_cash = round(cash_dr, 2)
        cb_row.receipt_premier = round(prem_dr, 2)
        cb_row.receipt_bank = round(other_bank_dr, 2)
        cb_row.receipt_total = round(total_inflow, 2)

        # Clear payment side
        cb_row.payment_date = None
        cb_row.payment_vno = ""
        cb_row.payment_code = ""
        cb_row.payment_desc = ""
        cb_row.payment_ref = ""
        cb_row.payment_cash = 0.0
        cb_row.payment_premier = 0.0
        cb_row.payment_bank = 0.0
        cb_row.payment_total = 0.0
        cb_row.remarks = voucher.narration or ""

    db.session.flush()
    recalculate_cash_book_running_balances()
    db.session.commit()

def delete_voucher_from_cash_book(voucher_no):
    """Removes a deleted voucher from the Cash & Bank Book and updates running balances."""
    if not voucher_no:
        return
    rows = CashBookRow.query.filter(
        db.or_(CashBookRow.payment_vno == voucher_no, CashBookRow.receipt_voucher == voucher_no)
    ).all()
    for r in rows:
        db.session.delete(r)
    db.session.flush()
    recalculate_cash_book_running_balances()
    db.session.commit()

def generate_voucher_no(vtype):
    """Generates next guaranteed unique sequential voucher number e.g. PV-2026-0001, RV-2026-0001."""
    prefix_map = {
        'Payment': 'PV',
        'Receipt': 'RV',
        'Journal': 'JV',
        'Contra': 'CV'
    }
    prefix = prefix_map.get(vtype, 'JV')
    current_year = date.today().year

    # Find highest existing number suffix
    matching_vouchers = JournalEntry.query.filter(
        JournalEntry.voucher_no.like(f"{prefix}-{current_year}-%")
    ).all()

    max_num = 0
    for v in matching_vouchers:
        parts = v.voucher_no.split('-')
        if len(parts) >= 3 and parts[-1].isdigit():
            val = int(parts[-1])
            if val > max_num:
                max_num = val

    next_num = max(len(matching_vouchers), max_num) + 1
    candidate = f"{prefix}-{current_year}-{next_num:04d}"

    session_vouchers = {obj.voucher_no for obj in db.session.new if isinstance(obj, JournalEntry)}
    while candidate in session_vouchers or JournalEntry.query.filter_by(voucher_no=candidate).first():
        next_num += 1
        candidate = f"{prefix}-{current_year}-{next_num:04d}"

    return candidate


# ==============================================================================
# ROUTES: AUTHENTICATION & MULTI-USER MANAGEMENT
# ==============================================================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            # Check if user is Administrator requiring OTP 2FA verification
            if user.role == 'Admin' or user.username.lower() in ['admin', 'zillur']:
                otp = generate_otp()
                send_otp_email(ADMIN_EMAIL, otp)
                session['pending_otp'] = {
                    'user_id': user.id,
                    'username': user.username,
                    'role': user.role,
                    'full_name': user.full_name,
                    'otp': otp,
                    'expires_at': (datetime.utcnow() + timedelta(minutes=10)).isoformat(),
                    'next_url': request.args.get('next') or url_for('dashboard')
                }
                flash(f'Admin Security Verification: A 6-digit OTP code has been sent to {ADMIN_EMAIL}.', 'info')
                return redirect(url_for('verify_otp'))
            else:
                session['user_id'] = user.id
                session['username'] = user.username
                session['role'] = user.role
                flash(f'Welcome back, {user.full_name} ({user.role})!', 'success')
                next_url = request.args.get('next') or url_for('dashboard')
                return redirect(next_url)
        else:
            flash('Invalid username or password.', 'danger')

    return render_template('auth/login.html')

@app.route('/verify-otp', methods=['GET', 'POST'])
def verify_otp():
    pending = session.get('pending_otp')
    if not pending:
        flash('No pending authentication session found. Please sign in.', 'warning')
        return redirect(url_for('login'))

    if request.method == 'POST':
        submitted_otp = request.form.get('otp', '').strip()
        expires_at = datetime.fromisoformat(pending['expires_at'])

        if datetime.utcnow() > expires_at:
            flash('The verification code has expired. Please request a new code.', 'danger')
            return redirect(url_for('verify_otp'))

        if submitted_otp == pending['otp']:
            # Successful verification
            session['user_id'] = pending['user_id']
            session['username'] = pending['username']
            session['role'] = pending['role']
            next_url = pending.get('next_url') or url_for('dashboard')
            full_name = pending.get('full_name', pending['username'])
            session.pop('pending_otp', None)
            flash(f'Welcome back, {full_name}! Administrator identity verified via {ADMIN_EMAIL}.', 'success')
            return redirect(next_url)
        else:
            flash('Invalid verification code. Please enter the correct 6-digit OTP.', 'danger')

    return render_template(
        'auth/verify_otp.html',
        target_email=ADMIN_EMAIL,
        dev_otp=pending.get('otp')
    )

@app.route('/resend-otp')
def resend_otp():
    pending = session.get('pending_otp')
    if not pending:
        flash('No active sign-in attempt found.', 'warning')
        return redirect(url_for('login'))

    new_otp = generate_otp()
    pending['otp'] = new_otp
    pending['expires_at'] = (datetime.utcnow() + timedelta(minutes=10)).isoformat()
    session['pending_otp'] = pending

    send_otp_email(ADMIN_EMAIL, new_otp)
    flash(f'A fresh 6-digit OTP has been sent to {ADMIN_EMAIL}.', 'info')
    return redirect(url_for('verify_otp'))

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out safely.', 'info')
    return redirect(url_for('login'))

@app.route('/settings/users')
@login_required
@admin_required
def user_settings():
    users = User.query.order_by(User.id).all()
    return render_template('settings/users.html', users=users)

@app.route('/settings/users/new', methods=['POST'])
@login_required
@admin_required
def create_user():
    username = request.form.get('username', '').strip()
    full_name = request.form.get('full_name', '').strip()
    role = request.form.get('role', 'General User')
    password = request.form.get('password', '')

    if not username or not password:
        flash('Username and password are required.', 'danger')
        return redirect(url_for('user_settings'))

    existing = User.query.filter_by(username=username).first()
    if existing:
        flash(f'User with username "{username}" already exists.', 'danger')
        return redirect(url_for('user_settings'))

    new_user = User(username=username, password=password, full_name=full_name, role=role)
    db.session.add(new_user)
    db.session.commit()

    flash(f'User "{username}" ({role}) created successfully.', 'success')
    return redirect(url_for('user_settings'))

@app.route('/settings/users/<int:user_id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_user(user_id):
    if session.get('user_id') == user_id:
        flash('You cannot delete your own logged-in user account.', 'danger')
        return redirect(url_for('user_settings'))

    user = User.query.get_or_404(user_id)
    uname = user.username
    db.session.delete(user)
    db.session.commit()

    flash(f'User "{uname}" deleted successfully.', 'info')
    return redirect(url_for('user_settings'))


# ==============================================================================
# ROUTES: FRONT PAGE PORTAL & DASHBOARD
# ==============================================================================

@app.route('/')
def home():
    # If not logged in, auto-authenticate default admin for smooth pairing experience
    if 'user_id' not in session:
        admin = User.query.filter_by(role='Admin').first()
        if admin:
            session['user_id'] = admin.id
            session['username'] = admin.username
            session['role'] = admin.role

    total_accounts = Account.query.count()
    total_vouchers = JournalEntry.query.count()
    total_employees = Employee.query.count()
    active_budgets = MonthlyBudget.query.count()
    total_remittances = RemittanceRecord.query.count()

    accounts = Account.query.all()
    total_assets = sum(a.balance for a in accounts if a.type == 'Asset')
    total_liabilities = sum(a.balance for a in accounts if a.type == 'Liability')
    total_revenue = sum(a.balance for a in accounts if a.type == 'Revenue')
    total_expense = sum(a.balance for a in accounts if a.type == 'Expense')
    net_surplus = total_revenue - total_expense
    liquid_cash = sum(a.balance for a in accounts if a.is_cash_or_bank)

    active_cid = session.get('active_company_id', 1)
    active_comp = CompanyManager.get_by_id(active_cid)

    return render_template(
        'home.html',
        total_accounts=total_accounts,
        total_vouchers=total_vouchers,
        total_employees=total_employees,
        active_budgets=active_budgets,
        total_remittances=total_remittances,
        total_assets=total_assets,
        total_liabilities=total_liabilities,
        total_revenue=total_revenue,
        total_expense=total_expense,
        net_surplus=net_surplus,
        liquid_cash=liquid_cash,
        active_company=active_comp,
        today_date=date.today()
    )


@app.route('/dashboard')
def dashboard():
    # If not logged in, auto-authenticate default admin for smooth pairing experience
    if 'user_id' not in session:
        admin = User.query.filter_by(role='Admin').first()
        if admin:
            session['user_id'] = admin.id
            session['username'] = admin.username
            session['role'] = admin.role

    recalculate_account_balances()
    accounts = Account.query.all()

    # Calculate high-level financial indicators
    total_assets = sum(a.balance for a in accounts if a.type == 'Asset')
    total_liabilities = sum(a.balance for a in accounts if a.type == 'Liability')
    total_equity_accounts = sum(a.balance for a in accounts if a.type == 'Equity')
    total_revenue = sum(a.balance for a in accounts if a.type == 'Revenue')
    total_expense = sum(a.balance for a in accounts if a.type == 'Expense')
    net_surplus = round(total_revenue - total_expense, 2)
    total_equity = round(total_equity_accounts + net_surplus, 2)

    # Liquid Cash & Bank Position
    liquid_accounts = [a for a in accounts if a.is_cash_or_bank]
    total_liquid = sum(a.balance for a in liquid_accounts)

    recent_vouchers = JournalEntry.query.order_by(JournalEntry.date.desc(), JournalEntry.id.desc()).limit(8).all()

    # Compute Monthly Cashflow & Voucher Activity using efficient SQL aggregation
    from collections import defaultdict, Counter
    import json

    monthly_rows = db.session.query(
        func.strftime('%Y-%m', JournalEntry.date).label('month'),
        JournalEntry.voucher_type,
        func.sum(TransactionLine.debit).label('total'),
        func.count(func.distinct(JournalEntry.id)).label('vcount')
    ).join(TransactionLine, JournalEntry.id == TransactionLine.journal_entry_id)\
     .filter(JournalEntry.date.isnot(None))\
     .group_by('month', JournalEntry.voucher_type)\
     .order_by('month').all()

    months_dict = defaultdict(lambda: {'receipts': 0.0, 'payments': 0.0, 'contra': 0.0, 'vcount': 0})
    for m, vtype, tot, vcnt in monthly_rows:
        if not m:
            continue
        months_dict[m]['vcount'] += vcnt
        if vtype == 'Receipt':
            months_dict[m]['receipts'] += float(tot)
        elif vtype == 'Payment':
            months_dict[m]['payments'] += float(tot)
        elif vtype == 'Contra':
            months_dict[m]['contra'] += float(tot)

    sorted_months = sorted(months_dict.keys())
    month_names = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
        '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
        '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }
    def fmt_month(ym):
        parts = ym.split('-')
        return f"{month_names.get(parts[1], parts[1])} {parts[0]}" if len(parts) == 2 else ym

    chart_months = [fmt_month(m) for m in sorted_months]
    chart_receipts = [round(months_dict[m]['receipts'], 2) for m in sorted_months]
    chart_payments = [round(months_dict[m]['payments'], 2) for m in sorted_months]
    chart_net = [round(months_dict[m]['receipts'] - months_dict[m]['payments'], 2) for m in sorted_months]
    chart_vcounts = [months_dict[m]['vcount'] for m in sorted_months]

    # Top Expense Cost Centers
    exp_accounts = sorted([a for a in accounts if a.type == 'Expense' and a.balance > 0], key=lambda x: x.balance, reverse=True)
    total_exp_calc = sum(a.balance for a in exp_accounts) or 1.0
    top_expenses = []
    other_exp = 0.0
    for idx, a in enumerate(exp_accounts):
        if idx < 6:
            top_expenses.append({
                'name': a.name,
                'code': a.code,
                'amount': round(a.balance, 2),
                'pct': round((a.balance / total_exp_calc) * 100, 1)
            })
        else:
            other_exp += a.balance
    if other_exp > 0:
        top_expenses.append({
            'name': 'Other Operating Expenses',
            'code': 'EXP-OTH',
            'amount': round(other_exp, 2),
            'pct': round((other_exp / total_exp_calc) * 100, 1)
        })

    # Asset & Liquid Composition
    pos_assets = sorted([a for a in accounts if a.type == 'Asset' and a.balance > 0], key=lambda x: x.balance, reverse=True)
    total_pos_assets = sum(a.balance for a in pos_assets) or 1.0
    asset_breakdown = [{
        'name': a.name,
        'code': a.code,
        'amount': round(a.balance, 2),
        'pct': round((a.balance / total_pos_assets) * 100, 1),
        'is_liquid': a.is_cash_or_bank
    } for a in pos_assets[:7]]

    # Voucher Type Totals
    pv_total = sum(chart_payments)
    rv_total = sum(chart_receipts)
    cv_total = sum(months_dict[m]['contra'] for m in sorted_months)
    voucher_stats = [
        {'type': 'Payment Vouchers (PV)', 'code': 'PV', 'amount': pv_total, 'color': '#ef4444', 'count': sum(1 for v in recent_vouchers if v.voucher_type == 'Payment')},
        {'type': 'Receipt Vouchers (RV)', 'code': 'RV', 'amount': rv_total, 'color': '#10b981', 'count': sum(1 for v in recent_vouchers if v.voucher_type == 'Receipt')},
        {'type': 'Contra Vouchers (CV)', 'code': 'CV', 'amount': cv_total, 'color': '#f59e0b', 'count': sum(1 for v in recent_vouchers if v.voucher_type == 'Contra')},
    ]

    # Departmental Employee Count & Payroll
    dept_counter = Counter(e.department for e in Employee.query.all() if e.department).most_common(6)
    dept_distribution = [{'dept': d, 'count': c} for d, c in dept_counter]

    analytics_package = {
        'months': chart_months,
        'receipts': chart_receipts,
        'payments': chart_payments,
        'net': chart_net,
        'voucher_counts': chart_vcounts,
        'top_expenses': top_expenses,
        'asset_breakdown': asset_breakdown,
        'voucher_stats': voucher_stats,
        'dept_distribution': dept_distribution,
        'liquid_accounts': [{'name': a.name, 'code': a.code, 'balance': a.balance} for a in liquid_accounts],
        'totals': {
            'assets': total_assets,
            'liabilities': total_liabilities,
            'equity': total_equity,
            'revenue': total_revenue,
            'expense': total_expense,
            'net_surplus': net_surplus,
            'liquid': total_liquid
        }
    }

    analytics_json = json.dumps(analytics_package)

    # Bank Statement Balances (Each shown in a separate box on the dashboard)
    bank_statement_boxes = []
    all_banks = Account.query.filter(
        db.or_(
            Account.category == 'Bank',
            Account.name.ilike('%bank%'),
            Account.is_cash_or_bank == True
        ),
        Account.type.notin_(['Expense', 'Revenue'])
    ).order_by(Account.code.asc()).all()

    # Also include any accounts that have uploaded bank statements
    stmt_acc_ids = [r[0] for r in db.session.query(BankStatement.bank_account_id).distinct().all() if r[0]]
    if stmt_acc_ids:
        more_accs = Account.query.filter(Account.id.in_(stmt_acc_ids)).all()
        for ma in more_accs:
            if ma not in all_banks:
                all_banks.append(ma)

    seen_acc_ids = set()
    for b_acc in all_banks:
        if b_acc.id in seen_acc_ids:
            continue
        seen_acc_ids.add(b_acc.id)

        latest_line = BankStatementLine.query.filter_by(bank_account_id=b_acc.id).order_by(
            BankStatementLine.txn_date.desc(),
            BankStatementLine.id.desc()
        ).first()
        stmt_count = BankStatementLine.query.filter_by(bank_account_id=b_acc.id).count()
        latest_stmt = BankStatement.query.filter_by(bank_account_id=b_acc.id).order_by(BankStatement.to_date.desc()).first()

        stmt_bal = latest_line.balance if latest_line else (latest_stmt.closing_balance if latest_stmt else None)
        as_of = latest_line.txn_date if latest_line else (latest_stmt.to_date if latest_stmt else None)

        gl_bal = b_acc.balance
        variance = (stmt_bal - gl_bal) if stmt_bal is not None else None

        bank_statement_boxes.append({
            'account_id': b_acc.id,
            'name': b_acc.name,
            'code': b_acc.code,
            'has_statement': stmt_count > 0,
            'statement_balance': stmt_bal,
            'book_balance': gl_bal,
            'variance': variance,
            'as_of_date': as_of,
            'line_count': stmt_count,
            'statement_title': latest_stmt.statement_title if latest_stmt else ""
        })

    return render_template(
        'dashboard.html',
        total_assets=total_assets,
        total_liabilities=total_liabilities,
        total_equity=total_equity,
        total_revenue=total_revenue,
        total_expense=total_expense,
        net_surplus=net_surplus,
        liquid_accounts=liquid_accounts,
        total_liquid=total_liquid,
        recent_vouchers=recent_vouchers,
        analytics_json=analytics_json,
        analytics=analytics_package,
        bank_statement_boxes=bank_statement_boxes
    )


# ==============================================================================
# ROUTES: VOUCHER POSTING, EDITING, DELETING & REGISTER
# ==============================================================================

@app.route('/vouchers')
def voucher_list():
    vtype = request.args.get('type')
    q = request.args.get('q', '').strip()

    query = JournalEntry.query
    if vtype:
        query = query.filter(JournalEntry.voucher_type == vtype)
    if q:
        query = query.filter(
            (JournalEntry.voucher_no.ilike(f'%{q}%')) |
            (JournalEntry.narration.ilike(f'%{q}%')) |
            (JournalEntry.reference.ilike(f'%{q}%'))
        )

    vouchers = query.order_by(JournalEntry.date.desc(), JournalEntry.id.desc()).all()
    return render_template('vouchers/voucher_list.html', vouchers=vouchers, selected_type=vtype, search_query=q)

@app.route('/vouchers/new', methods=['GET', 'POST'])
def create_voucher():
    if request.method == 'POST':
        vtype = request.form.get('voucher_type')
        vno = request.form.get('voucher_no')
        vdate_str = request.form.get('date')
        reference = request.form.get('reference')
        narration = request.form.get('narration')

        account_ids = request.form.getlist('account_id[]')
        particulars_list = request.form.getlist('particulars[]')
        debits = request.form.getlist('debit[]')
        credits = request.form.getlist('credit[]')

        try:
            vdate = datetime.strptime(vdate_str, '%Y-%m-%d').date()
        except (ValueError, TypeError):
            vdate = date.today()

        valid_lines = []
        total_dr = 0.0
        total_cr = 0.0

        for i in range(len(account_ids)):
            acc_id_str = account_ids[i] if i < len(account_ids) else None
            if not acc_id_str:
                continue

            part = particulars_list[i] if i < len(particulars_list) else ''
            dr = float(debits[i]) if i < len(debits) and debits[i] else 0.0
            cr = float(credits[i]) if i < len(credits) and credits[i] else 0.0

            if dr > 0 or cr > 0:
                valid_lines.append({
                    'account_id': int(acc_id_str),
                    'particulars': part,
                    'debit': round(dr, 2),
                    'credit': round(cr, 2)
                })
                total_dr += dr
                total_cr += cr

        total_dr = round(total_dr, 2)
        total_cr = round(total_cr, 2)

        # Validation: Double Entry Principle
        if len(valid_lines) < 2:
            flash('A voucher requires at least two transaction lines (Double-Entry).', 'danger')
            return redirect(url_for('create_voucher', type=vtype))

        if abs(total_dr - total_cr) >= 0.01:
            flash(f'Voucher is out of balance! Total Debits (${total_dr:,.2f}) must equal Total Credits (${total_cr:,.2f}).', 'danger')
            return redirect(url_for('create_voucher', type=vtype))

        if total_dr <= 0:
            flash('Voucher total amount must be greater than zero.', 'danger')
            return redirect(url_for('create_voucher', type=vtype))

        existing = JournalEntry.query.filter_by(voucher_no=vno).first()
        if existing:
            vno = generate_voucher_no(vtype)

        attachment_filename = None
        orig_filename = None
        if 'attachment' in request.files:
            file = request.files['attachment']
            if file and file.filename:
                orig_filename = file.filename
                ext = os.path.splitext(file.filename)[1].lower()
                clean_name = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', os.path.splitext(file.filename)[0])
                ts = int(datetime.utcnow().timestamp())
                attachment_filename = f"v_{vno}_{ts}_{clean_name[:30]}{ext}"
                upload_dir = os.path.join(basedir, 'static', 'uploads', 'vouchers')
                os.makedirs(upload_dir, exist_ok=True)
                file.save(os.path.join(upload_dir, attachment_filename))

        entry = JournalEntry(
            voucher_no=vno,
            voucher_type=vtype,
            date=vdate,
            reference=reference,
            narration=narration,
            total_amount=total_dr,
            attachment=attachment_filename,
            original_filename=orig_filename
        )
        db.session.add(entry)
        db.session.flush()

        for item in valid_lines:
            line = TransactionLine(
                journal_entry_id=entry.id,
                account_id=item['account_id'],
                debit=item['debit'],
                credit=item['credit'],
                particulars=item['particulars']
            )
            db.session.add(line)

        db.session.commit()
        recalculate_account_balances()
        sync_voucher_to_cash_book(entry)

        flash(f'Voucher {vno} ({vtype}) has been successfully posted to the General Ledger and synchronized with the Cash & Bank Book!', 'success')
        return redirect(url_for('view_voucher', voucher_id=entry.id))

    # GET Request
    selected_type = request.args.get('type', 'Payment')
    suggested_no = generate_voucher_no(selected_type)
    accounts = Account.query.order_by(Account.code).all()
    today_str = date.today().strftime('%Y-%m-%d')

    return render_template(
        'vouchers/create_voucher.html',
        selected_type=selected_type,
        suggested_no=suggested_no,
        accounts=accounts,
        today=today_str
    )

@app.route('/vouchers/<voucher_id>')
def view_voucher(voucher_id):
    if str(voucher_id).isdigit():
        voucher = JournalEntry.query.get(int(voucher_id))
    else:
        voucher = JournalEntry.query.filter_by(voucher_no=str(voucher_id)).first()
    
    if not voucher:
        flash(f'Voucher "{voucher_id}" not found in General Ledger.', 'warning')
        return redirect(url_for('voucher_list'))

    total_debit = sum(line.debit for line in voucher.lines)
    total_credit = sum(line.credit for line in voucher.lines)

    return render_template(
        'vouchers/view_voucher.html',
        voucher=voucher,
        total_debit=total_debit,
        total_credit=total_credit
    )

@app.route('/vouchers/<int:voucher_id>/edit', methods=['GET', 'POST'])
def edit_voucher(voucher_id):
    voucher = JournalEntry.query.get_or_404(voucher_id)
    accounts = Account.query.order_by(Account.code).all()

    if request.method == 'POST':
        vtype = request.form.get('voucher_type')
        vno = request.form.get('voucher_no')
        vdate_str = request.form.get('date')
        reference = request.form.get('reference')
        narration = request.form.get('narration')

        account_ids = request.form.getlist('account_id[]')
        particulars_list = request.form.getlist('particulars[]')
        debits = request.form.getlist('debit[]')
        credits = request.form.getlist('credit[]')

        try:
            vdate = datetime.strptime(vdate_str, '%Y-%m-%d').date()
        except (ValueError, TypeError):
            vdate = voucher.date

        valid_lines = []
        total_dr = 0.0
        total_cr = 0.0

        for i in range(len(account_ids)):
            acc_id_str = account_ids[i] if i < len(account_ids) else None
            if not acc_id_str:
                continue

            part = particulars_list[i] if i < len(particulars_list) else ''
            dr = float(debits[i]) if i < len(debits) and debits[i] else 0.0
            cr = float(credits[i]) if i < len(credits) and credits[i] else 0.0

            if dr > 0 or cr > 0:
                valid_lines.append({
                    'account_id': int(acc_id_str),
                    'particulars': part,
                    'debit': round(dr, 2),
                    'credit': round(cr, 2)
                })
                total_dr += dr
                total_cr += cr

        total_dr = round(total_dr, 2)
        total_cr = round(total_cr, 2)

        if len(valid_lines) < 2:
            flash('A voucher requires at least two transaction lines (Double-Entry).', 'danger')
            return render_template('vouchers/edit_voucher.html', voucher=voucher, accounts=accounts)

        if abs(total_dr - total_cr) >= 0.01:
            flash(f'Voucher is out of balance! Total Debits (${total_dr:,.2f}) must equal Total Credits (${total_cr:,.2f}).', 'danger')
            return render_template('vouchers/edit_voucher.html', voucher=voucher, accounts=accounts)

        # Update header
        voucher.voucher_no = vno
        voucher.voucher_type = vtype
        voucher.date = vdate
        voucher.reference = reference
        voucher.narration = narration
        voucher.total_amount = total_dr

        if request.form.get('remove_attachment') == '1':
            voucher.attachment = None
            voucher.original_filename = None

        if 'attachment' in request.files:
            file = request.files['attachment']
            if file and file.filename:
                orig_filename = file.filename
                ext = os.path.splitext(file.filename)[1].lower()
                clean_name = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', os.path.splitext(file.filename)[0])
                ts = int(datetime.utcnow().timestamp())
                attachment_filename = f"v_{vno}_{ts}_{clean_name[:30]}{ext}"
                upload_dir = os.path.join(basedir, 'static', 'uploads', 'vouchers')
                os.makedirs(upload_dir, exist_ok=True)
                file.save(os.path.join(upload_dir, attachment_filename))
                voucher.attachment = attachment_filename
                voucher.original_filename = orig_filename

        # Delete old lines and re-create updated lines
        TransactionLine.query.filter_by(journal_entry_id=voucher.id).delete()
        for item in valid_lines:
            line = TransactionLine(
                journal_entry_id=voucher.id,
                account_id=item['account_id'],
                debit=item['debit'],
                credit=item['credit'],
                particulars=item['particulars']
            )
            db.session.add(line)

        db.session.commit()
        recalculate_account_balances()
        sync_voucher_to_cash_book(voucher)

        flash(f'Voucher {voucher.voucher_no} has been successfully updated and synchronized with the Cash & Bank Book.', 'success')
        return redirect(url_for('view_voucher', voucher_id=voucher.id))

    return render_template('vouchers/edit_voucher.html', voucher=voucher, accounts=accounts)

@app.route('/vouchers/<int:voucher_id>/delete', methods=['POST'])
def delete_voucher(voucher_id):
    voucher = JournalEntry.query.get_or_404(voucher_id)
    vno = voucher.voucher_no

    # Unlink any payroll record referencing this voucher
    payroll = PayrollRecord.query.filter_by(voucher_id=voucher.id).first()
    if payroll:
        payroll.voucher_id = None

    db.session.delete(voucher)
    db.session.commit()
    delete_voucher_from_cash_book(vno)
    recalculate_account_balances()

    flash(f'Voucher {vno} was permanently deleted and removed from the Cash & Bank Book.', 'info')
    return redirect(url_for('voucher_list'))


# ==============================================================================
# ROUTES: CHART OF ACCOUNTS & SUBLEDGERS (EDIT, DELETE, HIERARCHY)
# ==============================================================================

@app.route('/accounts')
def chart_of_accounts():
    recalculate_account_balances()
    accounts = Account.query.order_by(Account.code).all()
    return render_template('accounts/chart_of_accounts.html', accounts=accounts)

@app.route('/add_account', methods=['GET', 'POST'])
def add_account():
    candidate_parents = Account.query.filter_by(parent_id=None).order_by(Account.code).all()

    if request.method == 'POST':
        code = request.form.get('code')
        name = request.form.get('name')
        acc_type = request.form.get('type')
        category = request.form.get('category')
        parent_id_str = request.form.get('parent_id')
        parent_id = int(parent_id_str) if parent_id_str else None
        is_cash_bank = True if request.form.get('is_cash_or_bank') else False
        opening_bal = float(request.form.get('opening_balance') or 0.0)

        if not code:
            max_id = db.session.query(func.max(Account.id)).scalar() or 0
            prefix = {'Asset': '1', 'Liability': '2', 'Equity': '3', 'Revenue': '4', 'Expense': '5'}.get(acc_type, '9')
            code = f"{prefix}{max_id + 10:03d}"

        if not category:
            category = f"{acc_type} General"

        new_acc = Account(
            code=code,
            name=name,
            type=acc_type,
            category=category,
            parent_id=parent_id,
            is_cash_or_bank=is_cash_bank,
            opening_balance=opening_bal,
            balance=opening_bal
        )
        try:
            db.session.add(new_acc)
            db.session.commit()
            flash(f'Account "{name}" [{code}] created successfully.', 'success')
            return redirect(url_for('chart_of_accounts'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error creating account: {str(e)}', 'danger')

    return render_template('accounts/add_account.html', candidate_parents=candidate_parents)

@app.route('/accounts/<int:account_id>/edit', methods=['GET', 'POST'])
def edit_account(account_id):
    account = Account.query.get_or_404(account_id)
    candidate_parents = Account.query.filter(Account.id != account.id, Account.parent_id == None).order_by(Account.code).all()

    if request.method == 'POST':
        code = request.form.get('code')
        name = request.form.get('name')
        acc_type = request.form.get('type')
        category = request.form.get('category')
        parent_id_str = request.form.get('parent_id')
        parent_id = int(parent_id_str) if parent_id_str else None
        is_cash_bank = True if request.form.get('is_cash_or_bank') else False
        opening_bal = float(request.form.get('opening_balance') or 0.0)

        account.code = code
        account.name = name
        account.type = acc_type
        account.category = category
        account.parent_id = parent_id
        account.is_cash_or_bank = is_cash_bank
        account.opening_balance = opening_bal

        db.session.commit()
        recalculate_account_balances()

        flash(f'Account [{account.code}] {account.name} updated successfully.', 'success')
        return redirect(url_for('chart_of_accounts'))

    return render_template('accounts/edit_account.html', account=account, candidate_parents=candidate_parents)

@app.route('/accounts/<int:account_id>/delete', methods=['POST'])
def delete_account(account_id):
    account = Account.query.get_or_404(account_id)

    # Protection: check if account has transaction lines
    line_count = TransactionLine.query.filter_by(account_id=account.id).count()
    if line_count > 0:
        flash(f'Cannot delete account [{account.code}] {account.name}: It is referenced in {line_count} journal transaction line(s).', 'danger')
        return redirect(url_for('chart_of_accounts'))

    # Protection: check if account has subledgers
    sub_count = Account.query.filter_by(parent_id=account.id).count()
    if sub_count > 0:
        flash(f'Cannot delete account [{account.code}] {account.name}: It has {sub_count} active subledger(s) assigned.', 'danger')
        return redirect(url_for('chart_of_accounts'))

    acc_name = account.name
    db.session.delete(account)
    db.session.commit()
    recalculate_account_balances()

    flash(f'Account "{acc_name}" deleted successfully.', 'info')
    return redirect(url_for('chart_of_accounts'))

@app.route('/accounts/clear-ledger-data', methods=['POST'])
def clear_ledger_data_by_date():
    """
    Clears all journal entries and transaction lines across all accounts within a specified date condition (From Date to To Date).
    Optionally also clears corresponding CashBookRow records within that date range.
    """
    from_str = request.form.get('from_date', '').strip()
    to_str = request.form.get('to_date', '').strip()
    clear_cash_book = True if request.form.get('clear_cash_book') == '1' else False

    try:
        from_date = parse_accounting_date(from_str) or datetime.strptime(from_str, '%Y-%m-%d').date()
        to_date = parse_accounting_date(to_str) or datetime.strptime(to_str, '%Y-%m-%d').date()
        if isinstance(from_date, datetime):
            from_date = from_date.date()
        if isinstance(to_date, datetime):
            to_date = to_date.date()
    except Exception:
        flash('Invalid date range specified. Please provide valid From Date and To Date.', 'danger')
        return redirect(url_for('chart_of_accounts'))

    if from_date > to_date:
        flash('From Date cannot be later than To Date.', 'warning')
        return redirect(url_for('chart_of_accounts'))

    try:
        # Step 1: Find all Journal Entries within date range
        target_jes = JournalEntry.query.filter(
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        ).all()
        je_ids = [j.id for j in target_jes]
        cleared_lines = 0
        cleared_jes = len(je_ids)

        if je_ids:
            cleared_lines = TransactionLine.query.filter(TransactionLine.journal_entry_id.in_(je_ids)).delete(synchronize_session=False)
            JournalEntry.query.filter(JournalEntry.id.in_(je_ids)).delete(synchronize_session=False)

        cleared_cb_rows = 0
        if clear_cash_book:
            cb_query = CashBookRow.query.filter(
                db.or_(
                    db.and_(CashBookRow.payment_date >= from_date, CashBookRow.payment_date <= to_date),
                    db.and_(CashBookRow.receipt_date >= from_date, CashBookRow.receipt_date <= to_date)
                )
            )
            # Never delete row 1 (Opening Balance)
            cb_query = cb_query.filter(CashBookRow.id != 1)
            cleared_cb_rows = cb_query.delete(synchronize_session=False)

        db.session.commit()
        recalculate_account_balances()
        if clear_cash_book:
            recalculate_cash_book_running_balances()

        flash(
            f'Ledger clear completed successfully for period {from_date} to {to_date}: '
            f'Deleted {cleared_jes} voucher(s), {cleared_lines} transaction line(s)'
            + (f', and {cleared_cb_rows} cash book row(s).' if clear_cash_book else '.')
            + ' All ledger account balances have been automatically recalculated.',
            'success'
        )
    except Exception as e:
        db.session.rollback()
        flash(f'Error clearing ledger data: {str(e)}', 'danger')

    return redirect(url_for('chart_of_accounts'))

@app.route('/accounts/sync-all-to-gl', methods=['GET', 'POST'])
def sync_all_cash_book_to_gl():
    """
    Synchronizes all CashBookRow transactions into General Ledger JournalEntries and TransactionLines.
    Ensures 100% parity between Cash & Bank Book and the General Ledger / Dashboard.
    """
    try:
        cash_acc = Account.query.filter_by(code='BS-001').first()
        prem_acc = Account.query.filter_by(code='BS-002').first()
        tbn_acc = Account.query.filter_by(code='BS-035').first()
        acc_dict = {a.code: a for a in Account.query.all()}

        rows = CashBookRow.query.order_by(CashBookRow.row_order.asc(), CashBookRow.id.asc()).all()
        synced_count = 0

        for r in rows:
            desc = (r.receipt_desc or '').lower()
            is_open = 'open' in desc or 'opein' in desc
            if is_open:
                continue

            r_code = str(r.receipt_code or '').strip()
            p_code = str(r.payment_code or '').strip()
            r_cash = clean_accounting_num(r.receipt_cash)
            r_prem = clean_accounting_num(r.receipt_premier)
            r_total = clean_accounting_num(r.receipt_total) or (r_cash + r_prem)
            p_cash = clean_accounting_num(r.payment_cash)
            p_prem = clean_accounting_num(r.payment_premier)
            p_total = clean_accounting_num(r.payment_total) or (p_cash + p_prem)
            r_date = r.receipt_date or date(2026, 1, 1)
            p_date = r.payment_date or date(2026, 1, 1)
            remarks = str(r.remarks or '').strip()

            # 1. Contra Transaction
            if r_code == 'BS-003' or p_code == 'BS-003':
                c_date = r_date or p_date
                amt = round(p_total or r_total, 2)
                existing_cv = JournalEntry.query.filter_by(voucher_type='Contra', date=c_date, total_amount=amt).first()
                if not existing_cv:
                    cv_num = r.payment_vno or r.receipt_voucher
                    if not cv_num or cv_num in ['-', '', '#N/A'] or JournalEntry.query.filter_by(voucher_no=cv_num).first():
                        cv_num = generate_voucher_no('Contra')
                    je = JournalEntry(
                        voucher_no=cv_num,
                        voucher_type='Contra',
                        date=c_date,
                        reference='Cash Contra',
                        narration=remarks or 'Cash Contra transfer',
                        total_amount=amt
                    )
                    db.session.add(je)
                    db.session.flush()
                    if r_cash > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars='Cash received'))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=0.0, credit=round(r_cash, 2), particulars='Premier Bank withdrawal'))
                    elif r_prem > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars='Premier Bank received'))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=0.0, credit=round(r_prem, 2), particulars='Cash disbursed'))
                    synced_count += 1

                # If there was ALSO a payment on that line
                if p_code and p_code != 'BS-003' and p_total > 0:
                    p_key = p_code if p_code not in ['-', '', '#N/A'] else 'IS-016'
                    target_acc = acc_dict.get(p_key) or acc_dict.get('IS-016')
                    p_vno = r.payment_vno
                    existing_pv = None
                    if p_vno and p_vno not in ['-', '', '#N/A']:
                        existing_pv = JournalEntry.query.filter_by(voucher_no=p_vno).first()
                    if not existing_pv and target_acc:
                        existing_pv = JournalEntry.query.join(TransactionLine).filter(
                            JournalEntry.date == p_date,
                            TransactionLine.account_id == target_acc.id,
                            TransactionLine.debit == round(p_total, 2)
                        ).first()
                    if not existing_pv and target_acc:
                        pv_num = p_vno if p_vno and p_vno not in ['-', '', '#N/A'] else None
                        if not pv_num or JournalEntry.query.filter_by(voucher_no=pv_num).first():
                            pv_num = generate_voucher_no('Payment')
                        je_p = JournalEntry(
                            voucher_no=pv_num,
                            voucher_type='Payment',
                            date=p_date,
                            reference=r.payment_ref or None,
                            narration=remarks or f'Payment for {r.payment_desc or target_acc.name}',
                            total_amount=round(p_total, 2)
                        )
                        db.session.add(je_p)
                        db.session.flush()
                        db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=r.payment_desc or 'Expense'))
                        if p_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars='Paid in Cash'))
                        if p_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars='Paid via Premier Bank'))
                        synced_count += 1
                continue

            # 2. Regular Receipt
            if r_total > 0:
                r_key = r_code if r_code and r_code not in ['-', '', '#N/A'] else 'BS-035'
                source_acc = acc_dict.get(r_key) or tbn_acc
                r_vno = r.receipt_voucher
                existing_rv = None
                if r_vno and r_vno not in ['-', '', '#N/A']:
                    existing_rv = JournalEntry.query.filter_by(voucher_no=r_vno).first()
                if not existing_rv and source_acc:
                    existing_rv = JournalEntry.query.filter_by(voucher_type='Receipt', date=r_date, total_amount=round(r_total, 2)).first()

                if not existing_rv and source_acc:
                    rv_num = r_vno if r_vno and r_vno not in ['-', '', '#N/A'] else None
                    if not rv_num or JournalEntry.query.filter_by(voucher_no=rv_num).first():
                        rv_num = generate_voucher_no('Receipt')
                    je_r = JournalEntry(
                        voucher_no=rv_num,
                        voucher_type='Receipt',
                        date=r_date,
                        reference=r.receipt_ref or None,
                        narration=remarks if (remarks and p_total == 0) else f'Receipt from {r.receipt_desc or source_acc.name}',
                        total_amount=round(r_total, 2)
                    )
                    db.session.add(je_r)
                    db.session.flush()
                    if r_cash > 0 and cash_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_r.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars='Cash Receipt'))
                    if r_prem > 0 and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_r.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars='Premier Bank Receipt'))
                    db.session.add(TransactionLine(journal_entry_id=je_r.id, account_id=source_acc.id, debit=0.0, credit=round(r_total, 2), particulars=r.receipt_desc or 'Receipt Credit'))
                    synced_count += 1

            # 3. Regular Payment
            if p_total > 0:
                p_key = p_code if p_code and p_code not in ['-', '', '#N/A'] else 'IS-016'
                target_acc = acc_dict.get(p_key) or acc_dict.get('IS-016')
                p_vno = r.payment_vno
                existing_pv = None
                if p_vno and p_vno not in ['-', '', '#N/A']:
                    existing_pv = JournalEntry.query.filter_by(voucher_no=p_vno).first()
                if not existing_pv and target_acc:
                    existing_pv = JournalEntry.query.join(TransactionLine).filter(
                        JournalEntry.date == p_date,
                        TransactionLine.account_id == target_acc.id,
                        TransactionLine.debit == round(p_total, 2)
                    ).first()

                if not existing_pv and target_acc:
                    pv_num = p_vno if p_vno and p_vno not in ['-', '', '#N/A'] else None
                    if not pv_num or JournalEntry.query.filter_by(voucher_no=pv_num).first():
                        pv_num = generate_voucher_no('Payment')
                    je_p = JournalEntry(
                        voucher_no=pv_num,
                        voucher_type='Payment',
                        date=p_date,
                        reference=r.payment_ref or None,
                        narration=remarks or f'Payment for {r.payment_desc or target_acc.name}',
                        total_amount=round(p_total, 2)
                    )
                    db.session.add(je_p)
                    db.session.flush()
                    db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=r.payment_desc or 'Expense'))
                    if p_cash > 0 and cash_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars='Paid in Cash'))
                    if p_prem > 0 and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_p.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars='Paid via Premier Bank'))
                    synced_count += 1

        db.session.commit()
        recalculate_account_balances()
        recalculate_cash_book_running_balances()
        flash(f'Synchronization successful! Synchronized {synced_count} voucher(s) into General Ledger. Cash & Bank Book and Dashboard are 100% matched.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error during synchronization: {str(e)}', 'danger')

    return redirect(url_for('chart_of_accounts'))

@app.route('/accounts/<int:account_id>/ledger')
def account_ledger(account_id):
    account = Account.query.get_or_404(account_id)

    default_from = "2026-01-01"
    default_to = "2026-12-31"
    from_str = request.args.get('from_date', default_from)
    to_str = request.args.get('to_date', default_to)

    try:
        from_date = datetime.strptime(from_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        from_date = date(2026, 1, 1)

    try:
        to_date = datetime.strptime(to_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        to_date = date(2026, 12, 31)

    # Prior movement strictly before from_date
    prior_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == account.id, JournalEntry.date < from_date).scalar()
    prior_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == account.id, JournalEntry.date < from_date).scalar()

    if account.type in ['Asset', 'Expense', 'Contra']:
        opening_balance = round(account.opening_balance + prior_dr - prior_cr, 2)
    else:
        opening_balance = round(account.opening_balance + prior_cr - prior_dr, 2)

    # All transactions between from_date and to_date
    lines = db.session.query(TransactionLine, JournalEntry)\
        .join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id)\
        .filter(
            TransactionLine.account_id == account.id,
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        )\
        .order_by(JournalEntry.date.asc(), JournalEntry.id.asc())\
        .all()

    ledger_entries = []
    current_running_bal = opening_balance
    total_debit = 0.0
    total_credit = 0.0

    for tl, je in lines:
        dr = tl.debit or 0.0
        cr = tl.credit or 0.0
        total_debit += dr
        total_credit += cr

        if account.type in ['Asset', 'Expense', 'Contra']:
            current_running_bal += (dr - cr)
        else:
            current_running_bal += (cr - dr)

        ledger_entries.append({
            'date': je.date,
            'voucher_id': je.id,
            'voucher_no': je.voucher_no,
            'voucher_type': je.voucher_type,
            'reference': je.reference,
            'narration': je.narration,
            'particulars': tl.particulars or je.narration,
            'debit': dr,
            'credit': cr,
            'running_balance': round(current_running_bal, 2)
        })

    closing_balance = round(current_running_bal, 2)
    net_movement = round(total_debit - total_credit if account.type in ['Asset', 'Expense', 'Contra'] else total_credit - total_debit, 2)
    all_accounts = Account.query.order_by(Account.code).all()

    return render_template(
        'accounts/account_ledger.html',
        account=account,
        from_date=from_date.strftime('%Y-%m-%d'),
        to_date=to_date.strftime('%Y-%m-%d'),
        opening_balance=opening_balance,
        ledger_entries=ledger_entries,
        total_debit=round(total_debit, 2),
        total_credit=round(total_credit, 2),
        net_movement=net_movement,
        closing_balance=closing_balance,
        all_accounts=all_accounts
    )

@app.route('/accounts/<int:account_id>/ledger/export-excel')
def export_account_ledger_excel(account_id):
    account = Account.query.get_or_404(account_id)
    from_str = request.args.get('from_date', '2026-01-01')
    to_str = request.args.get('to_date', '2026-12-31')

    try:
        from_date = datetime.strptime(from_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        from_date = date(2026, 1, 1)

    try:
        to_date = datetime.strptime(to_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        to_date = date(2026, 12, 31)

    prior_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == account.id, JournalEntry.date < from_date).scalar()
    prior_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == account.id, JournalEntry.date < from_date).scalar()

    if account.type in ['Asset', 'Expense', 'Contra']:
        opening_balance = round(account.opening_balance + prior_dr - prior_cr, 2)
    else:
        opening_balance = round(account.opening_balance + prior_cr - prior_dr, 2)

    lines = db.session.query(TransactionLine, JournalEntry)\
        .join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id)\
        .filter(
            TransactionLine.account_id == account.id,
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        )\
        .order_by(JournalEntry.date.asc(), JournalEntry.id.asc())\
        .all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Ledger {account.code[:15]}"

    title_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    meta_font = Font(name="Calibri", size=10, bold=False, color="475569")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    open_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws['A1'] = f"GENERAL LEDGER STATEMENT - [{account.code}] {account.name}"
    ws['A1'].font = title_font
    ws['A2'] = f"Classification: {account.type} | Category: {account.category} | Period: {from_date.strftime('%d-%b-%Y')} to {to_date.strftime('%d-%b-%Y')}"
    ws['A2'].font = meta_font

    headers = ["Date", "Voucher No.", "Type", "Ref", "Particulars & Narration", "Debit (৳)", "Credit (৳)", "Balance (৳)"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center" if col_idx in [1, 2, 3] else "left" if col_idx in [4, 5] else "right")

    # Opening balance row
    ws.cell(row=5, column=1, value=from_date.strftime('%d-%b-%Y')).alignment = Alignment(horizontal="center")
    ws.cell(row=5, column=2, value="")
    ws.cell(row=5, column=3, value="")
    ws.cell(row=5, column=4, value="")
    ws.cell(row=5, column=5, value="OPENING BALANCE B/F").font = Font(name="Calibri", bold=True)
    ws.cell(row=5, column=6, value="")
    ws.cell(row=5, column=7, value="")
    bal_cell = ws.cell(row=5, column=8, value=opening_balance)
    bal_cell.number_format = '#,##0.00'
    bal_cell.font = Font(name="Calibri", bold=True)
    for c in range(1, 9):
        ws.cell(row=5, column=c).fill = open_fill
        ws.cell(row=5, column=c).border = thin_border

    row_idx = 6
    running_bal = opening_balance
    tot_dr = 0.0
    tot_cr = 0.0

    for tl, je in lines:
        dr = tl.debit or 0.0
        cr = tl.credit or 0.0
        tot_dr += dr
        tot_cr += cr

        if account.type in ['Asset', 'Expense', 'Contra']:
            running_bal += (dr - cr)
        else:
            running_bal += (cr - dr)

        ws.cell(row=row_idx, column=1, value=je.date.strftime('%d-%b-%Y')).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=je.voucher_no).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=3, value=je.voucher_type).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=4, value=je.reference or "")
        ws.cell(row=row_idx, column=5, value=tl.particulars or je.narration)

        dr_c = ws.cell(row=row_idx, column=6, value=dr if dr > 0 else "")
        if dr > 0: dr_c.number_format = '#,##0.00'

        cr_c = ws.cell(row=row_idx, column=7, value=cr if cr > 0 else "")
        if cr > 0: cr_c.number_format = '#,##0.00'

        rb_c = ws.cell(row=row_idx, column=8, value=running_bal)
        rb_c.number_format = '#,##0.00'
        rb_c.font = Font(name="Calibri", bold=True)

        for c in range(1, 9):
            ws.cell(row=row_idx, column=c).border = thin_border

        row_idx += 1

    # Totals Row
    ws.cell(row=row_idx, column=5, value="TOTAL PERIOD MOVEMENTS").font = Font(name="Calibri", bold=True)
    t_dr = ws.cell(row=row_idx, column=6, value=tot_dr)
    t_dr.number_format = '#,##0.00'
    t_dr.font = Font(name="Calibri", bold=True, color="2563EB")
    t_cr = ws.cell(row=row_idx, column=7, value=tot_cr)
    t_cr.number_format = '#,##0.00'
    t_cr.font = Font(name="Calibri", bold=True, color="059669")
    t_bal = ws.cell(row=row_idx, column=8, value=running_bal)
    t_bal.number_format = '#,##0.00'
    t_bal.font = Font(name="Calibri", bold=True)
    for c in range(1, 9):
        ws.cell(row=row_idx, column=c).border = thin_border

    # Adjust column widths
    col_widths = {1: 14, 2: 15, 3: 12, 4: 15, 5: 45, 6: 16, 7: 16, 8: 18}
    for col_i, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = width

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"Ledger_{account.code}_{from_date.strftime('%Y%m%d')}_{to_date.strftime('%Y%m%d')}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==============================================================================
# ROUTES: PAYROLL MANAGEMENT & SALARY DISBURSEMENT
# ==============================================================================

@app.route('/payroll')
def payroll_dashboard():
    employees = Employee.query.order_by(Employee.employee_id).all()
    payroll_records = PayrollRecord.query.order_by(PayrollRecord.processed_date.desc(), PayrollRecord.id.desc()).all()

    total_basic = sum(e.basic_salary for e in employees if e.status == 'Active')
    total_allowances = sum(e.allowances for e in employees if e.status == 'Active')
    total_deductions = sum(e.deductions for e in employees if e.status == 'Active')
    total_gross_payroll = total_basic + total_allowances
    total_net_payroll = total_gross_payroll - total_deductions

    return render_template(
        'payroll/payroll_dashboard.html',
        employees=employees,
        payroll_records=payroll_records,
        total_basic=total_basic,
        total_allowances=total_allowances,
        total_deductions=total_deductions,
        total_gross_payroll=total_gross_payroll,
        total_net_payroll=total_net_payroll
    )

@app.route('/payroll/employees/new', methods=['GET', 'POST'])
def add_employee():
    if request.method == 'POST':
        emp_id = request.form.get('employee_id')
        name = request.form.get('name')
        designation = request.form.get('designation')
        department = request.form.get('department')
        basic = float(request.form.get('basic_salary') or 0.0)
        allowances = float(request.form.get('allowances') or 0.0)
        deductions = float(request.form.get('deductions') or 0.0)
        bank_acc = request.form.get('bank_account_no', '').strip() or None
        nid = request.form.get('nid_passport', '').strip() or None
        phone = request.form.get('phone', '').strip() or None
        email = request.form.get('email', '').strip() or None
        present_address = request.form.get('present_address', '').strip() or None
        j_date_str = request.form.get('joining_date')
        joining_date = parse_accounting_date(j_date_str) if j_date_str else None

        new_emp = Employee(
            employee_id=emp_id,
            name=name,
            designation=designation,
            department=department,
            basic_salary=basic,
            allowances=allowances,
            deductions=deductions,
            bank_account_no=bank_acc,
            nid_passport=nid,
            phone=phone,
            email=email,
            present_address=present_address,
            joining_date=joining_date,
            status='Active'
        )
        db.session.add(new_emp)
        db.session.commit()

        flash(f'Employee "{name}" ({emp_id}) added to payroll system.', 'success')
        return redirect(url_for('payroll_dashboard'))

    # Suggested employee ID
    emp_count = Employee.query.count()
    suggested_emp_id = f"EMP-{emp_count + 1:03d}"
    return render_template('payroll/add_employee.html', suggested_emp_id=suggested_emp_id)

@app.route('/payroll/employees/<int:employee_id>/edit', methods=['GET', 'POST'])
def edit_employee(employee_id):
    employee = Employee.query.get_or_404(employee_id)

    if request.method == 'POST':
        employee.employee_id = request.form.get('employee_id')
        employee.name = request.form.get('name')
        employee.designation = request.form.get('designation')
        employee.department = request.form.get('department')
        employee.basic_salary = float(request.form.get('basic_salary') or 0.0)
        employee.allowances = float(request.form.get('allowances') or 0.0)
        employee.deductions = float(request.form.get('deductions') or 0.0)
        employee.bank_account_no = request.form.get('bank_account_no', '').strip() or None
        employee.nid_passport = request.form.get('nid_passport', '').strip() or None
        employee.phone = request.form.get('phone', '').strip() or None
        employee.email = request.form.get('email', '').strip() or None
        employee.present_address = request.form.get('present_address', '').strip() or None
        j_date_str = request.form.get('joining_date')
        employee.joining_date = parse_accounting_date(j_date_str) if j_date_str else None

        db.session.commit()
        flash(f'Employee profile for "{employee.name}" updated successfully.', 'success')
        return redirect(url_for('employee_profile', emp_id=employee.id))

    return render_template('payroll/edit_employee.html', employee=employee)

@app.route('/payroll/employees/<int:employee_id>/delete', methods=['POST'])
def delete_employee(employee_id):
    employee = Employee.query.get_or_404(employee_id)
    emp_name = employee.name
    db.session.delete(employee)
    db.session.commit()

    flash(f'Employee "{emp_name}" deleted from payroll.', 'info')
    return redirect(url_for('payroll_dashboard'))

@app.route('/payroll/run', methods=['GET', 'POST'])
def run_payroll():
    employees = Employee.query.filter_by(status='Active').order_by(Employee.employee_id).all()
    liquid_accounts = Account.query.filter_by(is_cash_or_bank=True).all()
    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.code).all()

    total_basic = sum(e.basic_salary for e in employees)
    total_allowances = sum(e.allowances for e in employees)
    total_deductions = sum(e.deductions for e in employees)
    total_net = (total_basic + total_allowances) - total_deductions

    current_month = date.today().strftime('%Y-%m')
    today_str = date.today().strftime('%Y-%m-%d')

    if request.method == 'POST':
        payroll_month = request.form.get('payroll_month', current_month)
        processed_date_str = request.form.get('processed_date', today_str)
        post_to_gl = request.form.get('post_to_gl') == 'yes'
        payment_acc_val = request.form.get('payment_account_id')
        expense_acc_val = request.form.get('expense_account_id')

        try:
            processed_date = datetime.strptime(processed_date_str, '%Y-%m-%d').date()
        except (ValueError, TypeError):
            processed_date = date.today()

        # Check if already processed for this month
        existing_run = PayrollRecord.query.filter_by(payroll_month=payroll_month).first()
        if existing_run:
            flash(f'Payroll for month {payroll_month} has already been processed (Run #{existing_run.id}).', 'warning')
            return redirect(url_for('view_payroll_run', record_id=existing_run.id))

        payment_acc_id = int(payment_acc_val) if payment_acc_val and str(payment_acc_val).isdigit() else (liquid_accounts[0].id if liquid_accounts else 1)
        expense_acc_id = int(expense_acc_val) if expense_acc_val and str(expense_acc_val).isdigit() else None

        salary_pv = None
        pv_no = ""
        # 1. Only post Payment Voucher (PV) to GL if explicitly requested (keeping HR & Payroll separate by default)
        if post_to_gl and payment_acc_id and expense_acc_id:
            pv_no = generate_voucher_no('Payment')
            narration = f"Monthly staff salary disbursement for {payroll_month} ({len(employees)} active personnel)"
            
            salary_pv = JournalEntry(
                voucher_no=pv_no,
                voucher_type="Payment",
                date=processed_date,
                reference=f"PAYROLL-{payroll_month}",
                narration=narration,
                total_amount=round(total_net, 2)
            )
            db.session.add(salary_pv)
            db.session.flush()

            # Dr: Salaries Expense Account
            db.session.add(TransactionLine(
                journal_entry_id=salary_pv.id,
                account_id=expense_acc_id,
                debit=round(total_net, 2),
                credit=0.00,
                particulars=f"Salary Expense {payroll_month}"
            ))

            # Cr: Selected Bank / Cash Account
            db.session.add(TransactionLine(
                journal_entry_id=salary_pv.id,
                account_id=payment_acc_id,
                debit=0.00,
                credit=round(total_net, 2),
                particulars=f"Bank Disbursal {payroll_month}"
            ))

        # 2. Create Payroll Record
        payroll_rec = PayrollRecord(
            payroll_month=payroll_month,
            processed_date=processed_date,
            total_basic=total_basic,
            total_allowances=total_allowances,
            total_deductions=total_deductions,
            total_net=total_net,
            payment_account_id=payment_acc_id,
            voucher_id=salary_pv.id if salary_pv else None
        )
        db.session.add(payroll_rec)
        db.session.flush()

        # 3. Create Individual Employee Payroll Items
        for emp in employees:
            emp_net = (emp.basic_salary + emp.allowances) - emp.deductions
            item = PayrollItem(
                payroll_record_id=payroll_rec.id,
                employee_id=emp.id,
                basic_salary=emp.basic_salary,
                allowances=emp.allowances,
                deductions=emp.deductions,
                net_salary=round(emp_net, 2)
            )
            db.session.add(item)

        db.session.commit()
        if salary_pv:
            recalculate_account_balances()
            flash(f'Payroll for {payroll_month} disbursed successfully! Payment Voucher {pv_no} posted to General Ledger.', 'success')
        else:
            flash(f'Payroll for {payroll_month} saved successfully! HR & Payroll processed independently without affecting Accounts.', 'success')
        return redirect(url_for('view_payroll_run', record_id=payroll_rec.id))

    return render_template(
        'payroll/run_payroll.html',
        employees=employees,
        liquid_accounts=liquid_accounts,
        expense_accounts=expense_accounts,
        total_basic=total_basic,
        total_allowances=total_allowances,
        total_deductions=total_deductions,
        total_net=total_net,
        current_month=current_month,
        today=today_str
    )

@app.route('/payroll/run/<int:record_id>')
def view_payroll_run(record_id):
    record = PayrollRecord.query.get_or_404(record_id)
    return render_template('payroll/view_payroll_run.html', record=record)

@app.route('/payroll/payslip/<int:item_id>')
def view_payslip(item_id):
    item = PayrollItem.query.get_or_404(item_id)
    return render_template('payroll/payslip.html', item=item)


# ==============================================================================
# ROUTES: FINANCIAL REPORTS MODULE
# ==============================================================================

@app.route('/reports')
def reports_hub():
    return render_template('reports/reports_hub.html')

def clean_accounting_num(s):
    if not s or str(s).strip() in ("-", "", "#N/A", "None"):
        return 0.0
    s_clean = str(s).strip().replace(",", "")
    try:
        return float(s_clean)
    except Exception:
        return 0.0

def parse_accounting_date(date_val):
    if not date_val:
        return None
    if isinstance(date_val, datetime):
        return date_val.date()
    if isinstance(date_val, date):
        return date_val
    s = str(date_val).strip()
    if s in ("#N/A", "-", "", "None"):
        return None
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    return None

# 0. CASH & BANK BOOK (RECEIPTS & PAYMENTS JOURNAL AS PER CLIENT EXCEL)
@app.route('/reports/cash-book')
def cash_book_report():
    month = request.args.get('month', 'all')
    search = request.args.get('search', '').strip()
    from_str = request.args.get('from_date', '')
    to_str = request.args.get('to_date', '')

    query = CashBookRow.query

    month_options = [
        ('all', 'All Months (Full 2026)'),
        ('2026-01', 'January 2026'),
        ('2026-02', 'February 2026'),
        ('2026-03', 'March 2026'),
        ('2026-04', 'April 2026'),
        ('2026-05', 'May 2026'),
        ('2026-06', 'June 2026'),
        ('2026-07', 'July 2026'),
        ('2026-08', 'August 2026'),
        ('2026-09', 'September 2026'),
        ('2026-10', 'October 2026'),
        ('2026-11', 'November 2026'),
        ('2026-12', 'December 2026'),
    ]

    def is_opening_item(desc):
        if not desc:
            return False
        dl = str(desc).lower().strip()
        return "opening" in dl or "opeining" in dl

    if month and month != 'all':
        try:
            year, m = map(int, month.split('-'))
            import calendar
            last_day = calendar.monthrange(year, m)[1]
            m_start = date(year, m, 1)
            m_end = date(year, m, last_day)
            query = query.filter(
                db.or_(
                    db.and_(CashBookRow.payment_date >= m_start, CashBookRow.payment_date <= m_end),
                    db.and_(CashBookRow.receipt_date >= m_start, CashBookRow.receipt_date <= m_end),
                    CashBookRow.receipt_desc.ilike('%open%')
                )
            )
        except Exception:
            pass
    elif from_str and to_str:
        try:
            f_d = datetime.strptime(from_str, '%Y-%m-%d').date()
            t_d = datetime.strptime(to_str, '%Y-%m-%d').date()
            query = query.filter(
                db.or_(
                    db.and_(CashBookRow.payment_date >= f_d, CashBookRow.payment_date <= t_d),
                    db.and_(CashBookRow.receipt_date >= f_d, CashBookRow.receipt_date <= t_d),
                    CashBookRow.receipt_desc.ilike('%open%')
                )
            )
        except Exception:
            pass

    if search:
        search_pattern = f"%{search}%"
        query = query.filter(
            db.or_(
                CashBookRow.payment_vno.ilike(search_pattern),
                CashBookRow.payment_code.ilike(search_pattern),
                CashBookRow.payment_desc.ilike(search_pattern),
                CashBookRow.receipt_code.ilike(search_pattern),
                CashBookRow.receipt_desc.ilike(search_pattern),
                CashBookRow.remarks.ilike(search_pattern)
            )
        )

    rows = query.order_by(CashBookRow.row_order.asc(), CashBookRow.id.asc()).all()

    total_receipt_cash = sum(r.receipt_cash for r in rows if not is_opening_item(r.receipt_desc))
    total_receipt_premier = sum(r.receipt_premier for r in rows if not is_opening_item(r.receipt_desc))
    total_receipt_all = sum(r.receipt_total for r in rows if not is_opening_item(r.receipt_desc))

    total_payment_cash = sum(r.payment_cash for r in rows)
    total_payment_premier = sum(r.payment_premier for r in rows)
    total_payment_all = sum(r.payment_total for r in rows)

    # Closing balances for the selected period
    valid_bal_rows = [r for r in rows if r.balance_cash > 0 or r.balance_premier != 0]
    if valid_bal_rows:
        latest_cash = valid_bal_rows[-1].balance_cash
        latest_prem = valid_bal_rows[-1].balance_premier
    else:
        latest_row = CashBookRow.query.order_by(CashBookRow.row_order.desc(), CashBookRow.id.desc()).first()
        latest_cash = latest_row.balance_cash if latest_row else 0.0
        latest_prem = latest_row.balance_premier if latest_row else 0.0

    return render_template(
        'reports/cash_book.html',
        rows=rows,
        selected_month=month,
        search_query=search,
        from_date=from_str,
        to_date=to_str,
        month_options=month_options,
        total_receipt_cash=total_receipt_cash,
        total_receipt_premier=total_receipt_premier,
        total_receipt_all=total_receipt_all,
        total_payment_cash=total_payment_cash,
        total_payment_premier=total_payment_premier,
        total_payment_all=total_payment_all,
        latest_cash=latest_cash,
        latest_prem=latest_prem
    )

@app.route('/reports/cash-book/export-excel')
def export_cash_book_excel():
    month = request.args.get('month', 'all')
    query = CashBookRow.query
    if month and month != 'all':
        try:
            year, m = map(int, month.split('-'))
            import calendar
            last_day = calendar.monthrange(year, m)[1]
            m_start = date(year, m, 1)
            m_end = date(year, m, last_day)
            query = query.filter(
                db.or_(
                    db.and_(CashBookRow.payment_date >= m_start, CashBookRow.payment_date <= m_end),
                    db.and_(CashBookRow.receipt_date >= m_start, CashBookRow.receipt_date <= m_end),
                    CashBookRow.receipt_desc.ilike('%open%')
                )
            )
        except Exception:
            pass

    rows = query.order_by(CashBookRow.row_order.asc(), CashBookRow.id.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cash & Bank Book"

    green_header_fill = PatternFill(start_color="C2E0C6", end_color="C2E0C6", fill_type="solid")
    blue_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    peach_header_fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
    separator_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    dv_green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    prem_blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    light_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    hdr_font = Font(name="Calibri", size=11, bold=True)
    bold_font = Font(name="Calibri", size=10, bold=True)
    dv_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Row 1: Merged Top Headers as per client picture
    ws.merge_cells('A1:J1')
    ws['A1'] = "Receipts"
    ws['A1'].fill = green_header_fill
    ws['A1'].font = hdr_font
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws['K1'].fill = separator_fill

    ws.merge_cells('L1:T1')
    ws['L1'] = "Payments"
    ws['L1'].fill = blue_header_fill
    ws['L1'].font = hdr_font
    ws['L1'].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells('U1:W1')
    ws['U1'] = "Balance"
    ws['U1'].fill = peach_header_fill
    ws['U1'].font = hdr_font
    ws['U1'].alignment = Alignment(horizontal="center", vertical="center")

    ws['X1'] = "Remarks"
    ws['X1'].font = hdr_font
    ws['X1'].alignment = Alignment(horizontal="center", vertical="center")

    # Row 2: Sub-headers
    subheaders = [
        "Code", "Date", "", "Voucher", "Description", "Ref", "Cash", "Premier Bank", "Bank", "Total",
        "", # K
        "Date", "Voucher No.", "Code", "Description", "Ref", "Cash", "Bank", "Premier Bank", "Total",
        "Cash", "Bank", "PremierBank", "Remarks"
    ]
    for c_idx, sh in enumerate(subheaders, 1):
        cell = ws.cell(row=2, column=c_idx, value=sh)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        if c_idx <= 10:
            cell.fill = light_green_fill
        elif c_idx == 11:
            cell.fill = separator_fill
        elif c_idx <= 20:
            cell.fill = blue_header_fill
        elif c_idx <= 23:
            cell.fill = peach_header_fill

    # Highlight Premier Bank column headers in soft blue
    ws.cell(row=2, column=8).fill = prem_blue_fill
    ws.cell(row=2, column=19).fill = prem_blue_fill
    ws.cell(row=2, column=23).fill = prem_blue_fill

    row_num = 3
    for r in rows:
        ws.cell(row=row_num, column=1, value=r.receipt_code or "").alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=2, value=r.receipt_date.strftime('%d-%b-%y') if r.receipt_date else "").alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=3, value="")
        ws.cell(row=row_num, column=4, value=r.receipt_voucher or "").alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=5, value=r.receipt_desc or "")
        ws.cell(row=row_num, column=6, value=r.receipt_ref or "")

        c_rcash = ws.cell(row=row_num, column=7, value=r.receipt_cash if r.receipt_cash != 0 else "-")
        if r.receipt_cash != 0: c_rcash.number_format = '#,##0'

        c_rprem = ws.cell(row=row_num, column=8, value=r.receipt_premier if r.receipt_premier != 0 else "-")
        c_rprem.fill = prem_blue_fill
        if r.receipt_premier != 0: c_rprem.number_format = '#,##0'

        ws.cell(row=row_num, column=9, value="-").alignment = Alignment(horizontal="center")

        c_rtot = ws.cell(row=row_num, column=10, value=r.receipt_total if r.receipt_total != 0 else "-")
        c_rtot.font = bold_font
        if r.receipt_total != 0: c_rtot.number_format = '#,##0'

        ws.cell(row=row_num, column=11).fill = separator_fill

        ws.cell(row=row_num, column=12, value=r.payment_date.strftime('%d-%b-%y') if r.payment_date else "").alignment = Alignment(horizontal="center")
        
        c_vno = ws.cell(row=row_num, column=13, value=r.payment_vno or "")
        c_vno.alignment = Alignment(horizontal="center")
        if r.payment_vno and r.payment_vno.startswith("DV-"):
            c_vno.fill = dv_green_fill
            c_vno.font = dv_font

        ws.cell(row=row_num, column=14, value=r.payment_code or "").alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=15, value=r.payment_desc or "")
        ws.cell(row=row_num, column=16, value=r.payment_ref or "")

        c_pcash = ws.cell(row=row_num, column=17, value=r.payment_cash if r.payment_cash != 0 else "-")
        if r.payment_cash != 0: c_pcash.number_format = '#,##0'

        ws.cell(row=row_num, column=18, value="-").alignment = Alignment(horizontal="center")

        c_pprem = ws.cell(row=row_num, column=19, value=r.payment_premier if r.payment_premier != 0 else "-")
        c_pprem.fill = prem_blue_fill
        if r.payment_premier != 0: c_pprem.number_format = '#,##0'

        c_ptot = ws.cell(row=row_num, column=20, value=r.payment_total if r.payment_total != 0 else "-")
        c_ptot.font = bold_font
        if r.payment_total != 0: c_ptot.number_format = '#,##0'

        c_bcash = ws.cell(row=row_num, column=21, value=r.balance_cash if r.balance_cash != 0 else "-")
        c_bcash.font = bold_font
        if r.balance_cash != 0: c_bcash.number_format = '#,##0'

        ws.cell(row=row_num, column=22, value="-").alignment = Alignment(horizontal="center")

        c_bprem = ws.cell(row=row_num, column=23, value=r.balance_premier if r.balance_premier != 0 else "-")
        c_bprem.fill = prem_blue_fill
        c_bprem.font = bold_font
        if r.balance_premier != 0: c_bprem.number_format = '#,##0'

        ws.cell(row=row_num, column=24, value=r.remarks or "")

        for col_k in range(1, 25):
            if col_k != 11:
                ws.cell(row=row_num, column=col_k).border = thin_border

        row_num += 1

    widths = {
        1: 10, 2: 12, 3: 4, 4: 12, 5: 25, 6: 8, 7: 12, 8: 14, 9: 8, 10: 14,
        11: 3,
        12: 12, 13: 13, 14: 10, 15: 25, 16: 8, 17: 12, 18: 8, 19: 14, 20: 14,
        21: 14, 22: 8, 23: 14, 24: 50
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"Cash_and_Bank_Book_{month}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/reports/cash-book/import-excel', methods=['POST'])
def import_cash_book_excel():
    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Please select an Excel (.xlsx/.xls) or Tab-Separated (.tsv) file.', 'danger')
        return redirect(url_for('cash_book_report'))

    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.tsv') or filename.endswith('.csv')):
        flash('Unsupported format. Please upload an Excel (.xlsx / .xls) file.', 'warning')
        return redirect(url_for('cash_book_report'))

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active

        header_row = 2
        for r in range(1, min(10, ws.max_row + 1)):
            row_vals = [str(ws.cell(row=r, column=c).value or "").lower() for c in range(1, min(25, ws.max_column + 1))]
            if any("voucher" in v or "payment" in v or "receipt" in v for v in row_vals):
                header_row = r
                break

        imported_count = 0
        acc_dict = {a.code: a for a in Account.query.all()}
        cash_acc = acc_dict.get("BS-001")
        prem_acc = acc_dict.get("BS-002")

        max_order = db.session.query(func.coalesce(func.max(CashBookRow.row_order), 0)).scalar()

        for r_idx in range(header_row + 1, ws.max_row + 1):
            row_vals = [ws.cell(row=r_idx, column=c).value for c in range(1, 26)]
            if not any(v for v in row_vals if str(v).strip() and str(v).strip() != "-"):
                continue

            r_code = str(row_vals[0] or "").strip()
            r_date_raw = row_vals[1] or row_vals[2]
            r_voucher = str(row_vals[3] or "").strip()
            r_desc = str(row_vals[4] or "").strip()
            r_ref = str(row_vals[5] or "").strip()
            r_cash = clean_accounting_num(row_vals[6])
            r_prem = clean_accounting_num(row_vals[7])
            r_bank = clean_accounting_num(row_vals[8])
            r_total = clean_accounting_num(row_vals[9]) or (r_cash + r_prem)

            p_date_raw = row_vals[11]
            p_vno = str(row_vals[12] or "").strip()
            p_code = str(row_vals[13] or "").strip()
            p_desc = str(row_vals[14] or "").strip()
            p_ref = str(row_vals[15] or "").strip()
            p_cash = clean_accounting_num(row_vals[16])
            p_bank = clean_accounting_num(row_vals[17])
            p_prem = clean_accounting_num(row_vals[18])
            p_total = clean_accounting_num(row_vals[19]) or (p_cash + p_prem + p_bank)

            bal_cash = clean_accounting_num(row_vals[20])
            bal_bank = clean_accounting_num(row_vals[21])
            bal_prem = clean_accounting_num(row_vals[22])
            remarks = str(row_vals[23] or "").strip()

            r_date = parse_accounting_date(r_date_raw)
            p_date = parse_accounting_date(p_date_raw)

            # Check if CashBookRow already exists to prevent duplicate rows
            is_open_row = "open" in r_desc.lower() or "opein" in r_desc.lower()
            existing_cb = None
            if is_open_row:
                existing_cb = CashBookRow.query.filter(CashBookRow.receipt_desc.ilike('%open%')).first()
            elif p_vno and p_vno not in ["-", "", "#N/A"]:
                existing_cb = CashBookRow.query.filter_by(payment_vno=p_vno).first()
            elif r_voucher and r_voucher not in ["-", "", "#N/A"]:
                existing_cb = CashBookRow.query.filter_by(receipt_voucher=r_voucher).first()
            elif p_total > 0 and p_date:
                existing_cb = CashBookRow.query.filter(
                    CashBookRow.payment_date == p_date,
                    CashBookRow.payment_total == p_total,
                    CashBookRow.payment_desc == p_desc
                ).first()

            if not existing_cb:
                max_order += 1
                cb_row = CashBookRow(
                    row_order=max_order,
                    receipt_code=r_code if r_code != "#N/A" else "",
                    receipt_date=r_date,
                    receipt_voucher=r_voucher,
                    receipt_desc=r_desc if r_desc != "#N/A" else "",
                    receipt_ref=r_ref,
                    receipt_cash=r_cash,
                    receipt_premier=r_prem,
                    receipt_bank=r_bank,
                    receipt_total=r_total,
                    payment_date=p_date,
                    payment_vno=p_vno,
                    payment_code=p_code if p_code != "#N/A" else "",
                    payment_desc=p_desc,
                    payment_ref=p_ref,
                    payment_cash=p_cash,
                    payment_bank=p_bank,
                    payment_premier=p_prem,
                    payment_total=p_total,
                    balance_cash=bal_cash,
                    balance_bank=bal_bank,
                    balance_premier=bal_prem,
                    remarks=remarks
                )
                db.session.add(cb_row)

            # Case 1: Contra Transaction (BS-003 Cash withdraw / deposit)
            if r_code == "BS-003" or p_code == "BS-003":
                c_date = r_date or p_date
                if isinstance(c_date, datetime):
                    c_date = c_date.date()
                v_num = p_vno or r_voucher
                existing_cv = None
                if v_num and v_num not in ["-", "", "#N/A"]:
                    existing_cv = JournalEntry.query.filter_by(voucher_no=v_num).first()
                if not existing_cv:
                    existing_cv = JournalEntry.query.filter(
                        JournalEntry.voucher_type == "Contra",
                        JournalEntry.date == c_date,
                        JournalEntry.total_amount == round(p_total or r_total, 2)
                    ).first()

                if not existing_cv:
                    cv_num = v_num if v_num and v_num not in ["-", "", "#N/A"] else None
                    if not cv_num or JournalEntry.query.filter_by(voucher_no=cv_num).first():
                        cv_num = generate_voucher_no("Contra")
                    je = JournalEntry(
                        voucher_no=cv_num,
                        voucher_type="Contra",
                        date=c_date,
                        reference="Cash Contra",
                        narration=remarks or "Cash Contra transfer",
                        total_amount=round(p_total or r_total, 2)
                    )
                    db.session.add(je)
                    db.session.flush()
                    if r_cash > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars="Cash received"))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=0.0, credit=round(r_cash, 2), particulars="Premier Bank withdrawal"))
                    elif r_prem > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars="Premier Bank received"))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=0.0, credit=round(r_prem, 2), particulars="Cash disbursed"))

                # If there was ALSO a payment on that line
                if p_code and p_code != "BS-003" and p_total > 0:
                    p_key = p_code if p_code not in ["-", "", "#N/A"] else "IS-016"
                    if p_key not in acc_dict:
                        new_acc = Account(code=p_key, name=p_desc or p_key, type="Expense", category="Operating Expense")
                        db.session.add(new_acc)
                        db.session.flush()
                        acc_dict[p_key] = new_acc
                    target_acc = acc_dict[p_key]
                    v_num = generate_voucher_no('Payment')
                    je_pmt = JournalEntry(
                        voucher_no=v_num,
                        voucher_type="Payment",
                        date=c_date,
                        reference=p_ref or None,
                        narration=remarks or f"Payment for {p_desc or target_acc.name}",
                        total_amount=round(p_total, 2)
                    )
                    db.session.add(je_pmt)
                    db.session.flush()
                    db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=p_desc or "Expense"))
                    if p_cash > 0 and cash_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars="Cash Payment"))
                    if p_prem > 0 and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars="Premier Bank Payment"))
            else:
                # 1. Post Receipt Voucher to General Ledger
                if r_total > 0 and not is_open_row:
                    r_key = r_code if r_code and r_code not in ["-", "", "#N/A"] else "BS-035"
                    if r_key not in acc_dict:
                        new_acc = Account(code=r_key, name=r_desc or "Accounts-TBN-25", type="Revenue", category="Operating Revenue")
                        db.session.add(new_acc)
                        db.session.flush()
                        acc_dict[r_key] = new_acc

                    source_acc = acc_dict[r_key]
                    v_date = r_date or date.today()
                    if isinstance(v_date, datetime):
                        v_date = v_date.date()

                    # Deduplication check
                    existing_je = None
                    if r_voucher and r_voucher not in ["-", "", "#N/A"]:
                        existing_je = JournalEntry.query.filter_by(voucher_no=r_voucher).first()
                    if not existing_je:
                        existing_je = JournalEntry.query.join(TransactionLine).filter(
                            JournalEntry.date == v_date,
                            TransactionLine.account_id == source_acc.id,
                            TransactionLine.credit == round(r_total, 2)
                        ).first()

                    if not existing_je:
                        v_num = r_voucher if r_voucher and r_voucher not in ["-", "", "#N/A"] else None
                        if not v_num or JournalEntry.query.filter_by(voucher_no=v_num).first():
                            v_num = generate_voucher_no('Receipt')
                        je = JournalEntry(
                            voucher_no=v_num,
                            voucher_type="Receipt",
                            date=v_date,
                            reference=r_ref or None,
                            narration=remarks if (remarks and p_total == 0) else f"Receipt from {r_desc or source_acc.name}",
                            total_amount=round(r_total, 2)
                        )
                        db.session.add(je)
                        db.session.flush()
                        if r_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars="Cash Receipt"))
                        if r_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars="Premier Bank Receipt"))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=source_acc.id, debit=0.0, credit=round(r_total, 2), particulars=r_desc or "Receipt Credit"))

                # 2. Post Payment Voucher to General Ledger
                if p_total > 0:
                    p_key = p_code if p_code and p_code not in ["-", "", "#N/A"] else "IS-016"
                    if p_key not in acc_dict:
                        new_acc = Account(code=p_key, name=p_desc or p_key, type="Expense", category="Operating Expense")
                        db.session.add(new_acc)
                        db.session.flush()
                        acc_dict[p_key] = new_acc

                    target_acc = acc_dict[p_key]
                    v_date = p_date or date.today()
                    if isinstance(v_date, datetime):
                        v_date = v_date.date()

                    # Deduplication check
                    existing_je = None
                    if p_vno and p_vno not in ["-", "", "#N/A"]:
                        existing_je = JournalEntry.query.filter_by(voucher_no=p_vno).first()
                    if not existing_je:
                        existing_je = JournalEntry.query.join(TransactionLine).filter(
                            JournalEntry.date == v_date,
                            TransactionLine.account_id == target_acc.id,
                            TransactionLine.debit == round(p_total, 2)
                        ).first()

                    if not existing_je:
                        v_num = p_vno if p_vno and p_vno not in ["-", "", "#N/A"] else None
                        if not v_num or JournalEntry.query.filter_by(voucher_no=v_num).first():
                            v_num = generate_voucher_no('Payment')
                        je = JournalEntry(
                            voucher_no=v_num,
                            voucher_type="Payment",
                            date=v_date,
                            reference=p_ref or None,
                            narration=remarks or f"Payment for {p_desc or target_acc.name}",
                            total_amount=round(p_total, 2)
                        )
                        db.session.add(je)
                        db.session.flush()
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=p_desc or "Expense"))
                        if p_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars="Cash Payment"))
                        if p_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars="Premier Bank Payment"))

            imported_count += 1

        db.session.commit()
        recalculate_account_balances()
        recalculate_cash_book_running_balances()
        flash(f'Excel file imported successfully! Processed {imported_count} row(s) and synchronized the General Ledger without duplicates.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Error importing Excel file: {str(e)}', 'danger')

    return redirect(url_for('cash_book_report'))


@app.route('/reports/cash-book/update-date-range', methods=['POST'])
def update_cash_book_by_date_range():
    """
    Clears all previous data of accounts ledger (vouchers & lines) and cash book rows
    within the specified date condition (From Date to To Date), and updates/imports new
    data only within that date condition from the uploaded Excel file or recalculation.
    """
    from_str = request.form.get('from_date', '').strip()
    to_str = request.form.get('to_date', '').strip()
    action_type = request.form.get('action_type', 'replace_excel') # 'replace_excel' or 'clear_range'
    file = request.files.get('excel_file')

    try:
        from_date = parse_accounting_date(from_str) or datetime.strptime(from_str, '%Y-%m-%d').date()
        to_date = parse_accounting_date(to_str) or datetime.strptime(to_str, '%Y-%m-%d').date()
        if isinstance(from_date, datetime):
            from_date = from_date.date()
        if isinstance(to_date, datetime):
            to_date = to_date.date()
    except (ValueError, TypeError):
        flash('Invalid date range specified. Please provide valid From Date and To Date.', 'danger')
        return redirect(url_for('cash_book_report'))

    if from_date > to_date:
        flash('From Date cannot be later than To Date.', 'warning')
        return redirect(url_for('cash_book_report'))

    try:
        # Step 1: Clear previous ledger data & cash book data within the specified date condition
        target_jes = JournalEntry.query.filter(
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        ).all()
        je_ids = [j.id for j in target_jes]
        cleared_lines = 0
        cleared_jes = len(je_ids)

        if je_ids:
            cleared_lines = TransactionLine.query.filter(TransactionLine.journal_entry_id.in_(je_ids)).delete(synchronize_session=False)
            JournalEntry.query.filter(JournalEntry.id.in_(je_ids)).delete(synchronize_session=False)

        # Clear CashBookRows strictly within [from_date, to_date], protecting row 1 (Opening Balance)
        cb_query = CashBookRow.query.filter(
            db.or_(
                db.and_(CashBookRow.payment_date >= from_date, CashBookRow.payment_date <= to_date),
                db.and_(CashBookRow.receipt_date >= from_date, CashBookRow.receipt_date <= to_date)
            )
        )
        if from_date <= date(2026, 1, 1):
            cb_query = cb_query.filter(CashBookRow.id != 1)

        cleared_cb_rows = cb_query.delete(synchronize_session=False)
        db.session.flush()

        imported_count = 0
        if action_type == 'replace_excel' and file and file.filename:
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active

            header_row = 2
            for r in range(1, min(10, ws.max_row + 1)):
                row_vals = [str(ws.cell(row=r, column=c).value or "").lower() for c in range(1, min(25, ws.max_column + 1))]
                if any("voucher" in v or "payment" in v or "receipt" in v for v in row_vals):
                    header_row = r
                    break

            acc_dict = {a.code: a for a in Account.query.all()}
            cash_acc = acc_dict.get("BS-001")
            prem_acc = acc_dict.get("BS-002")

            max_order = db.session.query(func.coalesce(func.max(CashBookRow.row_order), 0)).scalar() or 0

            for r_idx in range(header_row + 1, ws.max_row + 1):
                row_vals = [ws.cell(row=r_idx, column=c).value for c in range(1, 26)]
                if not any(v for v in row_vals if str(v).strip() and str(v).strip() != "-"):
                    continue

                r_code = str(row_vals[0] or "").strip()
                r_date_raw = row_vals[1] or row_vals[2]
                r_voucher = str(row_vals[3] or "").strip()
                r_desc = str(row_vals[4] or "").strip()
                r_ref = str(row_vals[5] or "").strip()
                r_cash = clean_accounting_num(row_vals[6])
                r_prem = clean_accounting_num(row_vals[7])
                r_bank = clean_accounting_num(row_vals[8])
                r_total = clean_accounting_num(row_vals[9]) or (r_cash + r_prem)

                p_date_raw = row_vals[11]
                p_vno = str(row_vals[12] or "").strip()
                p_code = str(row_vals[13] or "").strip()
                p_desc = str(row_vals[14] or "").strip()
                p_ref = str(row_vals[15] or "").strip()
                p_cash = clean_accounting_num(row_vals[16])
                p_bank = clean_accounting_num(row_vals[17])
                p_prem = clean_accounting_num(row_vals[18])
                p_total = clean_accounting_num(row_vals[19]) or (p_cash + p_prem + p_bank)

                bal_cash = clean_accounting_num(row_vals[20])
                bal_bank = clean_accounting_num(row_vals[21])
                bal_prem = clean_accounting_num(row_vals[22])
                remarks = str(row_vals[23] or "").strip()

                r_date = parse_accounting_date(r_date_raw)
                p_date = parse_accounting_date(p_date_raw)
                if isinstance(r_date, datetime):
                    r_date = r_date.date()
                if isinstance(p_date, datetime):
                    p_date = p_date.date()

                # STRICT DATE CONDITION: Only process rows that fall strictly within [from_date, to_date]
                row_date = p_date or r_date
                if isinstance(row_date, datetime):
                    row_date = row_date.date()
                if not row_date or row_date < from_date or row_date > to_date:
                    continue

                is_open_row = "open" in r_desc.lower() or "opein" in r_desc.lower()
                if is_open_row and from_date > date(2026, 1, 1):
                    continue

                max_order += 1
                cb_row = CashBookRow(
                    row_order=max_order,
                    receipt_code=r_code if r_code != "#N/A" else "",
                    receipt_date=r_date,
                    receipt_voucher=r_voucher,
                    receipt_desc=r_desc if r_desc != "#N/A" else "",
                    receipt_ref=r_ref,
                    receipt_cash=r_cash,
                    receipt_premier=r_prem,
                    receipt_bank=r_bank,
                    receipt_total=r_total,
                    payment_date=p_date,
                    payment_vno=p_vno,
                    payment_code=p_code if p_code != "#N/A" else "",
                    payment_desc=p_desc,
                    payment_ref=p_ref,
                    payment_cash=p_cash,
                    payment_bank=p_bank,
                    payment_premier=p_prem,
                    payment_total=p_total,
                    balance_cash=bal_cash,
                    balance_bank=bal_bank,
                    balance_premier=bal_prem,
                    remarks=remarks
                )
                db.session.add(cb_row)

                # Post Contra Transaction
                if r_code == "BS-003" or p_code == "BS-003":
                    c_date = r_date or p_date
                    if isinstance(c_date, datetime):
                        c_date = c_date.date()
                    v_num = p_vno or r_voucher
                    if not v_num or v_num in ["-", "", "#N/A"] or JournalEntry.query.filter_by(voucher_no=v_num).first():
                        v_num = generate_voucher_no("Contra")
                    je = JournalEntry(
                        voucher_no=v_num,
                        voucher_type="Contra",
                        date=c_date,
                        reference="Cash Contra",
                        narration=remarks or "Cash Contra transfer",
                        total_amount=round(p_total or r_total, 2)
                    )
                    db.session.add(je)
                    db.session.flush()
                    if r_cash > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars="Cash received"))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=0.0, credit=round(r_cash, 2), particulars="Premier Bank withdrawal"))
                    elif r_prem > 0 and cash_acc and prem_acc:
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars="Premier Bank received"))
                        db.session.add(TransactionLine(journal_entry_id=je.id, account_id=cash_acc.id, debit=0.0, credit=round(r_prem, 2), particulars="Cash disbursed"))

                    # If there was ALSO a payment on that line (e.g. Line 1533 Conveyance 140)
                    if p_code and p_code != "BS-003" and p_total > 0:
                        p_key = p_code if p_code not in ["-", "", "#N/A"] else "IS-016"
                        if p_key not in acc_dict:
                            new_acc = Account(code=p_key, name=p_desc or p_key, type="Expense", category="Operating Expense")
                            db.session.add(new_acc)
                            db.session.flush()
                            acc_dict[p_key] = new_acc
                        target_acc = acc_dict[p_key]
                        v_num = generate_voucher_no('Payment')
                        je_pmt = JournalEntry(
                            voucher_no=v_num,
                            voucher_type="Payment",
                            date=c_date,
                            reference=p_ref or None,
                            narration=remarks or f"Payment for {p_desc or target_acc.name}",
                            total_amount=round(p_total, 2)
                        )
                        db.session.add(je_pmt)
                        db.session.flush()
                        db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=p_desc))
                        if p_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars="Cash Payment"))
                        if p_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars="Premier Bank Payment"))
                else:
                    # 1. Post Receipt Voucher
                    if r_total > 0 and not is_open_row:
                        r_key = r_code if r_code and r_code not in ["-", "", "#N/A"] else "BS-035"
                        if r_key not in acc_dict:
                            new_acc = Account(code=r_key, name=r_desc or "Accounts-TBN-25", type="Revenue", category="Operating Revenue")
                            db.session.add(new_acc)
                            db.session.flush()
                            acc_dict[r_key] = new_acc

                        source_acc = acc_dict[r_key]
                        v_num = r_voucher if r_voucher and r_voucher not in ["-", "", "#N/A"] else None
                        if not v_num or JournalEntry.query.filter_by(voucher_no=v_num).first():
                            v_num = generate_voucher_no('Receipt')
                        v_date = r_date or date.today()
                        if isinstance(v_date, datetime):
                            v_date = v_date.date()
                        je_rcpt = JournalEntry(
                            voucher_no=v_num,
                            voucher_type="Receipt",
                            date=v_date,
                            reference=r_ref or None,
                            narration=remarks if (remarks and p_total == 0) else f"Receipt from {r_desc or source_acc.name}",
                            total_amount=round(r_total, 2)
                        )
                        db.session.add(je_rcpt)
                        db.session.flush()
                        if r_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_rcpt.id, account_id=cash_acc.id, debit=round(r_cash, 2), credit=0.0, particulars="Cash Receipt"))
                        if r_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_rcpt.id, account_id=prem_acc.id, debit=round(r_prem, 2), credit=0.0, particulars="Premier Bank Receipt"))
                        db.session.add(TransactionLine(journal_entry_id=je_rcpt.id, account_id=source_acc.id, debit=0.0, credit=round(r_total, 2), particulars=r_desc or "Receipt Credit"))

                    # 2. Post Payment Voucher
                    if p_total > 0:
                        p_key = p_code if p_code and p_code not in ["-", "", "#N/A"] else "IS-016"
                        if p_key not in acc_dict:
                            new_acc = Account(code=p_key, name=p_desc or p_key, type="Expense", category="Operating Expense")
                            db.session.add(new_acc)
                            db.session.flush()
                            acc_dict[p_key] = new_acc

                        target_acc = acc_dict[p_key]
                        v_num = p_vno if p_vno and p_vno not in ["-", "", "#N/A"] else None
                        if not v_num or JournalEntry.query.filter_by(voucher_no=v_num).first():
                            v_num = generate_voucher_no('Payment')
                        v_date = p_date or date.today()
                        if isinstance(v_date, datetime):
                            v_date = v_date.date()
                        je_pmt = JournalEntry(
                            voucher_no=v_num,
                            voucher_type="Payment",
                            date=v_date,
                            reference=p_ref or None,
                            narration=remarks or f"Payment for {p_desc or target_acc.name}",
                            total_amount=round(p_total, 2)
                        )
                        db.session.add(je_pmt)
                        db.session.flush()
                        db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=target_acc.id, debit=round(p_total, 2), credit=0.0, particulars=p_desc or "Expense"))
                        if p_cash > 0 and cash_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=cash_acc.id, debit=0.0, credit=round(p_cash, 2), particulars="Cash Payment"))
                        if p_prem > 0 and prem_acc:
                            db.session.add(TransactionLine(journal_entry_id=je_pmt.id, account_id=prem_acc.id, debit=0.0, credit=round(p_prem, 2), particulars="Premier Bank Payment"))

                imported_count += 1

        db.session.commit()
        recalculate_cash_book_running_balances()
        recalculate_account_balances()

        flash(f'Date condition ({from_date} to {to_date}) executed successfully! Cleared {cleared_jes} previous voucher(s), {cleared_lines} ledger line(s), and {cleared_cb_rows} cash book row(s). Updated {imported_count} new record(s) into General Ledger and Cash Book.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Error during date condition update: {str(e)}', 'danger')

    month_arg = from_date.strftime('%Y-%m')
    return redirect(url_for('cash_book_report', month=month_arg))

# 1. TRIAL BALANCE
@app.route('/reports/trial-balance')
def trial_balance():
    as_of_date_str = request.args.get('as_of_date', date.today().strftime('%Y-%m-%d'))
    try:
        as_of_date = datetime.strptime(as_of_date_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        as_of_date = date.today()

    accounts = Account.query.order_by(Account.code).all()
    rows = []
    total_period_dr = 0.0
    total_period_cr = 0.0
    total_closing_dr = 0.0
    total_closing_cr = 0.0

    for acc in accounts:
        period_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= as_of_date).scalar()

        period_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= as_of_date).scalar()

        total_period_dr += period_dr
        total_period_cr += period_cr

        if acc.type in ['Asset', 'Expense', 'Contra']:
            net = acc.opening_balance + period_dr - period_cr
            closing_dr = net if net >= 0 else 0.0
            closing_cr = abs(net) if net < 0 else 0.0
        else: # Liability, Equity, Revenue
            net = acc.opening_balance + period_cr - period_dr
            closing_cr = net if net >= 0 else 0.0
            closing_dr = abs(net) if net < 0 else 0.0

        total_period_dr += period_dr
        total_period_cr += period_cr
        total_closing_dr += closing_dr
        total_closing_cr += closing_cr

        rows.append({
            'account': acc,
            'period_dr': period_dr,
            'period_cr': period_cr,
            'closing_dr': closing_dr,
            'closing_cr': closing_cr
        })

    diff = round(abs(total_closing_dr - total_closing_cr), 2)
    is_balanced = (diff < 0.01)

    return render_template(
        'reports/trial_balance.html',
        as_of_date=as_of_date.strftime('%Y-%m-%d'),
        as_of_date_obj=as_of_date,
        rows=rows,
        total_period_dr=total_period_dr,
        total_period_cr=total_period_cr,
        total_closing_dr=total_closing_dr,
        total_closing_cr=total_closing_cr,
        is_balanced=is_balanced,
        diff=diff
    )

# 2. RECEIPTS AND PAYMENTS STATEMENT
@app.route('/reports/receipts-payments')
def receipts_payments():
    default_from = date(date.today().year, 1, 1).strftime('%Y-%m-%d')
    default_to = date.today().strftime('%Y-%m-%d')

    from_str = request.args.get('from_date', default_from)
    to_str = request.args.get('to_date', default_to)

    try:
        from_date = datetime.strptime(from_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        from_date = date(date.today().year, 1, 1)

    try:
        to_date = datetime.strptime(to_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        to_date = date.today()

    cash_bank_accounts = Account.query.filter_by(is_cash_or_bank=True).all()
    cash_bank_ids = [a.id for a in cash_bank_accounts]

    opening_breakdown = []
    total_opening = 0.0
    for acc in cash_bank_accounts:
        prior_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date < from_date).scalar()
        prior_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date < from_date).scalar()

        acc_opening = acc.opening_balance + prior_dr - prior_cr
        opening_breakdown.append({'name': acc.name, 'balance': acc_opening})
        total_opening += acc_opening

    receipts_dict = {}
    inflow_vouchers = JournalEntry.query.join(TransactionLine)\
        .filter(
            TransactionLine.account_id.in_(cash_bank_ids),
            TransactionLine.debit > 0,
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        ).distinct().all()

    for v in inflow_vouchers:
        cash_dr_amount = sum(l.debit for l in v.lines if l.account_id in cash_bank_ids)
        credit_lines = [l for l in v.lines if l.credit > 0 and l.account_id not in cash_bank_ids]
        total_credit_lines = sum(l.credit for l in credit_lines)

        for cl in credit_lines:
            weight = cl.credit / total_credit_lines if total_credit_lines > 0 else 1.0
            allocated_amount = cash_dr_amount * weight
            key = (cl.account.name, cl.account.type)
            receipts_dict[key] = receipts_dict.get(key, 0.0) + allocated_amount

    receipts = [{'account_name': k[0], 'type': k[1], 'amount': amt} for k, amt in receipts_dict.items()]
    total_receipts = sum(r['amount'] for r in receipts)

    payments_dict = {}
    outflow_vouchers = JournalEntry.query.join(TransactionLine)\
        .filter(
            TransactionLine.account_id.in_(cash_bank_ids),
            TransactionLine.credit > 0,
            JournalEntry.date >= from_date,
            JournalEntry.date <= to_date
        ).distinct().all()

    for v in outflow_vouchers:
        cash_cr_amount = sum(l.credit for l in v.lines if l.account_id in cash_bank_ids)
        debit_lines = [l for l in v.lines if l.debit > 0 and l.account_id not in cash_bank_ids]
        total_debit_lines = sum(l.debit for l in debit_lines)

        for dl in debit_lines:
            weight = dl.debit / total_debit_lines if total_debit_lines > 0 else 1.0
            allocated_amount = cash_cr_amount * weight
            key = (dl.account.name, dl.account.type)
            payments_dict[key] = payments_dict.get(key, 0.0) + allocated_amount

    payments = [{'account_name': k[0], 'type': k[1], 'amount': amt} for k, amt in payments_dict.items()]
    total_payments = sum(p['amount'] for p in payments)

    closing_breakdown = []
    total_closing = 0.0
    for acc in cash_bank_accounts:
        net_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= to_date).scalar()
        net_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= to_date).scalar()

        acc_closing = acc.opening_balance + net_dr - net_cr
        closing_breakdown.append({'name': acc.name, 'balance': acc_closing})
        total_closing += acc_closing

    return render_template(
        'reports/receipts_payments.html',
        from_date=from_date.strftime('%Y-%m-%d'),
        to_date=to_date.strftime('%Y-%m-%d'),
        opening_breakdown=opening_breakdown,
        total_opening=total_opening,
        receipts=receipts,
        total_receipts=total_receipts,
        payments=payments,
        total_payments=total_payments,
        closing_breakdown=closing_breakdown,
        total_closing=total_closing
    )

# 3. INCOME & EXPENDITURE STATEMENT (P&L)
@app.route('/reports/income-expenditure')
def income_expenditure():
    default_from = date(date.today().year, 1, 1).strftime('%Y-%m-%d')
    default_to = date.today().strftime('%Y-%m-%d')

    from_str = request.args.get('from_date', default_from)
    to_str = request.args.get('to_date', default_to)

    try:
        from_date = datetime.strptime(from_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        from_date = date(date.today().year, 1, 1)

    try:
        to_date = datetime.strptime(to_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        to_date = date.today()

    revenue_accounts = Account.query.filter_by(type='Revenue').order_by(Account.code).all()
    income_operating = []
    income_other = []
    total_income = 0.0

    for acc in revenue_accounts:
        dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date >= from_date, JournalEntry.date <= to_date).scalar()
        cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date >= from_date, JournalEntry.date <= to_date).scalar()
        net = round(cr - dr, 2)
        if net != 0:
            item = {'code': acc.code, 'name': acc.name, 'amount': net}
            if 'Operating' in acc.category:
                income_operating.append(item)
            else:
                income_other.append(item)
            total_income += net

    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.code).all()
    exp_operating = []
    exp_admin = []
    exp_financial = []
    total_expenditure = 0.0

    for acc in expense_accounts:
        dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date >= from_date, JournalEntry.date <= to_date).scalar()
        cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date >= from_date, JournalEntry.date <= to_date).scalar()
        net = round(dr - cr, 2)
        if net != 0:
            item = {'code': acc.code, 'name': acc.name, 'amount': net}
            if 'Operating' in acc.category:
                exp_operating.append(item)
            elif 'Financial' in acc.category:
                exp_financial.append(item)
            else:
                exp_admin.append(item)
            total_expenditure += net

    total_income = round(total_income, 2)
    total_expenditure = round(total_expenditure, 2)
    net_surplus = round(total_income - total_expenditure, 2)

    return render_template(
        'reports/income_expenditure.html',
        from_date=from_date.strftime('%Y-%m-%d'),
        to_date=to_date.strftime('%Y-%m-%d'),
        income_operating=income_operating,
        income_other=income_other,
        total_income=total_income,
        exp_operating=exp_operating,
        exp_admin=exp_admin,
        exp_financial=exp_financial,
        total_expenditure=total_expenditure,
        net_surplus=net_surplus
    )

# 4. BALANCE SHEET
@app.route('/reports/balance-sheet')
def balance_sheet():
    as_of_date_str = request.args.get('as_of_date', date.today().strftime('%Y-%m-%d'))
    try:
        as_of_date = datetime.strptime(as_of_date_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        as_of_date = date.today()

    accounts = Account.query.order_by(Account.code).all()

    current_assets = []
    non_current_assets = []
    current_liabilities = []
    non_current_liabilities = []
    equity_accounts = []

    total_current_assets = 0.0
    total_non_current_assets = 0.0
    total_current_liabilities = 0.0
    total_non_current_liabilities = 0.0
    total_equity_direct = 0.0

    for acc in accounts:
        dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= as_of_date).scalar()
        cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
            .join(JournalEntry)\
            .filter(TransactionLine.account_id == acc.id, JournalEntry.date <= as_of_date).scalar()

        if acc.type in ['Asset', 'Contra']:
            bal = round(acc.opening_balance + dr - cr, 2)
            item = {'code': acc.code, 'name': acc.name, 'balance': bal}
            if 'Non-Current' in acc.category or 'Fixed' in acc.category:
                non_current_assets.append(item)
                total_non_current_assets += bal
            else:
                current_assets.append(item)
                total_current_assets += bal

        elif acc.type == 'Liability':
            bal = round(acc.opening_balance + cr - dr, 2)
            item = {'code': acc.code, 'name': acc.name, 'balance': bal}
            if 'Non-Current' in acc.category:
                non_current_liabilities.append(item)
                total_non_current_liabilities += bal
            else:
                current_liabilities.append(item)
                total_current_liabilities += bal

        elif acc.type == 'Equity':
            bal = round(acc.opening_balance + cr - dr, 2)
            equity_accounts.append({'code': acc.code, 'name': acc.name, 'balance': bal})
            total_equity_direct += bal

    rev_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
        .join(JournalEntry).join(Account)\
        .filter(Account.type == 'Revenue', JournalEntry.date <= as_of_date).scalar()
    rev_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
        .join(JournalEntry).join(Account)\
        .filter(Account.type == 'Revenue', JournalEntry.date <= as_of_date).scalar()
    total_rev_period = rev_cr - rev_dr

    exp_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
        .join(JournalEntry).join(Account)\
        .filter(Account.type == 'Expense', JournalEntry.date <= as_of_date).scalar()
    exp_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
        .join(JournalEntry).join(Account)\
        .filter(Account.type == 'Expense', JournalEntry.date <= as_of_date).scalar()
    total_exp_period = exp_dr - exp_cr

    current_surplus = round(total_rev_period - total_exp_period, 2)

    total_assets = round(total_current_assets + total_non_current_assets, 2)
    total_liabilities = round(total_current_liabilities + total_non_current_liabilities, 2)
    total_equity = round(total_equity_direct + current_surplus, 2)
    total_liabilities_and_equity = round(total_liabilities + total_equity, 2)

    diff = round(abs(total_assets - total_liabilities_and_equity), 2)
    is_balanced = (diff < 0.01)

    return render_template(
        'reports/balance_sheet.html',
        as_of_date=as_of_date.strftime('%Y-%m-%d'),
        as_of_date_obj=as_of_date,
        current_assets=current_assets,
        non_current_assets=non_current_assets,
        total_current_assets=total_current_assets,
        total_non_current_assets=total_non_current_assets,
        total_assets=total_assets,
        current_liabilities=current_liabilities,
        non_current_liabilities=non_current_liabilities,
        total_current_liabilities=total_current_liabilities,
        total_non_current_liabilities=total_non_current_liabilities,
        total_liabilities=total_liabilities,
        equity_accounts=equity_accounts,
        current_surplus=current_surplus,
        total_equity=total_equity,
        total_liabilities_and_equity=total_liabilities_and_equity,
        is_balanced=is_balanced,
        diff=diff
    )


# ==============================================================================
# BANK STATEMENT PDF PARSING & REPORTING MODULE
# ==============================================================================

def extract_bank_statement_pdf(stream):
    """
    Intelligently extracts transactions from a bank statement PDF using pypdf.
    Handles multi-line wrapped transaction descriptions, multiple pages,
    dynamic debit/credit/balance detection, reference number extraction,
    and boilerplate header/footer filtering.
    """
    reader = pypdf.PdfReader(stream)
    start_row_regex = re.compile(r'^\s*(?:(\d+)\s+)?(\d{1,2}-[A-Za-z]{3}-\d{4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b')
    money_regex = re.compile(r'[-+]?\d{1,3}(?:,\d{3})*\.\d{2}')

    all_lines = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        all_lines.extend(txt.split('\n'))

    tx_blocks = []
    current_block = None

    for line in all_lines:
        s = line.strip()
        if not s:
            continue
        # Skip header / footer lines that are not part of transaction rows
        if any(h in s for h in ['Account Statement', 'Branch Name', 'A/C No', 'Statement Period', 'Page 1 of', 'Total Debit :', 'Total Credit :']):
            if 'Total Debit' in s or 'Total Credit' in s:
                if current_block:
                    tx_blocks.append(current_block)
                    current_block = None
            continue
        if s.startswith('SI Date Narration') or s.startswith('Sl Date Particulars') or ('Opening Balance' in s and not start_row_regex.match(s)):
            continue

        m = start_row_regex.match(s)
        if m:
            if current_block:
                tx_blocks.append(current_block)
            current_block = [s]
        elif current_block:
            current_block.append(s)

    if current_block:
        tx_blocks.append(current_block)

    boilerplates = [
        r'Address\s*:.*?(?:BDT|Currency\s*:\s*[A-Z]+)',
        r'Mobile\s*:.*?(?:REGULAR|Status\s*:\s*[A-Z]+)',
        r'Notification\s*:.*?(?:Branches\)|Branches)',
        r'N:B\.?\s*:.*?(?:signature\.|signature)',
        r'Should have any query.*?(?:Branch\.|Branch)',
        r'Page\s+\d+\s+of\s+\d+',
        r'Branch Name\s*:.*',
        r'A/C No\s*:.*',
        r'A/C Type\s*:.*',
        r'Statement Period.*?(?:AM|PM|\d{4})',
        r'Generation Date.*?(?:AM|PM|\d{4})'
    ]

    parsed_items = []
    prev_balance = None

    for idx, b in enumerate(tx_blocks):
        full_text = ' '.join(b)
        m = start_row_regex.match(full_text)
        if not m:
            continue

        d_str = m.group(2)
        parsed_date = parse_accounting_date(d_str)
        if not parsed_date:
            continue

        nums = money_regex.findall(full_text)
        if not nums:
            continue

        clean_nums = []
        for x in nums:
            try:
                clean_nums.append(float(x.replace(',', '')))
            except Exception:
                pass

        if not clean_nums:
            continue

        text_after_date = full_text[m.end():].strip()

        w_dr = 0.0
        d_cr = 0.0
        bal = 0.0

        if len(clean_nums) >= 3:
            w_dr = clean_nums[-3]
            d_cr = clean_nums[-2]
            bal = clean_nums[-1]
            amt_strings = nums[-3:]
        elif len(clean_nums) == 2:
            amt = clean_nums[-2]
            bal = clean_nums[-1]
            amt_strings = nums[-2:]
            if prev_balance is not None:
                if bal < prev_balance:
                    w_dr = amt
                else:
                    d_cr = amt
            else:
                is_dr = any(k in full_text.upper() for k in ['DR', 'DEBIT', 'WITHDRAWAL', 'PAYMENT', 'PURCHASE', 'FEE', 'CHG'])
                if is_dr:
                    w_dr = amt
                else:
                    d_cr = amt
        else:
            amt = clean_nums[-1]
            amt_strings = nums[-1:]
            is_dr = any(k in full_text.upper() for k in ['DR', 'DEBIT', 'WITHDRAWAL', 'PAYMENT', 'PURCHASE', 'FEE', 'CHG'])
            if is_dr:
                w_dr = amt
            else:
                d_cr = amt
            bal = prev_balance if prev_balance is not None else 0.0

        prev_balance = bal

        # Clean description text
        narration = text_after_date
        for a_s in amt_strings:
            r_pos = narration.rfind(a_s)
            if r_pos != -1:
                narration = narration[:r_pos] + narration[r_pos+len(a_s):]

        for pat in boilerplates:
            narration = re.sub(pat, ' ', narration, flags=re.I)

        narration = re.sub(r'^\s*\d+\s+', '', narration) # strip any leftover leading serial number
        narration = re.sub(r'\s+', ' ', narration).strip()

        if not narration and w_dr == 0 and d_cr == 0:
            narration = "Opening Balance"

        ref_m = re.search(r'(?:RRN|INST NO\.?|TRX|REF|CHQ|NO\.?)[-#\s]*([A-Za-z0-9]+)', narration, re.I)
        cheque_ref = ref_m.group(0).strip() if ref_m else ""

        parsed_items.append({
            'date': parsed_date,
            'description': narration or "Bank Transaction",
            'cheque_ref': cheque_ref,
            'withdrawal': w_dr,
            'deposit': d_cr,
            'balance': bal
        })

    # Fallback to single-line scanner if no blocks were matched
    if not parsed_items:
        date_regex = re.compile(r'(\d{1,2}[-/.]\w{3,}[-/.]\d{2,4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})')
        for line in all_lines:
            line_str = line.strip()
            if not line_str:
                continue
            d_match = date_regex.search(line_str)
            if d_match:
                d_str = d_match.group(1)
                parsed_date = parse_accounting_date(d_str)
                if not parsed_date:
                    continue
                nums = money_regex.findall(line_str)
                if not nums:
                    continue
                clean_nums = []
                for num_s in nums:
                    try:
                        clean_nums.append(float(num_s.replace(',', '')))
                    except Exception:
                        pass
                if not clean_nums:
                    continue
                clean_line = line_str.replace(d_str, ' ')
                for num_s in nums:
                    clean_line = clean_line.replace(num_s, ' ')
                desc = re.sub(r'\s+', ' ', clean_line).strip()
                ref_match = re.search(r'(?:CHQ|REF|TRX|INV|VOUCHER|FT|TT)[-#\s]*[A-Za-z0-9]+|\b\d{4,8}\b', desc, re.I)
                cheque_ref = ref_match.group(0).strip() if ref_match else ""
                w_dr = clean_nums[0] if len(clean_nums) >= 2 else 0.0
                d_cr = clean_nums[1] if len(clean_nums) >= 3 else 0.0
                bal = clean_nums[-1] if clean_nums else 0.0
                parsed_items.append({
                    'date': parsed_date,
                    'description': desc or "Bank Transaction",
                    'cheque_ref': cheque_ref,
                    'withdrawal': w_dr,
                    'deposit': d_cr,
                    'balance': bal
                })

    return parsed_items


def extract_transactions_from_file(file_storage):
    """
    Extracts transactions from a single PDF, Excel (.xlsx, .xls), or CSV file.
    Normalizes: date, description, cheque_ref, withdrawal, deposit, balance, source_file.
    """
    filename = file_storage.filename or ""
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ""
    items = []

    if ext == 'pdf':
        items = extract_bank_statement_pdf(file_storage.stream)
        for it in items:
            it['source_file'] = filename
    elif ext in ('xlsx', 'xls'):
        wb = openpyxl.load_workbook(file_storage.stream, data_only=True)
        ws = wb.active

        # Detect header row dynamically
        header_row_idx = 1
        col_map = {'date': 0, 'desc': 1, 'ref': 2, 'wdr': 3, 'dep': 4, 'bal': 5}
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if r_idx > 12:
                break
            row_strs = [str(c or '').lower().strip() for c in row]
            if any('date' in s for s in row_strs) and any(('particular' in s or 'desc' in s or 'detail' in s or 'narration' in s or 'transaction' in s) for s in row_strs):
                header_row_idx = r_idx
                for c_idx, s in enumerate(row_strs):
                    if 'date' in s and 'val' not in s:
                        col_map['date'] = c_idx
                    elif any(k in s for k in ['particular', 'desc', 'detail', 'narration', 'transaction']):
                        col_map['desc'] = c_idx
                    elif any(k in s for k in ['cheque', 'chq', 'ref', 'instrument', 'trn']):
                        col_map['ref'] = c_idx
                    elif any(k in s for k in ['withdraw', 'debit', 'dr', 'payment']):
                        col_map['wdr'] = c_idx
                    elif any(k in s for k in ['deposit', 'credit', 'cr', 'receipt']):
                        col_map['dep'] = c_idx
                    elif any(k in s for k in ['balance', 'bal']):
                        col_map['bal'] = c_idx
                break

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not row or not any(row):
                continue
            d_val = row[col_map['date']] if len(row) > col_map['date'] else None
            d = parse_accounting_date(d_val)
            if not d:
                continue
            desc = str(row[col_map['desc']] or 'Bank Transaction') if len(row) > col_map['desc'] and row[col_map['desc']] is not None else 'Bank Transaction'
            ref = str(row[col_map['ref']] or '') if len(row) > col_map['ref'] and row[col_map['ref']] is not None else ''
            w = clean_accounting_num(row[col_map['wdr']]) if len(row) > col_map['wdr'] else 0.0
            dep = clean_accounting_num(row[col_map['dep']]) if len(row) > col_map['dep'] else 0.0
            bal = clean_accounting_num(row[col_map['bal']]) if len(row) > col_map['bal'] else 0.0
            items.append({
                'date': d,
                'description': desc.strip(),
                'cheque_ref': ref.strip(),
                'withdrawal': w,
                'deposit': dep,
                'balance': bal,
                'source_file': filename
            })

    elif ext == 'csv':
        content = file_storage.stream.read().decode('utf-8', errors='ignore')
        reader = csv.reader(content.splitlines())
        rows = list(reader)
        if rows:
            header_idx = 0
            col_map = {'date': 0, 'desc': 1, 'ref': 2, 'wdr': 3, 'dep': 4, 'bal': 5}
            for r_idx, row in enumerate(rows[:12]):
                row_strs = [s.lower().strip() for s in row]
                if any('date' in s for s in row_strs):
                    header_idx = r_idx
                    for c_idx, s in enumerate(row_strs):
                        if 'date' in s and 'val' not in s: col_map['date'] = c_idx
                        elif any(k in s for k in ['particular', 'desc', 'detail', 'narration']): col_map['desc'] = c_idx
                        elif any(k in s for k in ['cheque', 'chq', 'ref']): col_map['ref'] = c_idx
                        elif any(k in s for k in ['withdraw', 'debit', 'dr']): col_map['wdr'] = c_idx
                        elif any(k in s for k in ['deposit', 'credit', 'cr']): col_map['dep'] = c_idx
                        elif any(k in s for k in ['balance', 'bal']): col_map['bal'] = c_idx
                    break
            for row in rows[header_idx + 1:]:
                if not row or not any(row):
                    continue
                d = parse_accounting_date(row[col_map['date']]) if len(row) > col_map['date'] else None
                if not d:
                    continue
                desc = row[col_map['desc']] if len(row) > col_map['desc'] else 'Transaction'
                ref = row[col_map['ref']] if len(row) > col_map['ref'] else ''
                w = clean_accounting_num(row[col_map['wdr']]) if len(row) > col_map['wdr'] else 0.0
                dep = clean_accounting_num(row[col_map['dep']]) if len(row) > col_map['dep'] else 0.0
                bal = clean_accounting_num(row[col_map['bal']]) if len(row) > col_map['bal'] else 0.0
                items.append({
                    'date': d,
                    'description': desc.strip(),
                    'cheque_ref': ref.strip(),
                    'withdrawal': w,
                    'deposit': dep,
                    'balance': bal,
                    'source_file': filename
                })

    return items


def intelligent_analyze_statement_items(raw_items, bank_account_id):
    """
    Intelligent Pre-Posting Engine:
    - Detects duplicates against database records (preventing double posting)
    - Detects intra-batch duplicates across multiple files
    - Verifies running balance arithmetic
    - Automatically classifies items for safe posting
    """
    # Sort chronologically
    sorted_items = sorted(raw_items, key=lambda x: x['date'] if x['date'] else date.min)

    existing_lines = BankStatementLine.query.filter_by(bank_account_id=bank_account_id).all() if bank_account_id else []

    analyzed = []
    seen_in_batch = set()
    dup_count = 0
    clean_count = 0
    math_warnings_count = 0

    prev_balance = None

    for it in sorted_items:
        d = it['date']
        w = round(float(it.get('withdrawal') or 0.0), 2)
        dep = round(float(it.get('deposit') or 0.0), 2)
        bal = round(float(it.get('balance') or 0.0), 2)
        desc = it.get('description', 'Bank Transaction')
        ref = it.get('cheque_ref', '')
        src = it.get('source_file', 'Uploaded File')

        is_dup = False
        dup_reason = ""

        # 1. Match against existing DB records for this bank account
        for ex in existing_lines:
            if ex.txn_date == d and abs(ex.withdrawal_dr - w) < 0.01 and abs(ex.deposit_cr - dep) < 0.01:
                # If cheque reference matches and is present
                if ref and ex.cheque_ref and ref.lower().strip() == ex.cheque_ref.lower().strip():
                    is_dup = True
                    dup_reason = f"Duplicate: Ref #{ref} already posted in line #{ex.id} on {ex.txn_date.strftime('%d-%b-%Y')}"
                    break
                # If particulars/description matches closely
                cd_it = re.sub(r'[^a-zA-Z0-9]', '', desc.lower())
                cd_ex = re.sub(r'[^a-zA-Z0-9]', '', ex.particulars.lower())
                if cd_it and (cd_it == cd_ex or (len(cd_it) > 10 and (cd_it in cd_ex or cd_ex in cd_it))):
                    is_dup = True
                    dup_reason = f"Duplicate: Exact transaction already recorded in line #{ex.id} on {ex.txn_date.strftime('%d-%b-%Y')}"
                    break

        # 2. Check intra-batch duplication across uploaded files
        batch_key = (d, w, dep, ref.strip().lower(), desc[:35].strip().lower())
        if not is_dup:
            if batch_key in seen_in_batch:
                is_dup = True
                dup_reason = "Duplicate: Identical entry repeated within uploaded batch files"
            else:
                seen_in_batch.add(batch_key)

        # 3. Arithmetic balance continuity verification
        has_math_warning = False
        math_msg = ""
        if prev_balance is not None and bal > 0:
            expected_bal = round(prev_balance + dep - w, 2)
            if abs(bal - expected_bal) > 1.0:
                has_math_warning = True
                math_msg = f"Balance step jump (Expected: ৳{expected_bal:,.2f}, File stated: ৳{bal:,.2f})"
                math_warnings_count += 1
        if bal > 0:
            prev_balance = bal

        if is_dup:
            dup_count += 1
            status = 'duplicate'
            selected = False
        else:
            clean_count += 1
            status = 'clean' if not has_math_warning else 'warning'
            selected = True

        analyzed.append({
            'date': d.strftime('%Y-%m-%d') if d else '',
            'date_display': d.strftime('%d-%b-%Y') if d else '',
            'description': desc,
            'cheque_ref': ref,
            'withdrawal': w,
            'deposit': dep,
            'balance': bal,
            'source_file': src,
            'is_duplicate': is_dup,
            'dup_reason': dup_reason,
            'has_math_warning': has_math_warning,
            'math_msg': math_msg,
            'status': status,
            'selected': selected,
            'notes': ''
        })

    return {
        'rows': analyzed,
        'items': analyzed,
        'total_rows': len(sorted_items),
        'clean_count': clean_count,
        'dup_count': dup_count,
        'math_warnings_count': math_warnings_count
    }


@app.route('/reports/bank-statement')
@app.route('/banking/statements')
def bank_statement_report():
    # 1. Fetch available bank accounts (excluding revenue/expense)
    bank_accounts = Account.query.filter(
        db.or_(
            Account.category == 'Bank',
            Account.is_cash_or_bank == True,
            Account.name.ilike('%bank%')
        ),
        Account.type.notin_(['Expense', 'Revenue'])
    ).order_by(Account.code.asc()).all()

    # Also include any accounts that already have bank statements in BankStatement or BankStatementLine
    stmt_account_ids = [r[0] for r in db.session.query(BankStatement.bank_account_id).distinct().all() if r[0]]
    if stmt_account_ids:
        additional_accounts = Account.query.filter(Account.id.in_(stmt_account_ids)).all()
        for a in additional_accounts:
            if a not in bank_accounts:
                bank_accounts.append(a)

    # Deduplicate accounts by ID
    unique_banks = []
    seen = set()
    for b in bank_accounts:
        if b.id not in seen:
            seen.add(b.id)
            unique_banks.append(b)
    bank_accounts = unique_banks

    selected_acc_id = request.args.get('account_id', type=int)
    if selected_acc_id is None and bank_accounts:
        # Default to Premier Bank if available, else first bank account
        premier = next((a for a in bank_accounts if 'Premier' in a.name), bank_accounts[0])
        selected_acc_id = premier.id

    selected_acc = Account.query.get(selected_acc_id) if (selected_acc_id and selected_acc_id > 0) else None

    # Recon mode: '1' = Reconcile with Accounts Bank Ledger; '0' = Standalone External Statement
    recon_mode_param = request.args.get('recon_mode', '')
    if recon_mode_param == '':
        recon_mode = True if selected_acc else False
    else:
        recon_mode = (recon_mode_param == '1')

    from_str = request.args.get('from_date', '')
    to_str = request.args.get('to_date', '')
    search = request.args.get('search', '').strip()

    query = BankStatementLine.query
    if selected_acc_id and selected_acc_id > 0:
        query = query.filter(BankStatementLine.bank_account_id == selected_acc_id)

    if from_str:
        f_d = parse_accounting_date(from_str)
        if f_d:
            query = query.filter(BankStatementLine.txn_date >= f_d)
    if to_str:
        t_d = parse_accounting_date(to_str)
        if t_d:
            query = query.filter(BankStatementLine.txn_date <= t_d)

    if search:
        s_pat = f"%{search}%"
        query = query.filter(
            db.or_(
                BankStatementLine.particulars.ilike(s_pat),
                BankStatementLine.cheque_ref.ilike(s_pat),
                BankStatementLine.notes.ilike(s_pat),
                BankStatementLine.matched_voucher_no.ilike(s_pat)
            )
        )

    lines = query.order_by(BankStatementLine.txn_date.asc(), BankStatementLine.id.asc()).all()

    total_withdrawals = sum(l.withdrawal_dr for l in lines)
    total_deposits = sum(l.deposit_cr for l in lines)
    reconciled_count = sum(1 for l in lines if l.is_reconciled)
    unreconciled_count = len(lines) - reconciled_count

    latest_line = lines[-1] if lines else None
    latest_balance = latest_line.balance if latest_line else (selected_acc.balance if selected_acc else 0.0)

    # General Ledger Comparison
    gl_book_balance = selected_acc.balance if selected_acc else 0.0
    reconciliation_variance = round(latest_balance - gl_book_balance, 2) if selected_acc else 0.0

    statements = BankStatement.query.filter_by(bank_account_id=selected_acc_id).order_by(BankStatement.from_date.desc()).all() if selected_acc_id else []

    return render_template(
        'banking/bank_statement_report.html',
        bank_accounts=bank_accounts,
        selected_account=selected_acc,
        recon_mode=recon_mode,
        lines=lines,
        statements=statements,
        total_withdrawals=total_withdrawals,
        total_deposits=total_deposits,
        reconciled_count=reconciled_count,
        unreconciled_count=unreconciled_count,
        latest_balance=latest_balance,
        gl_book_balance=gl_book_balance,
        reconciliation_variance=reconciliation_variance,
        from_date=from_str,
        to_date=to_str,
        search_query=search
    )


@app.route('/banking/banks/create', methods=['POST'])
def create_custom_bank_account():
    """
    Creates a new custom bank account for target bank statement tracking.
    Can be configured as a Standalone Statement Account (separate from GL)
    or linked with the Accounts General Ledger.
    """
    bank_name = request.form.get('bank_name', '').strip()
    bank_code = request.form.get('bank_code', '').strip()
    opening_bal = float(request.form.get('opening_balance') or 0.0)
    ledger_mode = request.form.get('ledger_mode', 'standalone') # 'standalone' or 'gl_linked'

    if not bank_name:
        flash('Bank Name is required to add a custom bank account.', 'danger')
        return redirect(url_for('bank_statement_report'))

    # Check if a bank account with this name already exists
    existing = Account.query.filter(func.lower(Account.name) == bank_name.lower()).first()
    if existing:
        flash(f'Bank "{existing.name}" ({existing.code}) already exists and has been selected.', 'info')
        return redirect(url_for('bank_statement_report', account_id=existing.id, recon_mode='0' if ledger_mode == 'standalone' else '1'))

    # Generate next code if blank
    if not bank_code:
        existing_codes = set(r[0] for r in db.session.query(Account.code).all() if r[0])
        idx = 1
        while f"BS-{idx:03d}" in existing_codes:
            idx += 1
        bank_code = f"BS-{idx:03d}"
    else:
        dup = Account.query.filter_by(code=bank_code).first()
        if dup:
            existing_codes = set(r[0] for r in db.session.query(Account.code).all() if r[0])
            idx = 1
            while f"BS-{idx:03d}" in existing_codes:
                idx += 1
            bank_code = f"BS-{idx:03d}"

    init_bal = opening_bal if ledger_mode == 'gl_linked' else 0.0
    new_bank = Account(
        code=bank_code,
        name=bank_name,
        type='Asset',
        category='Bank',
        is_cash_or_bank=True,
        opening_balance=init_bal,
        balance=init_bal
    )
    try:
        db.session.add(new_bank)
        db.session.commit()
        mode_label = "Standalone Statement Account (isolated from Ledger)" if ledger_mode == 'standalone' else "Linked with General Ledger"
        flash(f'Custom Target Bank "{bank_name}" [{bank_code}] created successfully as {mode_label}!', 'success')
        return redirect(url_for('bank_statement_report', account_id=new_bank.id, recon_mode='0' if ledger_mode == 'standalone' else '1'))
    except Exception as e:
        db.session.rollback()
        flash(f'Error creating custom target bank: {str(e)}', 'danger')
        return redirect(url_for('bank_statement_report'))


@app.route('/banking/statements/upload', methods=['POST'])
@app.route('/banking/statements/upload-pdf', methods=['POST'])
def upload_bank_statement_files():
    """
    Handles upload of ONE OR MORE bank statement files (PDF, Excel .xlsx/.xls, CSV).
    Uses Intelligent Analysis to detect and exclude duplicates to avoid double posting.
    Supports on-the-fly custom target bank creation!
    """
    uploaded_files = request.files.getlist('statement_files')
    if not uploaded_files or not any(f.filename for f in uploaded_files):
        # Fallback to single file input 'pdf_file' for backwards compatibility
        single_f = request.files.get('pdf_file')
        if single_f and single_f.filename:
            uploaded_files = [single_f]

    bank_account_raw = (request.form.get('bank_account_id') or '').strip()
    action_mode = request.form.get('action_mode', 'analyze') # 'analyze' or 'direct_post'
    custom_bank_name = request.form.get('custom_bank_name', '').strip()
    custom_bank_code = request.form.get('custom_bank_code', '').strip()
    custom_bank_bal = float(request.form.get('custom_bank_balance') or 0.0)
    custom_ledger_mode = request.form.get('custom_bank_ledger_mode', 'standalone')

    # If the user selected __new__ or entered a custom bank name:
    if bank_account_raw == '__new__' or (custom_bank_name and (not bank_account_raw or bank_account_raw == '0')):
        if not custom_bank_name:
            flash('Please enter a Bank Name for the new custom bank account.', 'danger')
            return redirect(url_for('bank_statement_report'))
        
        # Check if existing by name
        existing = Account.query.filter(func.lower(Account.name) == custom_bank_name.lower()).first()
        if existing:
            bank_acc = existing
            bank_account_id = existing.id
        else:
            existing_codes = set(r[0] for r in db.session.query(Account.code).all() if r[0])
            if not custom_bank_code or custom_bank_code in existing_codes:
                idx = 1
                while f"BS-{idx:03d}" in existing_codes:
                    idx += 1
                custom_bank_code = f"BS-{idx:03d}"
            
            init_bal = custom_bank_bal if custom_ledger_mode == 'gl_linked' else 0.0
            bank_acc = Account(
                code=custom_bank_code,
                name=custom_bank_name,
                type='Asset',
                category='Bank',
                is_cash_or_bank=True,
                opening_balance=init_bal,
                balance=init_bal
            )
            db.session.add(bank_acc)
            db.session.commit()
            bank_account_id = bank_acc.id
    else:
        try:
            bank_account_id = int(bank_account_raw)
        except (ValueError, TypeError):
            bank_account_id = None
        bank_acc = Account.query.get(bank_account_id) if bank_account_id else None

    if not bank_acc:
        flash('Invalid bank account selected. Please select a valid target bank account or add a new custom bank.', 'danger')
        return redirect(url_for('bank_statement_report'))

    try:
        raw_items = []
        filenames_processed = []

        for f in uploaded_files:
            if not f or not f.filename:
                continue
            filenames_processed.append(f.filename)
            extracted = extract_transactions_from_file(f)
            raw_items.extend(extracted)

        if not raw_items:
            flash(
                f'No tabular bank transactions could be extracted from the uploaded file(s) ({", ".join(filenames_processed)}). '
                'Please check that the file contains transaction dates and amounts, or add entries manually.',
                'warning'
            )
            return redirect(url_for('bank_statement_report', account_id=bank_account_id))

        # Run Intelligent Analysis
        analysis_result = intelligent_analyze_statement_items(raw_items, bank_account_id)

        # Option A: Direct Smart Post (Auto-skip duplicates)
        if action_mode == 'direct_post':
            clean_items = [it for it in analysis_result['rows'] if not it['is_duplicate']]
            if not clean_items:
                flash(
                    f'Intelligent Analysis Alert: All {len(raw_items)} transaction(s) in the uploaded file(s) already exist in {bank_acc.name}. '
                    'Zero records were posted to avoid double-posting.',
                    'info'
                )
                return redirect(url_for('bank_statement_report', account_id=bank_account_id))

            # Sort clean items
            clean_items.sort(key=lambda x: x['date'])
            from_d = parse_accounting_date(clean_items[0]['date']) or date.today()
            to_d = parse_accounting_date(clean_items[-1]['date']) or date.today()

            stmt_title = f"{bank_acc.name} Statement ({from_d.strftime('%b %Y')} - {to_d.strftime('%b %Y')})"
            tot_w = sum(it['withdrawal'] for it in clean_items)
            tot_d = sum(it['deposit'] for it in clean_items)
            close_b = clean_items[-1]['balance'] if clean_items else 0.0
            open_b = clean_items[0]['balance'] if clean_items else 0.0

            stmt = BankStatement(
                bank_account_id=bank_acc.id,
                statement_title=stmt_title,
                filename=" + ".join(filenames_processed[:3]) + (f" (+{len(filenames_processed)-3} more)" if len(filenames_processed) > 3 else ""),
                from_date=from_d,
                to_date=to_d,
                opening_balance=open_b,
                closing_balance=close_b,
                total_withdrawals=tot_w,
                total_deposits=tot_d,
                notes=f"Auto-imported via Smart Anti-Duplicate Engine from {len(filenames_processed)} file(s)."
            )
            db.session.add(stmt)
            db.session.flush()

            created_count = 0
            for it in clean_items:
                d_obj = parse_accounting_date(it['date']) or date.today()
                amount = it['deposit'] if it['deposit'] > 0 else it['withdrawal']
                matched_v = None
                if amount > 0:
                    is_dep = it['deposit'] > 0
                    matched_line = TransactionLine.query.join(JournalEntry).filter(
                        TransactionLine.account_id == bank_acc.id,
                        TransactionLine.debit == (amount if is_dep else 0.0),
                        TransactionLine.credit == (0.0 if is_dep else amount),
                        JournalEntry.date >= d_obj - timedelta(days=7),
                        JournalEntry.date <= d_obj + timedelta(days=7)
                    ).first()
                    if matched_line:
                        matched_v = matched_line.journal_entry

                b_line = BankStatementLine(
                    statement_id=stmt.id,
                    bank_account_id=bank_acc.id,
                    txn_date=d_obj,
                    value_date=d_obj,
                    particulars=it['description'][:490],
                    cheque_ref=it['cheque_ref'][:95],
                    withdrawal_dr=it['withdrawal'],
                    deposit_cr=it['deposit'],
                    balance=it['balance'],
                    notes=f"Source: {it.get('source_file', '')[:100]}",
                    is_reconciled=(matched_v is not None),
                    reconciled_date=d_obj if matched_v else None,
                    matched_voucher_id=matched_v.id if matched_v else None,
                    matched_voucher_no=matched_v.voucher_no if matched_v else None
                )
                db.session.add(b_line)
                created_count += 1

            db.session.commit()
            dup_msg = f" Automatically skipped {analysis_result['dup_count']} duplicate item(s) to avoid double posting." if analysis_result['dup_count'] > 0 else ""
            flash(
                f'Intelligent posting successful: Imported {created_count} clean transactions from {len(filenames_processed)} file(s) into {bank_acc.name}.{dup_msg}',
                'success'
            )
            return redirect(url_for('bank_statement_report', account_id=bank_account_id))

        # Option B: Interactive Pre-Posting Review & Analysis Screen
        all_bank_accounts = Account.query.filter(
            db.or_(
                Account.category == 'Bank',
                Account.is_cash_or_bank == True,
                Account.name.ilike('%bank%')
            ),
            Account.type.notin_(['Expense', 'Revenue'])
        ).order_by(Account.code.asc()).all()
        return render_template(
            'banking/bank_statement_analyze.html',
            bank_account=bank_acc,
            all_bank_accounts=all_bank_accounts,
            analysis=analysis_result,
            filenames=", ".join(filenames_processed),
            default_title=f"{bank_acc.name} Statement ({date.today().strftime('%b %Y')})"
        )

    except Exception as e:
        db.session.rollback()
        flash(f'Error analyzing bank statement file(s): {str(e)}', 'danger')
        return redirect(url_for('bank_statement_report', account_id=bank_account_id))


@app.route('/banking/statements/post-analyzed', methods=['POST'])
def post_analyzed_bank_statements():
    """
    Posts reviewed clean transactions from the Intelligent Analysis Screen into the bank statement report.
    """
    bank_account_id = request.form.get('bank_account_id', type=int)
    bank_acc = Account.query.get_or_404(bank_account_id)
    statement_title = request.form.get('statement_title', '').strip()
    statement_notes = request.form.get('statement_notes', '').strip()
    filename_summary = request.form.get('filename_summary', 'Multi-Source Import')

    # Read reviewed rows
    selected_indices = request.form.getlist('selected_rows[]')
    dates = request.form.getlist('row_date[]')
    descs = request.form.getlist('row_desc[]')
    refs = request.form.getlist('row_ref[]')
    wdrs = request.form.getlist('row_wdr[]')
    deps = request.form.getlist('row_dep[]')
    bals = request.form.getlist('row_bal[]')
    notes_list = request.form.getlist('row_notes[]')

    selected_set = set(selected_indices)
    clean_items = []

    for i in range(len(dates)):
        if str(i) in selected_set:
            d_obj = parse_accounting_date(dates[i]) or date.today()
            clean_items.append({
                'date': d_obj,
                'description': descs[i].strip() if i < len(descs) else 'Bank Transaction',
                'cheque_ref': refs[i].strip() if i < len(refs) else '',
                'withdrawal': float(wdrs[i] or 0.0) if i < len(wdrs) else 0.0,
                'deposit': float(deps[i] or 0.0) if i < len(deps) else 0.0,
                'balance': float(bals[i] or 0.0) if i < len(bals) else 0.0,
                'notes': notes_list[i].strip() if i < len(notes_list) else ''
            })

    if not clean_items:
        flash('No transactions were selected for posting. Bank statement report unchanged.', 'warning')
        return redirect(url_for('bank_statement_report', account_id=bank_account_id))

    clean_items.sort(key=lambda x: x['date'])
    from_d = clean_items[0]['date']
    to_d = clean_items[-1]['date']

    if not statement_title:
        statement_title = f"{bank_acc.name} Statement ({from_d.strftime('%b %Y')} - {to_d.strftime('%b %Y')})"

    tot_w = sum(it['withdrawal'] for it in clean_items)
    tot_d = sum(it['deposit'] for it in clean_items)
    close_b = clean_items[-1]['balance'] if clean_items else 0.0
    open_b = clean_items[0]['balance'] if clean_items else 0.0

    stmt = BankStatement(
        bank_account_id=bank_acc.id,
        statement_title=statement_title,
        filename=filename_summary,
        from_date=from_d,
        to_date=to_d,
        opening_balance=open_b,
        closing_balance=close_b,
        total_withdrawals=tot_w,
        total_deposits=tot_d,
        notes=statement_notes
    )
    db.session.add(stmt)
    db.session.flush()

    created_count = 0
    for it in clean_items:
        d_obj = it['date']
        amount = it['deposit'] if it['deposit'] > 0 else it['withdrawal']
        matched_v = None
        if amount > 0:
            is_dep = it['deposit'] > 0
            matched_line = TransactionLine.query.join(JournalEntry).filter(
                TransactionLine.account_id == bank_acc.id,
                TransactionLine.debit == (amount if is_dep else 0.0),
                TransactionLine.credit == (0.0 if is_dep else amount),
                JournalEntry.date >= d_obj - timedelta(days=7),
                JournalEntry.date <= d_obj + timedelta(days=7)
            ).first()
            if matched_line:
                matched_v = matched_line.journal_entry

        b_line = BankStatementLine(
            statement_id=stmt.id,
            bank_account_id=bank_acc.id,
            txn_date=d_obj,
            value_date=d_obj,
            particulars=it['description'][:490],
            cheque_ref=it['cheque_ref'][:95],
            withdrawal_dr=it['withdrawal'],
            deposit_cr=it['deposit'],
            balance=it['balance'],
            notes=it['notes'][:490],
            is_reconciled=(matched_v is not None),
            reconciled_date=d_obj if matched_v else None,
            matched_voucher_id=matched_v.id if matched_v else None,
            matched_voucher_no=matched_v.voucher_no if matched_v else None
        )
        db.session.add(b_line)
        created_count += 1

    db.session.commit()
    flash(
        f'Verified and posted {created_count} clean transactions into {bank_acc.name}. Zero duplicate posting ensured.',
        'success'
    )
    return redirect(url_for('bank_statement_report', account_id=bank_account_id))


@app.route('/banking/statements/delete-data', methods=['POST'])
def delete_bank_statement_data():
    """
    Deletes bank statement report data safely.
    CRITICAL: Does NOT touch, delete, or alter any Accounts General Ledger vouchers or balances.
    """
    bank_account_id = request.form.get('bank_account_id', type=int)
    scope = request.form.get('scope', 'all')  # 'all' or 'statement_batch'
    statement_id = request.form.get('statement_id', type=int)

    del_lines = 0
    del_stmts = 0

    if scope == 'statement_batch' and statement_id:
        del_lines = BankStatementLine.query.filter_by(statement_id=statement_id).delete()
        del_stmts = BankStatement.query.filter_by(id=statement_id).delete()
    else:  # 'all' for this bank account
        if bank_account_id:
            del_lines = BankStatementLine.query.filter_by(bank_account_id=bank_account_id).delete()
            del_stmts = BankStatement.query.filter_by(bank_account_id=bank_account_id).delete()
        else:
            del_lines = BankStatementLine.query.delete()
            del_stmts = BankStatement.query.delete()

    db.session.commit()
    flash(
        f'Successfully deleted {del_lines} bank statement record(s) and {del_stmts} statement batch(es). '
        'IMPORTANT: Your Accounts General Ledger vouchers, journal entries, and ledger account balances are completely separate and remain 100% untouched.',
        'success'
    )
    return redirect(url_for('bank_statement_report', account_id=bank_account_id))


@app.route('/banking/statements/line/<int:line_id>/delete', methods=['POST'])
def delete_bank_statement_line(line_id):
    """
    Deletes an individual bank statement row. Does NOT affect the General Ledger.
    """
    line = BankStatementLine.query.get_or_404(line_id)
    acc_id = line.bank_account_id
    db.session.delete(line)
    db.session.commit()
    flash('Bank statement line deleted. Accounts General Ledger remains untouched.', 'success')
    return redirect(url_for('bank_statement_report', account_id=acc_id))


@app.route('/banking/statements/line/<int:line_id>/edit', methods=['POST'])
def edit_bank_statement_line(line_id):
    """
    Edits any and all fields of an individual bank statement transaction:
    date, value date, particulars, cheque_ref, debit, credit, balance, notes,
    reconciliation status, and matched voucher number.
    """
    line = BankStatementLine.query.get_or_404(line_id)

    date_str = request.form.get('txn_date')
    val_date_str = request.form.get('value_date')

    if date_str:
        try:
            line.txn_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except Exception:
            pass
    if val_date_str:
        try:
            line.value_date = datetime.strptime(val_date_str, '%Y-%m-%d').date()
        except Exception:
            pass

    line.particulars = request.form.get('particulars', line.particulars).strip()
    line.cheque_ref = request.form.get('cheque_ref', line.cheque_ref or '').strip()
    line.withdrawal_dr = float(request.form.get('withdrawal_dr', '0').replace(',', '') or 0.0)
    line.deposit_cr = float(request.form.get('deposit_cr', '0').replace(',', '') or 0.0)
    line.balance = float(request.form.get('balance', '0').replace(',', '') or 0.0)
    line.notes = request.form.get('notes', line.notes or '').strip()

    # Reconciled status
    is_recon = request.form.get('is_reconciled')
    if is_recon is not None:
        line.is_reconciled = (is_recon in ('1', 'true', 'on', 'yes'))
        if line.is_reconciled and not line.reconciled_date:
            line.reconciled_date = date.today()
        elif not line.is_reconciled:
            line.reconciled_date = None

    v_no = request.form.get('matched_voucher_no', '').strip()
    line.matched_voucher_no = v_no if v_no else None

    db.session.commit()
    flash(f'Bank statement transaction #{line.id} updated successfully.', 'success')
    return redirect(url_for('bank_statement_report', account_id=line.bank_account_id))


@app.route('/banking/statements/bulk-update', methods=['POST'])
def bulk_update_bank_statement_lines():
    """
    Saves multiple inline edits submitted directly from the spreadsheet table mode.
    """
    bank_account_id = request.form.get('bank_account_id', type=int)
    line_ids = request.form.getlist('line_ids')

    updated_count = 0
    for lid_str in line_ids:
        try:
            lid = int(lid_str)
        except (ValueError, TypeError):
            continue
        line = BankStatementLine.query.get(lid)
        if not line:
            continue

        d_str = request.form.get(f'date_{lid}')
        if d_str:
            d_obj = parse_accounting_date(d_str)
            if d_obj:
                line.txn_date = d_obj

        line.particulars = request.form.get(f'particulars_{lid}', line.particulars).strip()
        line.cheque_ref = request.form.get(f'ref_{lid}', line.cheque_ref or '').strip()
        line.withdrawal_dr = float(request.form.get(f'withdrawal_{lid}', '0').replace(',', '') or 0.0)
        line.deposit_cr = float(request.form.get(f'deposit_{lid}', '0').replace(',', '') or 0.0)
        line.balance = float(request.form.get(f'balance_{lid}', '0').replace(',', '') or 0.0)
        line.notes = request.form.get(f'notes_{lid}', line.notes or '').strip()
        updated_count += 1

    db.session.commit()
    flash(f'Successfully updated {updated_count} bank statement transaction(s) directly.', 'success')
    return redirect(url_for('bank_statement_report', account_id=bank_account_id))


@app.route('/banking/statements/<int:stmt_id>/edit-batch', methods=['POST'])
@app.route('/banking/statements/<int:stmt_id>/update-notes', methods=['POST'])
def edit_bank_statement_batch(stmt_id):
    """
    Updates the batch title, dates, opening/closing balance, and notes for an imported statement batch.
    """
    stmt = BankStatement.query.get_or_404(stmt_id)
    title = request.form.get('statement_title', '').strip()
    from_date_str = request.form.get('from_date', '').strip()
    to_date_str = request.form.get('to_date', '').strip()
    open_bal_str = request.form.get('opening_balance')
    close_bal_str = request.form.get('closing_balance')
    notes = request.form.get('notes', '').strip()

    if title:
        stmt.statement_title = title
    if from_date_str:
        f_d = parse_accounting_date(from_date_str)
        if f_d:
            stmt.from_date = f_d
    if to_date_str:
        t_d = parse_accounting_date(to_date_str)
        if t_d:
            stmt.to_date = t_d
    if open_bal_str is not None:
        try:
            stmt.opening_balance = float(open_bal_str.replace(',', '') or 0.0)
        except Exception:
            pass
    if close_bal_str is not None:
        try:
            stmt.closing_balance = float(close_bal_str.replace(',', '') or 0.0)
        except Exception:
            pass
    stmt.notes = notes
    db.session.commit()
    flash(f'Statement batch "{stmt.statement_title}" updated successfully.', 'success')
    return redirect(url_for('bank_statement_report', account_id=stmt.bank_account_id))


@app.route('/banking/statements/add-line', methods=['POST'])
def add_bank_statement_line():
    bank_account_id = request.form.get('bank_account_id', type=int)
    txn_date_str = request.form.get('txn_date')
    particulars = request.form.get('particulars')
    cheque_ref = request.form.get('cheque_ref', '')
    txn_type = request.form.get('txn_type', 'Deposit')
    amount = float(request.form.get('amount') or 0.0)
    balance = float(request.form.get('balance') or 0.0)
    notes = request.form.get('notes', '').strip()

    try:
        t_d = datetime.strptime(txn_date_str, '%Y-%m-%d').date()
    except Exception:
        t_d = date.today()

    w_dr = amount if txn_type == 'Withdrawal' else 0.0
    d_cr = amount if txn_type == 'Deposit' else 0.0

    line = BankStatementLine(
        bank_account_id=bank_account_id,
        txn_date=t_d,
        value_date=t_d,
        particulars=particulars,
        cheque_ref=cheque_ref,
        withdrawal_dr=w_dr,
        deposit_cr=d_cr,
        balance=balance,
        notes=notes,
        is_reconciled=False
    )
    db.session.add(line)
    db.session.commit()
    flash('Bank statement entry added successfully.', 'success')
    return redirect(url_for('bank_statement_report', account_id=bank_account_id))


@app.route('/reports/bank-statement/export-excel')
def export_bank_statement_excel():
    bank_acc_id = request.args.get('account_id', type=int)
    bank_acc = Account.query.get(bank_acc_id) if bank_acc_id else None

    query = BankStatementLine.query
    if bank_acc_id:
        query = query.filter_by(bank_account_id=bank_acc_id)

    lines = query.order_by(BankStatementLine.txn_date.asc(), BankStatementLine.id.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bank Statement"

    ws['A1'] = f"BANK STATEMENT REPORT - {bank_acc.name if bank_acc else 'ALL ACCOUNTS'}"
    ws['A1'].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    ws['A2'] = f"Generated on {date.today().strftime('%d-%b-%Y')} | Currency: BDT (৳)"
    ws['A2'].font = Font(name="Calibri", size=10, italic=True, color="595959")

    headers = ["Date", "Value Date", "Particulars / Description", "Cheque / Ref No", "Withdrawals (Debit)", "Deposits (Credit)", "Running Balance", "Recon Status", "Matched Voucher"]
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center" if col_idx not in (3, 4) else "left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    r_idx = 5
    for l in lines:
        ws.cell(row=r_idx, column=1, value=l.txn_date.strftime('%d-%b-%Y') if l.txn_date else '').alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2, value=l.value_date.strftime('%d-%b-%Y') if l.value_date else '').alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=3, value=l.particulars)
        ws.cell(row=r_idx, column=4, value=l.cheque_ref or '-')
        
        c_w = ws.cell(row=r_idx, column=5, value=l.withdrawal_dr)
        c_w.number_format = '#,##0.00'
        c_d = ws.cell(row=r_idx, column=6, value=l.deposit_cr)
        c_d.number_format = '#,##0.00'
        c_b = ws.cell(row=r_idx, column=7, value=l.balance)
        c_b.number_format = '#,##0.00'

        recon_str = "Reconciled" if l.is_reconciled else "Unreconciled"
        c_st = ws.cell(row=r_idx, column=8, value=recon_str)
        c_st.alignment = Alignment(horizontal="center")
        if l.is_reconciled:
            c_st.font = Font(color="008000", bold=True)
        else:
            c_st.font = Font(color="C00000")

        ws.cell(row=r_idx, column=9, value=l.matched_voucher_no or '-')

        for c_i in range(1, 10):
            ws.cell(row=r_idx, column=c_i).border = thin_border

        r_idx += 1

    # Totals Row
    ws.cell(row=r_idx, column=3, value="TOTALS:").font = Font(bold=True)
    c_tot_w = ws.cell(row=r_idx, column=5, value=sum(l.withdrawal_dr for l in lines))
    c_tot_w.font = Font(bold=True)
    c_tot_w.number_format = '#,##0.00'
    c_tot_d = ws.cell(row=r_idx, column=6, value=sum(l.deposit_cr for l in lines))
    c_tot_d.font = Font(bold=True)
    c_tot_d.number_format = '#,##0.00'
    tot_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for c_i in range(1, 10):
        ws.cell(row=r_idx, column=c_i).fill = tot_fill
        ws.cell(row=r_idx, column=c_i).border = thin_border

    # Adjust widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"Bank_Statement_{bank_acc.code if bank_acc else 'All'}_{date.today().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==============================================================================
# BANK RECONCILIATION MODULE (MONTH-END RECONCILIATION)
# ==============================================================================

@app.route('/banking/reconciliation')
def bank_reconciliation():
    bank_accounts = Account.query.filter_by(is_cash_or_bank=True).all()
    selected_acc_id = request.args.get('account_id', type=int)
    if not selected_acc_id and bank_accounts:
        premier = next((a for a in bank_accounts if 'Premier' in a.name), bank_accounts[0])
        selected_acc_id = premier.id

    selected_acc = Account.query.get(selected_acc_id) if selected_acc_id else None

    # Month selection
    recon_month = request.args.get('month', '2026-01')
    try:
        y, m = map(int, recon_month.split('-'))
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        as_of_date = date(y, m, last_day)
    except Exception:
        as_of_date = date.today()
        recon_month = as_of_date.strftime('%Y-%m')

    # 1. Statement Ending Balance
    latest_stmt_line = BankStatementLine.query.filter(
        BankStatementLine.bank_account_id == selected_acc_id,
        BankStatementLine.txn_date <= as_of_date
    ).order_by(BankStatementLine.txn_date.desc(), BankStatementLine.id.desc()).first()

    stmt_balance = latest_stmt_line.balance if latest_stmt_line else 0.0

    # 2. General Ledger Book Balance as of selected date
    prior_dr = db.session.query(func.coalesce(func.sum(TransactionLine.debit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == selected_acc_id, JournalEntry.date <= as_of_date).scalar()
    prior_cr = db.session.query(func.coalesce(func.sum(TransactionLine.credit), 0.0))\
        .join(JournalEntry)\
        .filter(TransactionLine.account_id == selected_acc_id, JournalEntry.date <= as_of_date).scalar()

    book_balance = (selected_acc.opening_balance if selected_acc else 0.0) + prior_dr - prior_cr

    # 3. GL Transactions for this bank account in month
    m_start = date(y, m, 1)
    gl_lines = TransactionLine.query.join(JournalEntry).filter(
        TransactionLine.account_id == selected_acc_id,
        JournalEntry.date >= m_start,
        JournalEntry.date <= as_of_date
    ).order_by(JournalEntry.date.asc(), JournalEntry.id.asc()).all()

    # 4. Bank Statement Transactions in month
    stmt_lines = BankStatementLine.query.filter(
        BankStatementLine.bank_account_id == selected_acc_id,
        BankStatementLine.txn_date >= m_start,
        BankStatementLine.txn_date <= as_of_date
    ).order_by(BankStatementLine.txn_date.asc(), BankStatementLine.id.asc()).all()

    # Reconciling items:
    # - Uncredited deposits (GL receipts Dr to Bank, not cleared in statement)
    uncredited_deposits = sum(l.debit for l in gl_lines if l.debit > 0 and not any(sl.matched_voucher_id == l.journal_entry_id for sl in stmt_lines))
    # - Outstanding cheques (GL payments Cr to Bank, not presented/cleared in statement)
    unpresented_cheques = sum(l.credit for l in gl_lines if l.credit > 0 and not any(sl.matched_voucher_id == l.journal_entry_id for sl in stmt_lines))
    # - Direct bank credits not in books (Statement credits not matched)
    direct_credits = sum(sl.deposit_cr for sl in stmt_lines if not sl.is_reconciled and sl.deposit_cr > 0)
    # - Bank charges / debits not in books (Statement debits not matched)
    bank_charges = sum(sl.withdrawal_dr for sl in stmt_lines if not sl.is_reconciled and sl.withdrawal_dr > 0)

    adjusted_bank = round(stmt_balance + uncredited_deposits - unpresented_cheques, 2)
    adjusted_book = round(book_balance + direct_credits - bank_charges, 2)
    variance = round(adjusted_bank - adjusted_book, 2)
    is_reconciled = (abs(variance) < 0.01)

    # Saved month-end reconciliation records
    past_recons = BankReconciliation.query.filter_by(bank_account_id=selected_acc_id).order_by(BankReconciliation.statement_date.desc()).all()

    return render_template(
        'banking/bank_reconciliation.html',
        bank_accounts=bank_accounts,
        selected_account=selected_acc,
        recon_month=recon_month,
        as_of_date=as_of_date,
        stmt_balance=stmt_balance,
        book_balance=book_balance,
        uncredited_deposits=uncredited_deposits,
        unpresented_cheques=unpresented_cheques,
        direct_credits=direct_credits,
        bank_charges=bank_charges,
        adjusted_bank=adjusted_bank,
        adjusted_book=adjusted_book,
        variance=variance,
        is_reconciled=is_reconciled,
        gl_lines=gl_lines,
        stmt_lines=stmt_lines,
        past_recons=past_recons
    )


@app.route('/banking/reconciliation/auto-match', methods=['POST'])
def bank_reconciliation_auto_match():
    bank_acc_id = request.form.get('account_id', type=int)
    recon_month = request.form.get('month', '2026-01')

    try:
        y, m = map(int, recon_month.split('-'))
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        m_start = date(y, m, 1)
        m_end = date(y, m, last_day)
    except Exception:
        m_start = date.today().replace(day=1)
        m_end = date.today()

    gl_lines = TransactionLine.query.join(JournalEntry).filter(
        TransactionLine.account_id == bank_acc_id,
        JournalEntry.date >= m_start,
        JournalEntry.date <= m_end
    ).all()

    stmt_lines = BankStatementLine.query.filter(
        BankStatementLine.bank_account_id == bank_acc_id,
        BankStatementLine.txn_date >= m_start,
        BankStatementLine.txn_date <= m_end
    ).all()

    match_count = 0
    for sl in stmt_lines:
        if sl.is_reconciled:
            continue
        amt = sl.deposit_cr if sl.deposit_cr > 0 else sl.withdrawal_dr
        is_dep = sl.deposit_cr > 0

        # Find matching GL entry with same amount
        for gl in gl_lines:
            gl_amt = gl.debit if is_dep else gl.credit
            if abs(gl_amt - amt) < 0.01:
                sl.is_reconciled = True
                sl.reconciled_date = sl.txn_date
                sl.matched_voucher_id = gl.journal_entry_id
                sl.matched_voucher_no = gl.journal_entry.voucher_no
                match_count += 1
                break

    db.session.commit()
    flash(f'Auto-match complete! {match_count} bank and general ledger entries matched successfully.', 'success')
    return redirect(url_for('bank_reconciliation', account_id=bank_acc_id, month=recon_month))


@app.route('/banking/reconciliation/save', methods=['POST'])
def save_bank_reconciliation():
    bank_acc_id = request.form.get('account_id', type=int)
    recon_month = request.form.get('month')
    statement_date_str = request.form.get('statement_date')
    stmt_balance = float(request.form.get('stmt_balance') or 0.0)
    book_balance = float(request.form.get('book_balance') or 0.0)
    uncredited = float(request.form.get('uncredited_deposits') or 0.0)
    unpresented = float(request.form.get('unpresented_cheques') or 0.0)
    bank_chg = float(request.form.get('bank_charges') or 0.0)
    direct_cr = float(request.form.get('direct_credits') or 0.0)
    adj_bank = float(request.form.get('adjusted_bank') or 0.0)
    adj_book = float(request.form.get('adjusted_book') or 0.0)
    variance = float(request.form.get('variance') or 0.0)
    notes = request.form.get('notes', '')

    stmt_date = parse_accounting_date(statement_date_str) or date.today()

    existing = BankReconciliation.query.filter_by(bank_account_id=bank_acc_id, recon_month=recon_month).first()
    if existing:
        recon = existing
    else:
        recon = BankReconciliation(bank_account_id=bank_acc_id, recon_month=recon_month, statement_date=stmt_date)
        db.session.add(recon)

    recon.statement_date = stmt_date
    recon.statement_balance = stmt_balance
    recon.book_balance = book_balance
    recon.uncredited_deposits = uncredited
    recon.unpresented_cheques = unpresented
    recon.bank_charges = bank_chg
    recon.direct_credits = direct_cr
    recon.adjusted_bank_balance = adj_bank
    recon.adjusted_book_balance = adj_book
    recon.variance = variance
    recon.status = "Reconciled" if abs(variance) < 0.01 else "Draft"
    recon.notes = notes
    recon.reconciled_at = datetime.utcnow()

    db.session.commit()
    flash(f'Bank Reconciliation for {recon_month} saved and locked successfully.', 'success')
    return redirect(url_for('bank_reconciliation', account_id=bank_acc_id, month=recon_month))


@app.route('/banking/reconciliation/report')
def bank_reconciliation_report():
    bank_acc_id = request.args.get('account_id', type=int)
    bank_accounts = Account.query.filter_by(is_cash_or_bank=True).all()
    if not bank_acc_id and bank_accounts:
        premier = next((a for a in bank_accounts if 'Premier' in a.name), bank_accounts[0])
        bank_acc_id = premier.id

    bank_acc = Account.query.get(bank_acc_id) if bank_acc_id else None
    if not bank_acc:
        flash("Please select a bank account to view reconciliation report.", "warning")
        return redirect(url_for('bank_reconciliation'))

    recon_month = request.args.get('month', '2026-01')
    recon = BankReconciliation.query.filter_by(bank_account_id=bank_acc.id, recon_month=recon_month).first()

    return render_template(
        'banking/bank_reconciliation_report.html',
        bank_account=bank_acc,
        recon=recon,
        month=recon_month
    )


# ==============================================================================
# EMPLOYEE DATABASE & MASTER IMPORT MODULE
# ==============================================================================

@app.route('/employees')
def employee_list():
    department = request.args.get('department', 'all')
    status = request.args.get('status', 'Active')
    search = request.args.get('search', '').strip()

    query = Employee.query
    if department != 'all':
        query = query.filter_by(department=department)
    if status != 'all':
        query = query.filter_by(status=status)
    if search:
        s_pat = f"%{search}%"
        query = query.filter(
            db.or_(
                Employee.employee_id.ilike(s_pat),
                Employee.name.ilike(s_pat),
                Employee.designation.ilike(s_pat),
                Employee.department.ilike(s_pat),
                Employee.nid_passport.ilike(s_pat)
            )
        )

    employees = query.order_by(Employee.employee_id.asc()).all()
    departments = db.session.query(Employee.department).distinct().all()
    dept_list = [d[0] for d in departments if d[0]]

    total_count = len(employees)
    active_count = sum(1 for e in employees if e.status == 'Active')
    total_payroll = sum(e.basic_salary + e.allowances for e in employees if e.status == 'Active')

    return render_template(
        'employees/employee_list.html',
        employees=employees,
        departments=dept_list,
        selected_dept=department,
        selected_status=status,
        search_query=search,
        total_count=total_count,
        active_count=active_count,
        total_payroll=total_payroll
    )


@app.route('/employees/<int:emp_id>/profile')
def employee_profile(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    salary_histories = EmployeeSalaryHistory.query.filter_by(employee_id=emp_id).order_by(EmployeeSalaryHistory.salary_month.desc()).all()
    fund_contributions = EmployeeFundContribution.query.filter_by(employee_id=emp_id).order_by(EmployeeFundContribution.contribution_month.desc()).all()
    total_fund = fund_contributions[0].cumulative_balance if fund_contributions else 0.0

    return render_template(
        'employees/employee_profile.html',
        employee=employee,
        salary_histories=salary_histories,
        fund_contributions=fund_contributions,
        total_fund=total_fund
    )


@app.route('/employees/import-excel', methods=['POST'])
def import_employees_excel():
    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Please select an Excel file to import employees.', 'warning')
        return redirect(url_for('employee_list'))

    try:
        wb = openpyxl.load_workbook(file.stream, data_only=True)
        ws = wb.active
        imported = 0
        updated = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue

            emp_code = str(row[0]).strip() if row[0] else ""
            if not emp_code:
                continue

            name = str(row[1]).strip() if len(row) > 1 and row[1] else "Staff Member"
            desig = str(row[2]).strip() if len(row) > 2 and row[2] else "Employee"
            dept = str(row[3]).strip() if len(row) > 3 and row[3] else "Administration"
            basic = clean_accounting_num(row[4]) if len(row) > 4 else 0.0
            allowances = clean_accounting_num(row[5]) if len(row) > 5 else 0.0
            deductions = clean_accounting_num(row[6]) if len(row) > 6 else 0.0
            bank_acc = str(row[7]).strip() if len(row) > 7 and row[7] else ""
            j_date = parse_accounting_date(row[8]) if len(row) > 8 and row[8] else None
            nid = str(row[9]).strip() if len(row) > 9 and row[9] else ""
            phone = str(row[10]).strip() if len(row) > 10 and row[10] else ""
            email = str(row[11]).strip() if len(row) > 11 and row[11] else ""
            addr = str(row[12]).strip() if len(row) > 12 and row[12] else ""

            emp = Employee.query.filter_by(employee_id=emp_code).first()
            if emp:
                emp.name = name
                emp.designation = desig
                emp.department = dept
                emp.basic_salary = basic
                emp.allowances = allowances
                emp.deductions = deductions
                emp.bank_account_no = bank_acc
                if j_date is not None:
                    emp.joining_date = j_date
                if nid:
                    emp.nid_passport = nid
                if phone:
                    emp.phone = phone
                if email:
                    emp.email = email
                if addr:
                    emp.present_address = addr
                updated += 1
            else:
                emp = Employee(
                    employee_id=emp_code,
                    name=name,
                    designation=desig,
                    department=dept,
                    basic_salary=basic,
                    allowances=allowances,
                    deductions=deductions,
                    bank_account_no=bank_acc,
                    joining_date=j_date,
                    nid_passport=nid,
                    phone=phone,
                    email=email,
                    present_address=addr,
                    status='Active'
                )
                db.session.add(emp)
                imported += 1

        db.session.commit()
        flash(f'Employee import completed: {imported} new profiles added, {updated} updated.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Error importing employee Excel file: {str(e)}', 'danger')

    return redirect(url_for('employee_list'))


@app.route('/employees/sample-excel')
def download_employee_sample_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employee Master"

    headers = ["Emp ID", "Full Name", "Designation", "Department", "Basic Salary", "Allowances", "Deductions", "Bank Account No", "Joining Date", "NID / Passport", "Phone", "Email", "Address"]
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sample_rows = [
        ["EMP-016", "Mohammad Tanvir Hasan", "Assistant Accounts Officer", "Finance & Accounts", 35000.0, 15000.0, 3000.0, "", "", "", "", "", ""],
        ["EMP-017", "Nusrat Jahan", "HR Coordinator", "Human Resources", 40000.0, 18000.0, 3500.0, "", "", "", "", "", ""]
    ]

    for r_idx, row_vals in enumerate(sample_rows, 2):
        for c_idx, val in enumerate(row_vals, 1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Employee_Master_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/employees/delete-all', methods=['POST'])
@app.route('/payroll/employees/delete-all', methods=['POST'])
def delete_all_employees():
    """
    Permanently clears all records from the Corporate Employee Database.
    Strict isolation constraint: ZERO impact on General Ledger, Accounts, Vouchers,
    Cash Book, Bank Statements, Reconciliations, or any financial accounting reports.
    """
    confirm = request.form.get('confirm_delete')
    if confirm != 'yes':
        flash('Operation cancelled: Please check the confirmation box to authorize deleting the employee database.', 'warning')
        return redirect(request.referrer or url_for('employee_list'))

    clear_associated = request.form.get('clear_associated', 'yes') == 'yes'

    try:
        emp_count = Employee.query.count()
        if emp_count == 0:
            flash('Employee database is already empty.', 'info')
            return redirect(request.referrer or url_for('employee_list'))

        deleted_funds = 0
        deleted_salaries = 0
        deleted_payroll_items = 0
        deleted_payroll_records = 0

        # Child tables referencing Employee:
        # Note: We ensure complete isolation from accounting data:
        # JournalEntry, TransactionLine, CashBookRow, Account, BankStatement are NEVER touched.
        if clear_associated:
            deleted_funds = EmployeeFundContribution.query.delete()
            deleted_salaries = EmployeeSalaryHistory.query.delete()
            deleted_payroll_items = PayrollItem.query.delete()
            # Only deletes the payroll record header; any posted GL payment voucher remains intact in JournalEntry
            deleted_payroll_records = PayrollRecord.query.delete()
        else:
            deleted_payroll_items = PayrollItem.query.delete()
            deleted_funds = EmployeeFundContribution.query.delete()
            deleted_salaries = EmployeeSalaryHistory.query.delete()

        # Delete all employee records
        deleted_emps = Employee.query.delete()
        db.session.commit()

        msg = f"Successfully purged all {deleted_emps} employee profiles from the Corporate Employee Database."
        if deleted_funds > 0 or deleted_salaries > 0:
            msg += f" (Associated records cleared: {deleted_funds} fund entries, {deleted_salaries} salary logs)."
        msg += " All Accounting Reports (Trial Balance, Cash Book, Balance Sheet, Vouchers) remain 100% unaffected."

        flash(msg, 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Error deleting employee database: {str(e)}', 'danger')

    return redirect(request.referrer or url_for('employee_list'))


# ==============================================================================
# EMPLOYEE CONTRIBUTION FUND MODULE (PF / GRATUITY / WELFARE)
# ==============================================================================

@app.route('/payroll/fund-contributions')
def fund_contributions_dashboard():
    employees = Employee.query.order_by(Employee.employee_id.asc()).all()

    fund_summaries = []
    total_fund_accumulated = 0.0
    total_emp_share = 0.0
    total_comp_share = 0.0
    total_welfare_share = 0.0
    total_profit_accrued = 0.0

    for emp in employees:
        records = EmployeeFundContribution.query.filter_by(employee_id=emp.id).order_by(EmployeeFundContribution.contribution_date.desc(), EmployeeFundContribution.id.desc()).all()
        emp_tot = sum(r.employee_share or 0.0 for r in records)
        comp_tot = sum(r.employer_share or 0.0 for r in records)
        welfare_tot = sum(r.welfare_fund or 0.0 for r in records)
        prof_tot = sum(r.profit_interest or 0.0 for r in records)
        cumul = records[0].cumulative_balance if records else 0.0

        fund_summaries.append({
            'employee': emp,
            'emp_share': emp_tot,
            'comp_share': comp_tot,
            'welfare_share': welfare_tot,
            'profit': prof_tot,
            'balance': cumul,
            'record_count': len(records)
        })

        total_emp_share += emp_tot
        total_comp_share += comp_tot
        total_welfare_share += welfare_tot
        total_profit_accrued += prof_tot
        total_fund_accumulated += cumul

    return render_template(
        'payroll/fund_contributions.html',
        fund_summaries=fund_summaries,
        total_fund_accumulated=total_fund_accumulated,
        total_emp_share=total_emp_share,
        total_comp_share=total_comp_share,
        total_welfare_share=total_welfare_share,
        total_profit_accrued=total_profit_accrued
    )


@app.route('/payroll/fund-contributions/employee/<int:emp_id>')
def employee_fund_ledger(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    records = EmployeeFundContribution.query.filter_by(employee_id=emp_id).order_by(EmployeeFundContribution.contribution_date.asc(), EmployeeFundContribution.id.asc()).all()

    tot_emp = sum(r.employee_share or 0.0 for r in records)
    tot_comp = sum(r.employer_share or 0.0 for r in records)
    tot_welfare = sum(r.welfare_fund or 0.0 for r in records)
    tot_prof = sum(r.profit_interest or 0.0 for r in records)
    closing_bal = records[-1].cumulative_balance if records else 0.0

    return render_template(
        'payroll/employee_fund_ledger.html',
        employee=employee,
        records=records,
        tot_emp=tot_emp,
        tot_comp=tot_comp,
        tot_welfare=tot_welfare,
        tot_prof=tot_prof,
        closing_bal=closing_bal
    )


@app.route('/payroll/fund-contributions/employee/<int:emp_id>/export-excel')
def export_employee_fund_excel(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    records = EmployeeFundContribution.query.filter_by(employee_id=emp_id).order_by(EmployeeFundContribution.contribution_date.asc(), EmployeeFundContribution.id.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contribution Fund Ledger"

    ws['A1'] = f"EMPLOYEE CONTRIBUTION FUND LEDGER - {employee.name.upper()} ({employee.employee_id})"
    ws['A1'].font = Font(name="Calibri", size=13, bold=True, color="1F4E79")
    ws['A2'] = f"Designation: {employee.designation} | Department: {employee.department} | Date: {date.today().strftime('%d-%b-%Y')}"
    ws['A2'].font = Font(name="Calibri", size=10, italic=True, color="595959")

    headers = ["Month", "Date", "PF Employee Share (৳)", "PF Employer Share (৳)", "Welfare Fund (৳)", "Profit / Interest (৳)", "Monthly Total (৳)", "Cumulative Fund (৳)", "Remarks"]
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for c_i, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=c_i, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    r_idx = 5
    for r in records:
        ws.cell(row=r_idx, column=1, value=r.contribution_month).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2, value=r.contribution_date.strftime('%d-%b-%Y') if r.contribution_date else '').alignment = Alignment(horizontal="center")

        for col_i, val in [(3, r.employee_share), (4, r.employer_share), (5, r.welfare_fund), (6, r.profit_interest), (7, r.total_monthly), (8, r.cumulative_balance)]:
            c = ws.cell(row=r_idx, column=col_i, value=val)
            c.number_format = '#,##0.00'

        ws.cell(row=r_idx, column=9, value=r.remarks or '')

        for c_i in range(1, 10):
            ws.cell(row=r_idx, column=c_i).border = thin_border
        r_idx += 1

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"Fund_Statement_{employee.employee_id}_{date.today().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


MONTH_NAME_LOOKUP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12
}

def clean_fund_num(val):
    """Safely parses financial numbers from diverse Excel representations."""
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '').replace('৳', '').replace('$', '').replace('BDT', '').replace('Tk', '').strip()
    if s == '-' or not s:
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def parse_fund_excel_period(raw_month, raw_date=None, sheet_year=None):
    """
    Parses contribution month (YYYY-MM) and date (date) from diverse Excel formats,
    specifically supporting multi-year data from 2018 to 2026.
    """
    if isinstance(raw_date, (datetime, date)):
        y = raw_date.year
        m = raw_date.month
        c_date = raw_date if isinstance(raw_date, date) and not isinstance(raw_date, datetime) else raw_date.date()
        return f"{y:04d}-{m:02d}", c_date

    if isinstance(raw_month, (datetime, date)):
        y = raw_month.year
        m = raw_month.month
        c_date = raw_month if isinstance(raw_month, date) and not isinstance(raw_month, datetime) else raw_month.date()
        return f"{y:04d}-{m:02d}", c_date

    s_m = str(raw_month or '').strip()
    s_d = str(raw_date or '').strip()

    # Try matching YYYY-MM or YYYY/MM with robust month regex
    m1 = re.search(r'(20\d\d)[-_/.](1[0-2]|0[1-9]|[1-9])(?!\d)', s_m)
    if m1:
        y = int(m1.group(1))
        m = int(m1.group(2))
        return f"{y:04d}-{m:02d}", date(y, m, 28)

    # Try matching MM-YYYY or MM/YYYY
    m2 = re.search(r'(1[0-2]|0[1-9]|[1-9])[-_/.](20\d\d)', s_m)
    if m2:
        m = int(m2.group(1))
        y = int(m2.group(2))
        return f"{y:04d}-{m:02d}", date(y, m, 28)

    # Try matching Month Name + Year (e.g. Jan-18, January 2018, Mar 2026)
    m3 = re.search(r'([a-zA-Z]+)[-_/\s]*(20\d\d|\d{2})', s_m)
    if m3:
        m_name = m3.group(1).lower()
        yr_str = m3.group(2)
        y = int(yr_str) if len(yr_str) == 4 else 2000 + int(yr_str)
        if m_name in MONTH_NAME_LOOKUP:
            m = MONTH_NAME_LOOKUP[m_name]
            return f"{y:04d}-{m:02d}", date(y, m, 28)

    # Try parsing raw date column if available
    if s_d:
        d_parsed = parse_accounting_date(s_d)
        if d_parsed:
            d_obj = d_parsed.date() if isinstance(d_parsed, datetime) else d_parsed
            return f"{d_obj.year:04d}-{d_obj.month:02d}", d_obj

    # If sheet_year exists (e.g. sheet named "2018" or "2024")
    if sheet_year and re.match(r'^20\d\d$', str(sheet_year)):
        y = int(sheet_year)
        s_clean = s_m.lower()
        for k, v in MONTH_NAME_LOOKUP.items():
            if k == s_clean or k in s_clean:
                return f"{y:04d}-{v:02d}", date(y, v, 28)
        if s_m.isdigit() and 1 <= int(s_m) <= 12:
            m = int(s_m)
            return f"{y:04d}-{m:02d}", date(y, m, 28)

    return None, None


@app.route('/payroll/fund-contributions/import-excel', methods=['POST'])
def import_fund_contributions_excel():
    """
    Imports Employee Contribution Fund (Provident & Welfare) entries from an uploaded Excel file.
    Supports:
    1. Horizontal Employee-Wise Matrix layout (e.g. csf.xlsx / csf 18-19.xlsx where rows are employees and columns are months).
    2. Vertical Tabular layout (e.g. standard template where each row is an entry).
    Matches by Employee Name and/or ID/SL number.
    Auto-creates missing employees in database.
    Chronologically imports/updates historical data from 2018 to 2026.
    Recalculates running cumulative fund balances.
    STRICT CONSTRAINT: Isolated from General Accounts (zero GL/Voucher impact).
    """
    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Please select an Excel file with contribution fund entries.', 'warning')
        return redirect(url_for('fund_contributions_dashboard'))

    try:
        wb = openpyxl.load_workbook(file.stream, data_only=True)
        
        # Build employee caches and collision-proof ID generator
        all_db_emps = Employee.query.all()
        used_emp_codes = {e.employee_id.strip().upper() for e in all_db_emps if e.employee_id}
        emp_by_code = {e.employee_id.strip().upper(): e for e in all_db_emps if e.employee_id}

        def normalize_emp_name(name):
            if not name:
                return ''
            s = re.sub(r'[\(\-_]\s*old\s*[\)]?', ' (old)', str(name).strip().lower())
            s = re.sub(r'^(?:0+)?([a-z])', r'\1', s)
            s = re.sub(r'\s*\(\d+\)$', '', s)
            return re.sub(r'\s+', ' ', s).strip()

        emp_by_name = {normalize_emp_name(e.name): e for e in all_db_emps if e.name}

        max_emp_seq = 0
        for code in used_emp_codes:
            nums = re.findall(r'\d+', code)
            if nums:
                max_emp_seq = max(max_emp_seq, int(nums[-1]))

        def get_next_unique_code(preferred_id=None, prefix="EMP-"):
            nonlocal max_emp_seq
            if preferred_id:
                clean = str(preferred_id).strip()
                if clean.isdigit():
                    cand = f"{prefix}{int(clean):03d}"
                    if cand.upper() not in used_emp_codes:
                        used_emp_codes.add(cand.upper())
                        max_emp_seq = max(max_emp_seq, int(clean))
                        return cand
                elif re.match(r'^(?:EMP|E)[-_]?(\d+)$', clean, re.I):
                    n = int(re.search(r'\d+', clean).group(0))
                    cand = f"{prefix}{n:03d}"
                    if cand.upper() not in used_emp_codes:
                        used_emp_codes.add(cand.upper())
                        max_emp_seq = max(max_emp_seq, n)
                        return cand
                elif clean.isalnum() and len(clean) <= 20 and clean.lower() not in ['tis', 'bcs', 'tbn', 'g&g', 'sl', 'id']:
                    cand = f"{prefix}{clean.upper()}"
                    if cand.upper() not in used_emp_codes:
                        used_emp_codes.add(cand.upper())
                        return cand
            while True:
                max_emp_seq += 1
                cand = f"{prefix}{max_emp_seq:03d}"
                if cand.upper() not in used_emp_codes:
                    used_emp_codes.add(cand.upper())
                    return cand

        created_employees_count = 0
        all_parsed_contributions = [] # dict(emp_id, month, date, emp_share, comp_share, welfare, profit, remarks)
        all_years = set()

        # Check if any sheet is Horizontal Employee-Wise Matrix
        is_matrix_layout = False
        matrix_sheet = None
        for ws in wb.worksheets:
            first_rows = []
            for r in range(1, min(10, ws.max_row + 1)):
                first_rows.extend([str(ws.cell(r, c).value or '').lower() for c in range(1, min(30, ws.max_column + 1))])
            combined_text = " ".join(first_rows)
            if 'opening balance' in combined_text or combined_text.count('pf employee') >= 2 or combined_text.count('pf employer') >= 2:
                is_matrix_layout = True
                matrix_sheet = ws
                break

        if is_matrix_layout and matrix_sheet:
            ws = matrix_sheet
            # ------------------------------------------------------------------
            # MODE A: HORIZONTAL EMPLOYEE-WISE MATRIX LAYOUT (csf.xlsx / multi-year)
            # ------------------------------------------------------------------
            section_rows = []
            for r in range(1, ws.max_row + 1):
                for c in range(1, 15):
                    val = str(ws.cell(r, c).value or '').lower()
                    if 'opening balance' in val:
                        section_rows.append(r)
                        break

            if not section_rows:
                section_rows = [2]

            for s_idx, hr in enumerate(section_rows):
                next_hr = section_rows[s_idx + 1] if s_idx + 1 < len(section_rows) else ws.max_row + 1
                sub_hr = hr + 1

                # Detect column mapping for this specific section
                id_c, sl_c, name_c, dept_c, doj_c, basic_c = None, None, None, None, None, None
                for c in range(1, min(15, ws.max_column + 1)):
                    v_hr = str(ws.cell(hr, c).value or '')
                    v_sub = str(ws.cell(sub_hr, c).value or '')
                    h_text = f"{v_hr} {v_sub}".strip().lower()
                    if re.search(r'\b(name|employee name|particulars)\b', h_text) and not name_c:
                        name_c = c
                    elif re.search(r'\b(id|emp id|employee id|emp code)\b', h_text) and not id_c:
                        id_c = c
                    elif re.search(r'\b(sl|sl no|serial)\b', h_text) and not sl_c:
                        sl_c = c
                    elif re.search(r'\b(dept|department|company)\b', h_text) and not dept_c:
                        dept_c = c
                    elif re.search(r'\b(joining|doj|starting|satarting)\b', h_text) and not doj_c:
                        doj_c = c
                    elif re.search(r'\b(basic|salary)\b', h_text) and not basic_c:
                        basic_c = c

                # Detect month columns for this section
                month_columns = []
                for c in range(1, ws.max_column + 1):
                    v_hr = ws.cell(hr, c).value
                    # Opening Balance (initial deposit for 2018-09 in the first section)
                    if s_idx == 0 and v_hr and 'opening balance' in str(v_hr).lower():
                        emp_c, comp_c = None, None
                        for off in range(3):
                            sv = str(ws.cell(sub_hr, c + off).value or '').lower()
                            if 'employee' in sv and not emp_c:
                                emp_c = c + off
                            elif 'employer' in sv and not comp_c:
                                comp_c = c + off
                        if emp_c and comp_c:
                            month_columns.append(('2018-09', date(2018, 9, 30), emp_c, comp_c, 'Opening Balance up to 2018-09'))
                    elif isinstance(v_hr, datetime):
                        m = v_hr.month
                        y = 2000 + v_hr.day if v_hr.day in range(15, 35) else v_hr.year
                        m_str = f"{y:04d}-{m:02d}"
                        emp_c, comp_c = None, None
                        for off in range(3):
                            sv = str(ws.cell(sub_hr, c + off).value or '').lower()
                            if 'employee' in sv and not emp_c:
                                emp_c = c + off
                            elif 'employer' in sv and not comp_c:
                                comp_c = c + off
                        if emp_c and comp_c and not any(x[0] == m_str for x in month_columns):
                            month_columns.append((m_str, date(y, m, 28), emp_c, comp_c, f"{m_str} PF contribution"))
                    elif isinstance(v_hr, str) and v_hr.strip():
                        pm, pd = parse_fund_excel_period(v_hr, None, None)
                        if pm:
                            emp_c, comp_c = None, None
                            for off in range(3):
                                sv = str(ws.cell(sub_hr, c + off).value or '').lower()
                                if 'employee' in sv and not emp_c:
                                    emp_c = c + off
                                elif 'employer' in sv and not comp_c:
                                    comp_c = c + off
                            if emp_c and comp_c and not any(x[0] == pm for x in month_columns):
                                month_columns.append((pm, pd or date(int(pm.split('-')[0]), int(pm.split('-')[1]), 28), emp_c, comp_c, f"{pm} PF contribution"))

                # Process employee rows in this section
                for r in range(sub_hr + 1, next_hr):
                    raw_name = str(ws.cell(r, name_c).value or '').strip() if name_c else ''
                    if not raw_name:
                        continue
                    if any(k in raw_name.lower() for k in ['total', 'grand', 'subtotal', 'average', 'particulars', 'signature', 'page', 'opening balance']):
                        continue
                    if raw_name.lower() in ['tis', 'bcs', 'tbn', 'g&g', 'department', 'name', 'id', 'sl no', 'serial']:
                        continue

                    norm_name = normalize_emp_name(raw_name)
                    id_val = str(ws.cell(r, id_c).value or '').strip() if id_c else ''
                    dept_val = str(ws.cell(r, dept_c).value or '').strip() if dept_c else 'General'
                    basic_val = clean_fund_num(ws.cell(r, basic_c).value) if basic_c else 0.0
                    doj_val = ws.cell(r, doj_c).value if doj_c else None
                    parsed_doj = parse_accounting_date(doj_val) if doj_val else None

                    # Match Employee: Name first, then corporate code
                    emp = emp_by_name.get(norm_name)
                    if not emp and id_val and id_val.isdigit():
                        emp = emp_by_code.get(f"EMP-{int(id_val):03d}")

                    # Auto-create Employee if not found
                    if not emp:
                        new_code = get_next_unique_code(preferred_id=id_val)
                        emp = Employee(
                            employee_id=new_code,
                            name=raw_name,
                            designation="Staff Member",
                            department=dept_val or "General",
                            basic_salary=basic_val or 0.0,
                            status="Active",
                            joining_date=parsed_doj or date(2018, 1, 1)
                        )
                        db.session.add(emp)
                        db.session.flush()
                        emp_by_code[new_code.upper()] = emp
                        emp_by_name[norm_name] = emp
                        created_employees_count += 1

                    # Extract monthly contributions
                    for m_str, c_date, emp_col, comp_col, desc in month_columns:
                        v_emp = clean_fund_num(ws.cell(r, emp_col).value)
                        v_comp = clean_fund_num(ws.cell(r, comp_col).value)
                        if v_emp > 0 or v_comp > 0:
                            all_parsed_contributions.append({
                                'emp_id': emp.id,
                                'month': m_str,
                                'date': c_date,
                                'emp_share': v_emp,
                                'comp_share': v_comp,
                                'welfare_share': 0.0,
                                'profit': 0.0,
                                'remarks': desc
                            })

        else:
            # ------------------------------------------------------------------
            # MODE B: VERTICAL TABULAR LAYOUT (Standard Template / Flat Rows)
            # ------------------------------------------------------------------
            for ws in wb.worksheets:
                sheet_year = None
                yr_match = re.search(r'\b(20\d\d)\b', ws.title)
                if yr_match:
                    sheet_year = yr_match.group(1)

                header_row_idx = 1
                col_map = {}
                for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                    if r_idx > 15:
                        break
                    row_str = [str(c or '').strip().lower() for c in row]
                    if any('emp' in s or 'id' in s or 'code' in s for s in row_str) and any('share' in s or 'fund' in s or 'pf' in s or 'welfare' in s or 'month' in s or 'salary' in s or 'basic' in s for s in row_str):
                        header_row_idx = r_idx
                        for col_idx, val in enumerate(row_str):
                            if not val:
                                continue
                            if any(k in val for k in ['emp id', 'employee id', 'emp code', 'employee code', 'id no', 'emp_no', 'staff id']) and 'emp_id' not in col_map:
                                col_map['emp_id'] = col_idx
                            elif any(k in val for k in ['name', 'employee name', 'staff name', 'emp name', 'full name']) and 'name' not in col_map:
                                col_map['name'] = col_idx
                            elif any(k in val for k in ['designation', 'desig', 'position', 'post', 'rank']) and 'designation' not in col_map:
                                col_map['designation'] = col_idx
                            elif any(k in val for k in ['department', 'dept', 'section', 'branch']) and 'department' not in col_map:
                                col_map['department'] = col_idx
                            elif any(k in val for k in ['basic pay', 'basic salary', 'basic']) and 'basic' not in col_map:
                                col_map['basic'] = col_idx
                            elif any(k in val for k in ['month', 'period', 'salary month', 'contribution month', 'year-month']) and 'month' not in col_map:
                                col_map['month'] = col_idx
                            elif any(k in val for k in ['date', 'txn date', 'deposit date', 'posting date']) and 'date' not in col_map:
                                col_map['date'] = col_idx
                            elif any(k in val for k in ['employee share', 'emp share', 'staff share', 'pf own', 'own share', 'pf employee', 'pf emp', 'own contribution', 'employee contribution']) and 'emp_share' not in col_map:
                                col_map['emp_share'] = col_idx
                            elif any(k in val for k in ['employer share', 'company share', 'comp share', 'office share', 'pf company', 'pf employer', 'pf comp', 'employer contribution', 'office contribution']) and 'comp_share' not in col_map:
                                col_map['comp_share'] = col_idx
                            elif any(k in val for k in ['welfare', 'welfare fund', 'kallyan', 'kallan', 'wf', 'w.f']) and 'welfare' not in col_map:
                                col_map['welfare'] = col_idx
                            elif any(k in val for k in ['profit', 'interest', 'dividend', 'accrual', 'return']) and 'profit' not in col_map:
                                col_map['profit'] = col_idx
                            elif any(k in val for k in ['remarks', 'narration', 'note', 'description']) and 'remarks' not in col_map:
                                col_map['remarks'] = col_idx
                        break

                for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                    if not row or not any(row):
                        continue

                    raw_code = str(row[col_map.get('emp_id', 0)] if 'emp_id' in col_map and col_map['emp_id'] < len(row) else (row[0] or '')).strip()
                    if not raw_code:
                        continue
                    if any(x in raw_code.lower() for x in ['total', 'grand', 'subtotal', 'sum', 'page', 'signature']):
                        continue

                    clean_code = raw_code
                    if clean_code.isdigit():
                        clean_code = f"EMP-{int(clean_code):03d}"
                    elif re.match(r'^(?:emp|e)[-_]?(\d+)$', clean_code, re.I):
                        m_num = re.search(r'\d+', clean_code).group(0)
                        clean_code = f"EMP-{int(m_num):03d}"

                    raw_name = str(row[col_map['name']] if 'name' in col_map and col_map['name'] < len(row) and row[col_map['name']] else (row[1] if len(row) > 1 and row[1] and not str(row[1]).replace('.', '').isdigit() else "")).strip()
                    desig = str(row[col_map['designation']] if 'designation' in col_map and col_map['designation'] < len(row) and row[col_map['designation']] else "").strip()
                    dept = str(row[col_map['department']] if 'department' in col_map and col_map['department'] < len(row) and row[col_map['department']] else "").strip()
                    basic_sal = clean_fund_num(row[col_map['basic']]) if 'basic' in col_map and col_map['basic'] < len(row) else 0.0

                    raw_m = row[col_map['month']] if 'month' in col_map and col_map['month'] < len(row) else (row[2] if len(row) > 2 else None)
                    raw_d = row[col_map['date']] if 'date' in col_map and col_map['date'] < len(row) else (row[3] if len(row) > 3 else None)

                    c_month, c_date = parse_fund_excel_period(raw_m, raw_d, sheet_year)
                    if not c_month:
                        continue

                    emp_share = clean_fund_num(row[col_map['emp_share']]) if 'emp_share' in col_map and col_map['emp_share'] < len(row) else (clean_fund_num(row[4]) if len(row) > 4 else 0.0)
                    comp_share = clean_fund_num(row[col_map['comp_share']]) if 'comp_share' in col_map and col_map['comp_share'] < len(row) else (clean_fund_num(row[5]) if len(row) > 5 else 0.0)
                    welfare_share = clean_fund_num(row[col_map['welfare']]) if 'welfare' in col_map and col_map['welfare'] < len(row) else (clean_fund_num(row[6]) if len(row) > 6 and 'welfare' in col_map else 0.0)
                    profit = clean_fund_num(row[col_map['profit']]) if 'profit' in col_map and col_map['profit'] < len(row) else (clean_fund_num(row[7 if 'welfare' in col_map else 6]) if len(row) > (7 if 'welfare' in col_map else 6) else 0.0)
                    remarks = str(row[col_map['remarks']] if 'remarks' in col_map and col_map['remarks'] < len(row) and row[col_map['remarks']] else f"PF & Welfare contribution ({c_month})").strip()

                    norm_name = normalize_emp_name(raw_name)

                    # Match Employee by Name or Code
                    emp = None
                    if norm_name:
                        emp = emp_by_name.get(norm_name)
                    if not emp and clean_code:
                        emp = emp_by_code.get(clean_code.upper())

                    # Auto-create if missing
                    if not emp:
                        new_code = get_next_unique_code(preferred_id=clean_code)
                        emp = Employee(
                            employee_id=new_code,
                            name=raw_name or f"Employee {new_code}",
                            designation=desig or "Staff Member",
                            department=dept or "General",
                            basic_salary=basic_sal or 0.0,
                            status="Active",
                            joining_date=c_date or date(2018, 1, 1)
                        )
                        db.session.add(emp)
                        db.session.flush()
                        emp_by_code[new_code.upper()] = emp
                        if norm_name:
                            emp_by_name[norm_name] = emp
                        created_employees_count += 1

                    all_parsed_contributions.append({
                        'emp_id': emp.id,
                        'month': c_month,
                        'date': c_date or date(int(c_month.split('-')[0]), int(c_month.split('-')[1]), 28),
                        'emp_share': emp_share,
                        'comp_share': comp_share,
                        'welfare_share': welfare_share,
                        'profit': profit,
                        'remarks': remarks
                    })

        db.session.commit()

        # ----------------------------------------------------------------------
        # BATCH DATABASE INGESTION & CHRONOLOGICAL CUMULATIVE RECALCULATION
        # ----------------------------------------------------------------------
        # Group by employee
        emp_records_map = {}
        for r in all_parsed_contributions:
            e_id = r['emp_id']
            if e_id not in emp_records_map:
                emp_records_map[e_id] = []
            emp_records_map[e_id].append(r)

        imported_count = 0
        updated_count = 0

        # Pre-fetch existing contributions for fast lookup
        existing_cfs = {(c.employee_id, c.contribution_month): c for c in EmployeeFundContribution.query.all()}

        for emp_id, records in emp_records_map.items():
            # Sort chronologically from 2018 to 2026
            records.sort(key=lambda x: (x['month'], x['date']))

            for rec in records:
                m_str = rec['month']
                try:
                    all_years.add(int(m_str.split('-')[0]))
                except:
                    pass

                monthly_tot = round(rec['emp_share'] + rec['comp_share'] + rec['welfare_share'] + rec['profit'], 2)
                key = (emp_id, m_str)

                if key in existing_cfs:
                    ecf = existing_cfs[key]
                    ecf.employee_share = rec['emp_share']
                    ecf.employer_share = rec['comp_share']
                    ecf.welfare_fund = rec['welfare_share']
                    ecf.profit_interest = rec['profit']
                    ecf.contribution_date = rec['date']
                    ecf.total_monthly = monthly_tot
                    ecf.remarks = rec['remarks']
                    updated_count += 1
                else:
                    new_cf = EmployeeFundContribution(
                        employee_id=emp_id,
                        contribution_month=m_str,
                        contribution_date=rec['date'],
                        employee_share=rec['emp_share'],
                        employer_share=rec['comp_share'],
                        welfare_fund=rec['welfare_share'],
                        profit_interest=rec['profit'],
                        total_monthly=monthly_tot,
                        cumulative_balance=0.0,
                        remarks=rec['remarks']
                    )
                    db.session.add(new_cf)
                    existing_cfs[key] = new_cf
                    imported_count += 1

        db.session.commit()

        # Recalculate running balances for all affected employees
        for emp_id in emp_records_map.keys():
            recalculate_employee_fund_balances(emp_id)

        years_range_str = f"{min(all_years)} to {max(all_years)}" if all_years else "2018 to 2026"
        layout_name = "Employee-Wise Matrix Layout" if is_matrix_layout else "Tabular Layout"
        flash(
            f"Successfully processed {layout_name}: Auto-created {created_employees_count} new employees. "
            f"Imported {imported_count} new and updated {updated_count} fund records across {len(emp_records_map)} employees "
            f"spanning {years_range_str}. (Employee Contribution Fund remains 100% separate from General Accounts).",
            "success"
        )

    except Exception as e:
        db.session.rollback()
        flash(f'Error importing fund contributions from Excel: {str(e)}', 'danger')

    return redirect(url_for('fund_contributions_dashboard'))


@app.route('/payroll/fund-contributions/sample-excel')
def download_fund_sample_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contribution Import"

    headers = [
        "Emp ID", 
        "Employee Name", 
        "Month (YYYY-MM)", 
        "Date (YYYY-MM-DD)", 
        "PF Employee Share (৳)", 
        "PF Employer Share (৳)", 
        "Welfare Fund (৳)", 
        "Profit / Interest (৳)", 
        "Total Monthly (৳)", 
        "Remarks"
    ]
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sample_rows = [
        ["EMP-001", "Md. Moshiur Rahman", "2026-03", "2026-03-31", 4894.0, 4894.0, 200.0, 150.0, 10138.0, "March 2026 PF & Welfare Fund"],
        ["EMP-002", "Syed Mohiuddin Ahmed", "2026-03", "2026-03-31", 5939.0, 5939.0, 200.0, 180.0, 12258.0, "March 2026 PF & Welfare Fund"],
        ["EMP-003", "Md. Zillur Rahman", "2026-03", "2026-03-31", 3800.0, 3800.0, 200.0, 120.0, 7920.0, "March 2026 PF & Welfare Fund"]
    ]

    for r_idx, row_vals in enumerate(sample_rows, 2):
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 16)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Employee_Contribution_Fund_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/payroll/fund-contributions/add-manual', methods=['POST'])
def add_fund_contribution_manual():
    emp_id = request.form.get('employee_id', type=int)
    c_month = request.form.get('contribution_month')
    c_date_str = request.form.get('contribution_date')
    emp_share = float(request.form.get('employee_share') or 0.0)
    comp_share = float(request.form.get('employer_share') or 0.0)
    welfare_share = float(request.form.get('welfare_fund') or 0.0)
    profit = float(request.form.get('profit_interest') or 0.0)
    remarks = request.form.get('remarks', 'Manual entry')

    c_date = parse_accounting_date(c_date_str) or date.today()
    if isinstance(c_date, datetime):
        c_date = c_date.date()

    latest_prev = EmployeeFundContribution.query.filter_by(employee_id=emp_id).order_by(EmployeeFundContribution.contribution_date.desc(), EmployeeFundContribution.id.desc()).first()
    prev_cumul = latest_prev.cumulative_balance if latest_prev else 0.0

    monthly_tot = round(emp_share + comp_share + welfare_share + profit, 2)
    new_cumul = round(prev_cumul + monthly_tot, 2)

    cf = EmployeeFundContribution(
        employee_id=emp_id,
        contribution_month=c_month,
        contribution_date=c_date,
        employee_share=emp_share,
        employer_share=comp_share,
        welfare_fund=welfare_share,
        profit_interest=profit,
        total_monthly=monthly_tot,
        cumulative_balance=new_cumul,
        remarks=remarks
    )
    db.session.add(cf)
    db.session.commit()

    flash('Fund contribution saved successfully.', 'success')
    return redirect(url_for('employee_fund_ledger', emp_id=emp_id))


def recalculate_employee_fund_balances(emp_id):
    """Recalculates sequential cumulative balances for an employee's fund records."""
    records = EmployeeFundContribution.query.filter_by(employee_id=emp_id).order_by(
        EmployeeFundContribution.contribution_date.asc(), 
        EmployeeFundContribution.id.asc()
    ).all()
    running = 0.0
    for r in records:
        emp_share = r.employee_share or 0.0
        comp_share = r.employer_share or 0.0
        welfare = r.welfare_fund or 0.0
        profit = r.profit_interest or 0.0
        r.total_monthly = round(emp_share + comp_share + welfare + profit, 2)
        running = round(running + r.total_monthly, 2)
        r.cumulative_balance = running
    db.session.commit()


@app.route('/payroll/fund-contributions/clear', methods=['POST'])
def clear_fund_contributions():
    """
    Clears previous data for Employee Contribution Fund (Provident & Welfare).
    Supports:
    - ALL records (total fund wipe to start fresh)
    - By date range (From Date to To Date)
    - By specific month (e.g. 2026-03)
    - For all employees or a selected employee
    Maintains 100% complete isolation from Accounts / General Ledger.
    """
    clear_scope = request.form.get('clear_scope', 'all')
    emp_filter = request.form.get('employee_id', type=int)
    from_str = request.form.get('from_date', '').strip()
    to_str = request.form.get('to_date', '').strip()
    month_str = request.form.get('month_str', '').strip()
    redirect_emp_id = request.form.get('redirect_emp_id', type=int)

    query = EmployeeFundContribution.query

    if emp_filter and emp_filter > 0:
        query = query.filter(EmployeeFundContribution.employee_id == emp_filter)

    scope_desc = "all previous records"
    if clear_scope == 'date_range' and from_str and to_str:
        try:
            f_date = parse_accounting_date(from_str) or datetime.strptime(from_str, '%Y-%m-%d').date()
            t_date = parse_accounting_date(to_str) or datetime.strptime(to_str, '%Y-%m-%d').date()
            if isinstance(f_date, datetime):
                f_date = f_date.date()
            if isinstance(t_date, datetime):
                t_date = t_date.date()
            if f_date > t_date:
                flash('From Date cannot be later than To Date.', 'warning')
                return redirect(url_for('fund_contributions_dashboard'))
            query = query.filter(
                EmployeeFundContribution.contribution_date >= f_date, 
                EmployeeFundContribution.contribution_date <= t_date
            )
            scope_desc = f"records between {f_date.strftime('%d-%b-%Y')} and {t_date.strftime('%d-%b-%Y')}"
        except Exception as e:
            flash(f'Invalid date range specified: {str(e)}', 'danger')
            return redirect(url_for('fund_contributions_dashboard'))

    elif clear_scope == 'month' and month_str:
        query = query.filter(EmployeeFundContribution.contribution_month == month_str)
        scope_desc = f"records for month {month_str}"

    matching_records = query.all()
    count_deleted = len(matching_records)
    affected_emp_ids = list(set([r.employee_id for r in matching_records]))

    for r in matching_records:
        db.session.delete(r)
    db.session.commit()

    # Recalculate chronological running balances for any remaining records
    for e_id in affected_emp_ids:
        recalculate_employee_fund_balances(e_id)

    emp_name_info = ""
    if emp_filter and emp_filter > 0:
        emp_obj = Employee.query.get(emp_filter)
        if emp_obj:
            emp_name_info = f" for {emp_obj.name} ({emp_obj.employee_id})"

    flash(f'Successfully cleared {count_deleted} Employee Contribution Fund records ({scope_desc}{emp_name_info}). General Accounts and Ledger calculations are completely untouched.', 'success')

    if redirect_emp_id:
        return redirect(url_for('employee_fund_ledger', emp_id=redirect_emp_id))
    return redirect(url_for('fund_contributions_dashboard'))


# ==============================================================================
# PREVIOUS SALARY HISTORY & COMPENSATION REGISTRY MODULE (34-COLUMN SPECIFICATION)
# ==============================================================================

_SCHEMA_INITIALIZED = False

def ensure_salary_history_schema():
    """Ensure all required columns exist in SQLite database for employee_salary_history."""
    global _SCHEMA_INITIALIZED
    if _SCHEMA_INITIALIZED:
        return
    try:
        from sqlalchemy import inspect, text
        inspector = inspect(db.engine)
        existing_cols = {col['name'] for col in inspector.get_columns('employee_salary_history')}
        cols = [
            ('department', 'VARCHAR(100)'),
            ('designation', 'VARCHAR(100)'),
            ('joining_date', 'DATE'),
            ('pf_rate', 'FLOAT DEFAULT 0.0'),
            ('pf_employer', 'FLOAT DEFAULT 0.0'),
            ('pf_employee', 'FLOAT DEFAULT 0.0'),
            ('pf_total', 'FLOAT DEFAULT 0.0'),
            ('salary_payable', 'FLOAT DEFAULT 0.0'),
            ('govt_holiday', 'FLOAT DEFAULT 0.0'),
            ('holiday_extra_duty', 'FLOAT DEFAULT 0.0'),
            ('loyalty_allowance', 'FLOAT DEFAULT 0.0'),
            ('additional_duty', 'FLOAT DEFAULT 0.0'),
            ('office_order_allowance', 'FLOAT DEFAULT 0.0'),
            ('eid_bonus', 'FLOAT DEFAULT 0.0'),
            ('other_addition', 'FLOAT DEFAULT 0.0'),
            ('working_days', 'FLOAT DEFAULT 0.0'),
            ('absent_days', 'FLOAT DEFAULT 0.0'),
            ('late_days', 'FLOAT DEFAULT 0.0'),
            ('absenteeism_amount', 'FLOAT DEFAULT 0.0'),
            ('lateness_amount', 'FLOAT DEFAULT 0.0'),
            ('arrear', 'FLOAT DEFAULT 0.0'),
            ('tds', 'FLOAT DEFAULT 0.0'),
            ('net_pay_after_tds', 'FLOAT DEFAULT 0.0'),
            ('company_total_paid', 'FLOAT DEFAULT 0.0'),
        ]
        missing = [c for c in cols if c[0] not in existing_cols]
        if missing:
            with db.engine.begin() as conn:
                for col_name, col_type in missing:
                    conn.execute(text(f"ALTER TABLE employee_salary_history ADD COLUMN {col_name} {col_type}"))
        _SCHEMA_INITIALIZED = True
    except Exception as e:
        _SCHEMA_INITIALIZED = True

def normalize_salary_month(val):
    if not val:
        return datetime.today().strftime('%Y-%m')
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    if re.match(r'^\d{4}-\d{2}$', s):
        return s
    d = parse_accounting_date(s)
    if d:
        return d.strftime('%Y-%m')
    for fmt in ("%B %Y", "%b %Y", "%b-%y", "%b-%Y", "%m-%Y", "%m/%Y", "%Y/%m", "%m-%d-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m')
        except Exception:
            pass
    return s


@app.route('/payroll/salary-history')
def salary_history_dashboard():
    ensure_salary_history_schema()
    emp_id = request.args.get('employee_id', type=int)
    search_q = request.args.get('search', '').strip()
    month = request.args.get('month', 'all')

    query = EmployeeSalaryHistory.query.join(Employee)
    if emp_id:
        query = query.filter(EmployeeSalaryHistory.employee_id == emp_id)
    elif search_q:
        query = query.filter(
            (Employee.employee_id.ilike(f"%{search_q}%")) |
            (Employee.name.ilike(f"%{search_q}%"))
        )
    if month and month != 'all':
        query = query.filter(EmployeeSalaryHistory.salary_month == month)

    histories = query.order_by(EmployeeSalaryHistory.salary_month.desc(), Employee.employee_id.asc()).all()
    employees = Employee.query.order_by(Employee.employee_id.asc()).all()

    all_months = db.session.query(EmployeeSalaryHistory.salary_month).distinct().order_by(EmployeeSalaryHistory.salary_month.desc()).all()
    month_list = [m[0] for m in all_months if m[0]]

    total_basic = sum(h.basic_salary or 0.0 for h in histories)
    total_gross = sum(h.gross_salary or 0.0 for h in histories)
    total_net = sum(h.display_net for h in histories)
    total_ded = sum(h.total_deductions or 0.0 for h in histories)
    total_tds = sum(h.tds or h.tax_deduction or 0.0 for h in histories)
    total_pf_empr = sum(h.pf_employer or h.pf_deduction or 0.0 for h in histories)
    total_company_paid = sum(h.display_company_paid for h in histories)

    return render_template(
        'payroll/salary_history.html',
        histories=histories,
        employees=employees,
        month_list=month_list,
        selected_emp_id=emp_id,
        selected_month=month,
        search_q=search_q,
        total_basic=total_basic,
        total_gross=total_gross,
        total_net=total_net,
        total_ded=total_ded,
        total_tds=total_tds,
        total_pf_empr=total_pf_empr,
        total_company_paid=total_company_paid
    )


def format_excel_emp_code(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v)).strip()
    if isinstance(v, int) and not isinstance(v, bool):
        return str(v).strip()
    s = str(v).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        return s[:-2]
    return s

def is_excel_helper_or_total_row(row, emp_code, emp_name):
    code_lower = str(emp_code).strip().lower()
    name_lower = str(emp_name).strip().lower()
    keywords = {
        'id', 'idno', 'id no', 'sl', 'sl no', 'serial', 'code', 'emp id',
        'employee id', 'name', 'employee name', 'month', 'total', 'grand total',
        'sub total', 'subtotal', 'sum', 'nan', 'none'
    }
    if code_lower in keywords or name_lower in keywords:
        return True

    # Detect column number indicator rows: e.g. [1, 2, 3, 4, 5, ...]
    numeric_vals = []
    for cell in row[:8]:
        if cell is not None and str(cell).strip() != '':
            try:
                num = int(float(str(cell).strip().strip('()[]')))
                numeric_vals.append(num)
            except ValueError:
                pass
    if len(numeric_vals) >= 4:
        diffs = [numeric_vals[i+1] - numeric_vals[i] for i in range(len(numeric_vals)-1)]
        if all(d == 1 for d in diffs):
            return True

    try:
        c_num = int(float(code_lower))
        n_num = int(float(name_lower))
        if n_num == c_num + 1 or n_num == c_num:
            return True
    except ValueError:
        pass

    return False


@app.route('/payroll/salary-history/import-excel', methods=['POST'])
def import_salary_history_excel():
    ensure_salary_history_schema()
    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Please select an Excel file containing previous salary data.', 'warning')
        return redirect(url_for('salary_history_dashboard'))

    try:
        wb = openpyxl.load_workbook(file.stream, data_only=True)
        ws = wb.active
        imported = 0
        created_employees = 0
        updated_employees = 0

        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            flash('The uploaded Excel file is empty.', 'warning')
            return redirect(url_for('salary_history_dashboard'))

        # Dynamically find the real header row by scanning rows 1 to 10
        header_tokens = {
            'month', 'salarymonth', 'idno', 'id', 'empid', 'code', 'employeeno',
            'name', 'employeename', 'fullname', 'department', 'dept',
            'designation', 'desig', 'position', 'dateofjoining', 'doj', 'joiningdate',
            'basic', 'basicsalary', 'hrallowance', 'houserent', 'hra',
            'medicalallowance', 'medical', 'conveyanceallowance', 'conveyance',
            'grosspay', 'grosssalary', 'gross', 'pf', 'pfrate', 'pfemployer',
            'pfemployee', 'total', 'salary', 'salarypayable', 'workingday',
            'netpay', 'netsalary', 'tds', 'netpayaftertds'
        }

        header_row_idx = 0
        header_row = []
        col_map = {}
        max_header_matches = 0

        for r_i, r in enumerate(all_rows[:10]):
            if not r or not any(r):
                continue
            cur_map = {}
            matches = 0
            for idx, h in enumerate(r):
                if h is not None:
                    norm_h = re.sub(r'[^a-z0-9]', '', str(h).lower())
                    if norm_h:
                        cur_map[norm_h] = idx
                        if norm_h in header_tokens:
                            matches += 1
            if matches > max_header_matches:
                max_header_matches = matches
                header_row_idx = r_i
                header_row = r
                col_map = cur_map

        is_34_col = (
            max_header_matches >= 2 or
            any(k in col_map for k in ['idno', 'grosspay', 'netpayaftertds', 'loyaltyallowance', 'workingday']) or
            len(header_row) >= 20
        )

        # Pre-cache all existing employees in memory to eliminate SQLite lock contention and duplicates
        existing_emps = {str(e.employee_id).strip(): e for e in Employee.query.all()}
        existing_emps_ci = {str(e.employee_id).strip().lower(): e for e in Employee.query.all()}
        new_emps_to_add = {}

        # 1. Parse all rows and extract data
        parsed_records = []
        for row in all_rows[header_row_idx + 1:]:
            if not row or not any(row):
                continue

            def get_val(key_list, fallback_idx, default=""):
                for k in key_list:
                    if k in col_map and col_map[k] < len(row):
                        v = row[col_map[k]]
                        if v is not None and str(v).strip() != "":
                            return v
                if fallback_idx is not None and fallback_idx < len(row):
                    v = row[fallback_idx]
                    if v is not None and str(v).strip() != "":
                        return v
                return default

            if is_34_col:
                emp_code = format_excel_emp_code(get_val(['idno', 'id', 'empid', 'code', 'employeeno'], 1, ''))
                emp_name = str(get_val(['name', 'employeename', 'fullname'], 2, f"Staff {emp_code}")).strip()

                # Skip non-data rows (column indexing numbers, total rows, empty code)
                if not emp_code or is_excel_helper_or_total_row(row, emp_code, emp_name):
                    continue

                raw_month = get_val(['month', 'salarymonth', 'date'], 0, datetime.today().strftime('%Y-%m'))
                s_month = normalize_salary_month(raw_month)

                emp_dept = str(get_val(['department', 'dept'], 3, "General")).strip()
                emp_desig = str(get_val(['designation', 'desig', 'position'], 4, "Staff")).strip()
                raw_doj = get_val(['dateofjoining', 'doj', 'joiningdate'], 5, None)
                emp_doj = parse_accounting_date(raw_doj)

                b_sal = clean_accounting_num(get_val(['basic', 'basicsalary'], 6, 0.0))
                hr_all = clean_accounting_num(get_val(['hrallowance', 'houserent', 'hra'], 7, 0.0))
                med_all = clean_accounting_num(get_val(['medicalallowance', 'medical'], 8, 0.0))
                conv_all = clean_accounting_num(get_val(['conveyanceallowance', 'conveyance'], 9, 0.0))
                gross = clean_accounting_num(get_val(['grosspay', 'grosssalary', 'gross'], 10, (b_sal + hr_all + med_all + conv_all)))

                pf_rate = clean_accounting_num(get_val(['pf', 'pfrate'], 11, 0.0))
                pf_empr = clean_accounting_num(get_val(['pfemployer', 'employerpf'], 12, 0.0))
                pf_emp = clean_accounting_num(get_val(['pfemployee', 'employeepf'], 13, 0.0))
                pf_tot = clean_accounting_num(get_val(['total', 'pftotal', 'totalpf'], 14, (pf_empr + pf_emp)))

                sal_pay = clean_accounting_num(get_val(['salary', 'salarypayable'], 15, b_sal))
                govt_hol = clean_accounting_num(get_val(['govtholiday', 'govtholidays'], 16, 0.0))
                hol_duty = clean_accounting_num(get_val(['holidayextradutyallowance', 'extradutyallowance'], 17, 0.0))
                loyalty = clean_accounting_num(get_val(['loyaltyallowance755yrs1010yrs', 'loyaltyallowance'], 18, 0.0))
                add_duty = clean_accounting_num(get_val(['additionaldutyamountforholidayoffice', 'additionaldutyamountfor', 'additionalduty'], 19, 0.0))
                off_order = clean_accounting_num(get_val(['allowanceforofficeorder', 'officeorderallowance'], 20, 0.0))
                eid = clean_accounting_num(get_val(['eidbonus', 'bonus'], 21, 0.0))
                oth_add = clean_accounting_num(get_val(['other', 'otherallowance'], 22, 0.0))

                w_days = clean_accounting_num(get_val(['workingday', 'workingdays'], 23, 26))
                abs_days = clean_accounting_num(get_val(['absent', 'absentdays'], 24, 0))
                late_days = clean_accounting_num(get_val(['late', 'latedays'], 25, 0))
                abs_amt = clean_accounting_num(get_val(['absenteeismamount', 'absentamount'], 26, 0.0))
                late_amt = clean_accounting_num(get_val(['latenessamount', 'lateamount'], 27, 0.0))
                oth_ded = clean_accounting_num(get_val(['otherdeduction', 'otherdeductions'], 28, 0.0))
                tot_ded = clean_accounting_num(get_val(['deduction', 'deductions', 'totaldeductions'], 29, (abs_amt + late_amt + oth_ded + pf_emp)))

                arr = clean_accounting_num(get_val(['arrear', 'arrears'], 30, 0.0))
                net_pay = clean_accounting_num(get_val(['netpay', 'netsalary'], 31, 0.0))
                tds_val = clean_accounting_num(get_val(['tds', 'tax'], 32, 0.0))
                net_after_tds = clean_accounting_num(get_val(['netpayaftertds', 'nettakehome'], 33, 0.0))

                if gross <= 0:
                    gross = round(b_sal + hr_all + med_all + conv_all, 2)
                if net_pay <= 0:
                    tot_allw = govt_hol + hol_duty + loyalty + add_duty + off_order + eid + oth_add
                    net_pay = round(gross + tot_allw - tot_ded + arr, 2)
                if net_after_tds <= 0:
                    net_after_tds = round(net_pay - tds_val, 2)
                co_paid = round(net_after_tds + tds_val + pf_empr, 2)
                p_date = date.today()

            else:
                raw_col0 = row[0]
                raw_col1 = row[1] if len(row) > 1 else None

                # Check if col 0 is a month/date and col 1 is employee code
                if isinstance(raw_col0, (datetime, date)) or (raw_col0 and re.match(r'^\d{4}[-/]\d{2}', str(raw_col0))):
                    s_month = normalize_salary_month(raw_col0)
                    emp_code = format_excel_emp_code(raw_col1)
                    emp_name = str(row[2]).strip() if len(row) > 2 and row[2] else f"Staff {emp_code}"
                else:
                    emp_code = format_excel_emp_code(raw_col0)
                    s_month = normalize_salary_month(raw_col1) if raw_col1 else datetime.today().strftime('%Y-%m')
                    emp_name = f"Staff {emp_code}"

                if not emp_code or is_excel_helper_or_total_row(row, emp_code, emp_name):
                    continue

                p_date = parse_accounting_date(row[2]) if len(row) > 2 and row[2] else date.today()
                b_sal = clean_accounting_num(row[3]) if len(row) > 3 else 0.0
                hr_all = clean_accounting_num(row[4]) if len(row) > 4 else round(b_sal * 0.4, 2)
                med_all = clean_accounting_num(row[5]) if len(row) > 5 else round(b_sal * 0.1, 2)
                conv_all = clean_accounting_num(row[6]) if len(row) > 6 else round(b_sal * 0.05, 2)
                oth_add = clean_accounting_num(row[7]) if len(row) > 7 else 0.0
                gross = clean_accounting_num(row[8]) if len(row) > 8 else (b_sal + hr_all + med_all + conv_all + oth_add)
                pf_emp = clean_accounting_num(row[9]) if len(row) > 9 else round(b_sal * 0.08, 2)
                pf_empr = pf_emp
                pf_tot = pf_emp * 2
                pf_rate = 0.08
                tds_val = clean_accounting_num(row[10]) if len(row) > 10 else 0.0
                oth_ded = clean_accounting_num(row[11]) if len(row) > 11 else 0.0
                tot_ded = clean_accounting_num(row[12]) if len(row) > 12 else (pf_emp + tds_val + oth_ded)
                net_pay = clean_accounting_num(row[13]) if len(row) > 13 else (gross - tot_ded)
                net_after_tds = round(net_pay - tds_val, 2) if net_pay > tds_val else net_pay
                co_paid = round(net_after_tds + tds_val + pf_empr, 2)

                emp_dept = "General"
                emp_desig = "Staff"
                emp_doj = None
                sal_pay = b_sal
                govt_hol = hol_duty = loyalty = add_duty = off_order = eid = 0.0
                w_days = 26
                abs_days = late_days = abs_amt = late_amt = arr = 0.0

            # Store for batch processing
            parsed_records.append({
                'emp_code': emp_code,
                'emp_name': emp_name,
                'emp_dept': emp_dept,
                'emp_desig': emp_desig,
                'emp_doj': emp_doj,
                's_month': s_month,
                'p_date': p_date,
                'b_sal': b_sal,
                'hr_all': hr_all,
                'med_all': med_all,
                'conv_all': conv_all,
                'gross': gross,
                'pf_rate': pf_rate,
                'pf_empr': pf_empr,
                'pf_emp': pf_emp,
                'pf_tot': pf_tot,
                'sal_pay': sal_pay,
                'govt_hol': govt_hol,
                'hol_duty': hol_duty,
                'loyalty': loyalty,
                'add_duty': add_duty,
                'off_order': off_order,
                'eid': eid,
                'oth_add': oth_add,
                'w_days': w_days,
                'abs_days': abs_days,
                'late_days': late_days,
                'abs_amt': abs_amt,
                'late_amt': late_amt,
                'oth_ded': oth_ded,
                'tot_ded': tot_ded,
                'arr': arr,
                'net_pay': net_pay,
                'tds_val': tds_val,
                'net_after_tds': net_after_tds,
                'co_paid': co_paid
            })

            # Check Employee database in memory (unique ID preserved, no duplicates)
            code_ci = emp_code.lower()
            existing_e = existing_emps.get(emp_code) or existing_emps_ci.get(code_ci) or new_emps_to_add.get(emp_code) or new_emps_to_add.get(code_ci)

            if not existing_e:
                new_emp = Employee(
                    employee_id=emp_code,
                    name=emp_name or f"Staff {emp_code}",
                    department=emp_dept or "General",
                    designation=emp_desig or "Staff",
                    joining_date=emp_doj,
                    basic_salary=b_sal,
                    house_rent=hr_all,
                    medical_allowance=med_all,
                    conveyance=conv_all,
                    status="Active"
                )
                new_emps_to_add[emp_code] = new_emp
                new_emps_to_add[code_ci] = new_emp
                created_employees += 1
            else:
                if emp_name and not emp_name.startswith("Staff ") and (not existing_e.name or existing_e.name.startswith("Staff ")):
                    existing_e.name = emp_name
                if emp_dept and emp_dept != "General" and not existing_e.department:
                    existing_e.department = emp_dept
                if emp_desig and emp_desig != "Staff" and not existing_e.designation:
                    existing_e.designation = emp_desig
                if emp_doj and not existing_e.joining_date:
                    existing_e.joining_date = emp_doj
                if b_sal > 0 and (not existing_e.basic_salary or existing_e.basic_salary == 0):
                    existing_e.basic_salary = b_sal
                updated_employees += 1

        # 2. Batch commit new employees to get real database IDs
        unique_new_emps = list({id(e): e for e in new_emps_to_add.values()}.values())
        if unique_new_emps:
            db.session.add_all(unique_new_emps)
            db.session.commit()
            for e in unique_new_emps:
                code_str = str(e.employee_id).strip()
                existing_emps[code_str] = e
                existing_emps_ci[code_str.lower()] = e

        # 3. Load existing salary histories for the relevant employees and months
        emp_obj_map = {}
        for r in parsed_records:
            code = r['emp_code']
            code_ci = code.lower()
            emp_obj = existing_emps.get(code) or existing_emps_ci.get(code_ci)
            if emp_obj:
                emp_obj_map[code] = emp_obj

        emp_ids = [emp.id for emp in emp_obj_map.values()]
        existing_sh_query = EmployeeSalaryHistory.query.filter(EmployeeSalaryHistory.employee_id.in_(emp_ids)).all() if emp_ids else []
        sh_map = {(sh.employee_id, sh.salary_month): sh for sh in existing_sh_query}

        # 4. Create/update salary histories in memory
        for r in parsed_records:
            emp = emp_obj_map.get(r['emp_code'])
            if not emp:
                continue

            sh_key = (emp.id, r['s_month'])
            sh = sh_map.get(sh_key)
            if not sh:
                sh = EmployeeSalaryHistory(employee_id=emp.id, salary_month=r['s_month'], payment_date=r['p_date'])
                db.session.add(sh)
                sh_map[sh_key] = sh

            sh.payment_date = r['p_date']
            sh.department = r['emp_dept'] or emp.department
            sh.designation = r['emp_desig'] or emp.designation
            sh.joining_date = r['emp_doj'] or emp.joining_date
            sh.basic_salary = r['b_sal']
            sh.house_rent = r['hr_all']
            sh.medical = r['med_all']
            sh.conveyance = r['conv_all']
            sh.gross_salary = r['gross']
            sh.pf_rate = r['pf_rate']
            sh.pf_employer = r['pf_empr']
            sh.pf_employee = r['pf_emp']
            sh.pf_total = r['pf_tot']
            sh.salary_payable = r['sal_pay']
            sh.govt_holiday = r['govt_hol']
            sh.holiday_extra_duty = r['hol_duty']
            sh.loyalty_allowance = r['loyalty']
            sh.additional_duty = r['add_duty']
            sh.office_order_allowance = r['off_order']
            sh.eid_bonus = r['eid']
            sh.other_addition = r['oth_add']
            sh.working_days = r['w_days']
            sh.absent_days = r['abs_days']
            sh.late_days = r['late_days']
            sh.absenteeism_amount = r['abs_amt']
            sh.lateness_amount = r['late_amt']
            sh.other_deduction = r['oth_ded']
            sh.total_deductions = r['tot_ded']
            sh.arrear = r['arr']
            sh.net_salary = r['net_pay']
            sh.tds = r['tds_val']
            sh.net_pay_after_tds = r['net_after_tds']
            sh.company_total_paid = r['co_paid']
            sh.pf_deduction = r['pf_emp']
            sh.tax_deduction = r['tds_val']
            sh.other_allowance = (r['govt_hol'] + r['hol_duty'] + r['loyalty'] + r['add_duty'] + r['off_order'] + r['eid'] + r['oth_add'])

            imported += 1

        # 5. Commit all records in one transaction
        db.session.commit()
        flash(f'Successfully imported {imported} historical salary records. (Created {created_employees} new employee profiles, synced {updated_employees} existing profiles with unique ID).', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Error importing salary history Excel: {str(e)}', 'danger')

    return redirect(url_for('salary_history_dashboard'))


@app.route('/payroll/salary-history/clear', methods=['POST'])
def clear_salary_history():
    """Clear Historical Salary & Compensation Registry data as per date or month without touching employee database."""
    ensure_salary_history_schema()
    clear_mode = request.form.get('clear_mode', 'month') # 'month', 'date_range', 'all'
    confirm = request.form.get('confirm_clear')

    if confirm != 'yes':
        flash('Operation cancelled: Please tick the confirmation box to authorize clearing historical salary data.', 'warning')
        return redirect(url_for('salary_history_dashboard'))

    try:
        deleted_count = 0
        desc = ""
        if clear_mode == 'month':
            target_month = request.form.get('target_month', '').strip()
            if not target_month:
                flash('Please select or specify the salary month to clear.', 'warning')
                return redirect(url_for('salary_history_dashboard'))
            deleted_count = EmployeeSalaryHistory.query.filter_by(salary_month=target_month).delete()
            desc = f"month {target_month}"

        elif clear_mode == 'date_range':
            start_date_str = request.form.get('start_date')
            end_date_str = request.form.get('end_date')
            start_d = parse_accounting_date(start_date_str)
            end_d = parse_accounting_date(end_date_str)
            if not start_d or not end_d:
                flash('Please specify valid start and end dates.', 'warning')
                return redirect(url_for('salary_history_dashboard'))
            deleted_count = EmployeeSalaryHistory.query.filter(
                EmployeeSalaryHistory.payment_date >= start_d,
                EmployeeSalaryHistory.payment_date <= end_d
            ).delete()
            desc = f"date range {start_d.strftime('%d-%b-%Y')} to {end_d.strftime('%d-%b-%Y')}"

        elif clear_mode == 'all':
            deleted_count = EmployeeSalaryHistory.query.delete()
            desc = "all historical salary periods"

        db.session.commit()
        flash(f'Successfully cleared {deleted_count} historical salary records for {desc}. All Employee database records (IDs, names, departments) remain safely preserved.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error clearing historical salary data: {str(e)}', 'danger')

    return redirect(url_for('salary_history_dashboard'))


@app.route('/payroll/salary-history/add-record', methods=['POST'])
def add_salary_history_record():
    """Manual input of a single historical compensation record matching the 34-column registry."""
    ensure_salary_history_schema()
    try:
        emp_code = request.form.get('employee_id_code', '').strip()
        if not emp_code:
            flash('Employee ID number is required.', 'danger')
            return redirect(url_for('salary_history_dashboard'))

        s_month = normalize_salary_month(request.form.get('salary_month'))
        p_date = parse_accounting_date(request.form.get('payment_date')) or date.today()
        emp_name = request.form.get('employee_name', '').strip()
        dept = request.form.get('department', '').strip()
        desig = request.form.get('designation', '').strip()
        doj = parse_accounting_date(request.form.get('joining_date'))

        b_sal = clean_accounting_num(request.form.get('basic_salary'))
        hr_all = clean_accounting_num(request.form.get('house_rent'))
        med_all = clean_accounting_num(request.form.get('medical'))
        conv_all = clean_accounting_num(request.form.get('conveyance'))
        gross = clean_accounting_num(request.form.get('gross_salary')) or (b_sal + hr_all + med_all + conv_all)

        pf_rate = clean_accounting_num(request.form.get('pf_rate'))
        pf_empr = clean_accounting_num(request.form.get('pf_employer'))
        pf_emp = clean_accounting_num(request.form.get('pf_employee'))
        pf_tot = clean_accounting_num(request.form.get('pf_total')) or (pf_empr + pf_emp)

        sal_pay = clean_accounting_num(request.form.get('salary_payable')) or b_sal
        govt_hol = clean_accounting_num(request.form.get('govt_holiday'))
        hol_duty = clean_accounting_num(request.form.get('holiday_extra_duty'))
        loyalty = clean_accounting_num(request.form.get('loyalty_allowance'))
        add_duty = clean_accounting_num(request.form.get('additional_duty'))
        off_order = clean_accounting_num(request.form.get('office_order_allowance'))
        eid = clean_accounting_num(request.form.get('eid_bonus'))
        oth_add = clean_accounting_num(request.form.get('other_addition'))

        w_days = clean_accounting_num(request.form.get('working_days')) or 26.0
        abs_days = clean_accounting_num(request.form.get('absent_days'))
        late_days = clean_accounting_num(request.form.get('late_days'))
        abs_amt = clean_accounting_num(request.form.get('absenteeism_amount'))
        late_amt = clean_accounting_num(request.form.get('lateness_amount'))
        oth_ded = clean_accounting_num(request.form.get('other_deduction'))
        tot_ded = clean_accounting_num(request.form.get('total_deductions')) or (abs_amt + late_amt + oth_ded + pf_emp)

        arr = clean_accounting_num(request.form.get('arrear'))
        tot_allw = (govt_hol + hol_duty + loyalty + add_duty + off_order + eid + oth_add)
        net_pay = clean_accounting_num(request.form.get('net_salary')) or (gross + tot_allw - tot_ded + arr)
        tds_val = clean_accounting_num(request.form.get('tds'))
        net_after_tds = clean_accounting_num(request.form.get('net_pay_after_tds')) or (net_pay - tds_val)
        co_paid = clean_accounting_num(request.form.get('company_total_paid')) or (net_after_tds + tds_val + pf_empr)

        # Unique Employee logic: find or create, no duplicate!
        emp = Employee.query.filter_by(employee_id=emp_code).first()
        if not emp:
            emp = Employee(
                employee_id=emp_code,
                name=emp_name or f"Staff {emp_code}",
                department=dept or "General",
                designation=desig or "Staff",
                joining_date=doj,
                basic_salary=b_sal,
                house_rent=hr_all,
                medical_allowance=med_all,
                conveyance=conv_all,
                status="Active"
            )
            db.session.add(emp)
            db.session.flush()
        else:
            if emp_name: emp.name = emp_name
            if dept: emp.department = dept
            if desig: emp.designation = desig
            if doj and not emp.joining_date: emp.joining_date = doj
            if b_sal > 0: emp.basic_salary = b_sal

        sh = EmployeeSalaryHistory.query.filter_by(employee_id=emp.id, salary_month=s_month).first()
        if not sh:
            sh = EmployeeSalaryHistory(employee_id=emp.id, salary_month=s_month, payment_date=p_date)
            db.session.add(sh)

        sh.payment_date = p_date
        sh.department = dept or emp.department
        sh.designation = desig or emp.designation
        sh.joining_date = doj or emp.joining_date
        sh.basic_salary = b_sal
        sh.house_rent = hr_all
        sh.medical = med_all
        sh.conveyance = conv_all
        sh.gross_salary = gross
        sh.pf_rate = pf_rate
        sh.pf_employer = pf_empr
        sh.pf_employee = pf_emp
        sh.pf_total = pf_tot
        sh.salary_payable = sal_pay
        sh.govt_holiday = govt_hol
        sh.holiday_extra_duty = hol_duty
        sh.loyalty_allowance = loyalty
        sh.additional_duty = add_duty
        sh.office_order_allowance = off_order
        sh.eid_bonus = eid
        sh.other_addition = oth_add
        sh.working_days = w_days
        sh.absent_days = abs_days
        sh.late_days = late_days
        sh.absenteeism_amount = abs_amt
        sh.lateness_amount = late_amt
        sh.other_deduction = oth_ded
        sh.total_deductions = tot_ded
        sh.arrear = arr
        sh.net_salary = net_pay
        sh.tds = tds_val
        sh.net_pay_after_tds = net_after_tds
        sh.company_total_paid = co_paid
        sh.pf_deduction = pf_emp
        sh.tax_deduction = tds_val
        sh.other_allowance = tot_allw

        db.session.commit()
        flash(f'Historical compensation record for {emp.employee_id} - {emp.name} ({s_month}) saved successfully.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error saving record: {str(e)}', 'danger')

    return redirect(url_for('salary_history_dashboard'))


@app.route('/payroll/salary-history/sample-excel')
def download_salary_history_sample_excel():
    """Generate exact 34-column Excel import template matching client historical salary registry image."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salary & Compensation Registry"

    headers = [
        "Month",
        "ID no.",
        "Name",
        "Department",
        "Designation",
        "Date of Joining",
        "Basic",
        "HR Allowance",
        "Medical Allowance",
        "Conveyance Allowance",
        "Gross Pay",
        "PF",
        "PF Employer",
        "PF Employee",
        "Total",
        "Salary",
        "Govt. Holiday",
        "Holiday & Extra Duty Allowance",
        "Loyalty Allowance 7.5% 5yrs & 10% >= 10yrs",
        "Additional duty amount for holiday/office",
        "Allowance for office order",
        "Eid Bonus",
        "Other",
        "Working Day",
        "Absent",
        "Late",
        "Absenteeism Amount",
        "Lateness Amount",
        "Other Deduction",
        "Deduction",
        "Arrear",
        "Net Pay",
        "TDS",
        "Net pay after TDS"
    ]

    hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    ws.row_dimensions[1].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    sample_rows = [
        [
            "2025-01", "EMP-001", "Tanvir Ahmed", "Finance & Accounts", "Senior Accountant", "2020-01-15",
            50000.0, 20000.0, 5000.0, 3000.0, 78000.0,
            0.08, 4000.0, 4000.0, 8000.0,
            50000.0, 0.0, 1500.0, 3750.0, 1000.0, 500.0, 0.0, 0.0,
            26, 0, 1, 0.0, 250.0, 0.0, 4250.0, 0.0,
            80500.0, 2500.0, 78000.0
        ],
        [
            "2025-01", "EMP-002", "Farhana Yasmin", "Human Resources", "HR Manager", "2018-06-01",
            60000.0, 24000.0, 6000.0, 3500.0, 93500.0,
            0.08, 4800.0, 4800.0, 9600.0,
            60000.0, 0.0, 0.0, 6000.0, 0.0, 0.0, 0.0, 500.0,
            26, 1, 0, 1923.0, 0.0, 0.0, 6723.0, 0.0,
            93277.0, 3500.0, 89777.0
        ],
        [
            "2025-01", "EMP-003", "Md. Rafiqul Islam", "Operations", "Operations Supervisor", "2022-03-10",
            42000.0, 16800.0, 4200.0, 2500.0, 65500.0,
            0.08, 3360.0, 3360.0, 6720.0,
            42000.0, 1200.0, 2000.0, 0.0, 500.0, 0.0, 0.0, 0.0,
            26, 0, 0, 0.0, 0.0, 0.0, 3360.0, 0.0,
            65840.0, 1500.0, 64340.0
        ]
    ]

    for r_idx, row_vals in enumerate(sample_rows, 2):
        ws.row_dimensions[r_idx].height = 20
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if isinstance(val, (int, float)):
                if c_idx in [11, 23, 24, 25]: # Rate or days
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if c_idx == 11:
                        cell.number_format = '0.00%'
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00'
            else:
                cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 6] else "left", vertical="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Historical_Salary_Registry_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==============================================================================
# REPORT: EMPLOYEE CUMULATIVE SALARY & COMPANY TOTAL PAID REPORT
# ==============================================================================

@app.route('/payroll/employee-salary-report')
def employee_salary_report():
    """Comprehensive employee salary statement showing all salary information and company total paid."""
    ensure_salary_history_schema()
    search_id = request.args.get('search_id', '').strip()
    emp_param = request.args.get('employee_id')

    target_emp = None
    query_val = search_id or emp_param
    if query_val:
        q_str = str(query_val).strip()
        # 1. Match by employee_id code first (e.g. '19', 'EMP-001')
        target_emp = Employee.query.filter(Employee.employee_id.ilike(q_str)).first()
        # 2. If not found and is integer, check primary key id
        if not target_emp and q_str.isdigit():
            target_emp = Employee.query.get(int(q_str))
        # 3. If still not found, match by employee name
        if not target_emp:
            target_emp = Employee.query.filter(Employee.name.ilike(f"%{q_str}%")).first()

    employees = Employee.query.order_by(Employee.employee_id.asc()).all()

    histories = []
    totals = {
        'months_count': 0,
        'basic': 0.0,
        'hr': 0.0,
        'medical': 0.0,
        'conveyance': 0.0,
        'gross': 0.0,
        'pf_employer': 0.0,
        'pf_employee': 0.0,
        'pf_total': 0.0,
        'salary_payable': 0.0,
        'govt_holiday': 0.0,
        'holiday_duty': 0.0,
        'loyalty': 0.0,
        'additional_duty': 0.0,
        'office_order': 0.0,
        'eid_bonus': 0.0,
        'other_add': 0.0,
        'total_allowances': 0.0,
        'absenteeism': 0.0,
        'lateness': 0.0,
        'other_ded': 0.0,
        'total_ded': 0.0,
        'arrear': 0.0,
        'net_pay': 0.0,
        'tds': 0.0,
        'net_after_tds': 0.0,
        'company_total_paid': 0.0
    }

    if target_emp:
        histories = EmployeeSalaryHistory.query.filter_by(employee_id=target_emp.id).order_by(EmployeeSalaryHistory.salary_month.asc()).all()
        totals['months_count'] = len(histories)
        for h in histories:
            totals['basic'] += (h.basic_salary or 0.0)
            totals['hr'] += (h.house_rent or 0.0)
            totals['medical'] += (h.medical or 0.0)
            totals['conveyance'] += (h.conveyance or 0.0)
            totals['gross'] += (h.gross_salary or 0.0)

            p_empr = h.pf_employer or h.pf_deduction or 0.0
            p_emp = h.pf_employee or h.pf_deduction or 0.0
            p_tot = h.pf_total or (p_empr + p_emp)
            totals['pf_employer'] += p_empr
            totals['pf_employee'] += p_emp
            totals['pf_total'] += p_tot

            totals['salary_payable'] += (h.salary_payable or 0.0)
            totals['govt_holiday'] += (h.govt_holiday or 0.0)
            totals['holiday_duty'] += (h.holiday_extra_duty or 0.0)
            totals['loyalty'] += (h.loyalty_allowance or 0.0)
            totals['additional_duty'] += (h.additional_duty or 0.0)
            totals['office_order'] += (h.office_order_allowance or 0.0)
            totals['eid_bonus'] += (h.eid_bonus or 0.0)
            totals['other_add'] += (h.other_addition or h.other_allowance or 0.0)

            allw = (h.house_rent or 0.0) + (h.medical or 0.0) + (h.conveyance or 0.0) + \
                   (h.govt_holiday or 0.0) + (h.holiday_extra_duty or 0.0) + (h.loyalty_allowance or 0.0) + \
                   (h.additional_duty or 0.0) + (h.office_order_allowance or 0.0) + (h.eid_bonus or 0.0) + (h.other_addition or h.other_allowance or 0.0)
            totals['total_allowances'] += allw

            totals['absenteeism'] += (h.absenteeism_amount or 0.0)
            totals['lateness'] += (h.lateness_amount or 0.0)
            totals['other_ded'] += (h.other_deduction or 0.0)
            totals['total_ded'] += (h.total_deductions or 0.0)
            totals['arrear'] += (h.arrear or 0.0)
            totals['net_pay'] += (h.net_salary or 0.0)

            t_val = (h.tds or h.tax_deduction or 0.0)
            totals['tds'] += t_val

            net_after = h.net_pay_after_tds if (h.net_pay_after_tds is not None and h.net_pay_after_tds > 0) else (h.net_salary or 0.0)
            totals['net_after_tds'] += net_after

            co_paid = h.display_company_paid
            totals['company_total_paid'] += co_paid
    elif search_id:
        flash(f'No employee found matching ID or Name: "{search_id}". Please check the ID or select from the list.', 'warning')

    return render_template(
        'payroll/employee_salary_report.html',
        target_emp=target_emp,
        search_id=search_id or (target_emp.employee_id if target_emp else ''),
        employees=employees,
        histories=histories,
        totals=totals
    )


@app.route('/payroll/employee-salary-report/export-excel/<int:emp_id>')
def export_employee_salary_report_excel(emp_id):
    """Export complete cumulative salary ledger for an individual employee to Excel."""
    emp = Employee.query.get_or_404(emp_id)
    histories = EmployeeSalaryHistory.query.filter_by(employee_id=emp.id).order_by(EmployeeSalaryHistory.salary_month.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Salary Report - {emp.employee_id}"

    title_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    sec_font = Font(name="Calibri", size=11, bold=True, color="334155")
    val_font = Font(name="Calibri", size=10, bold=False, color="0F172A")
    bold_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    hdr_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    tot_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    co_paid_fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    co_paid_font = Font(name="Calibri", size=10, bold=True, color="065F46")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. Company Banner
    ws.cell(row=1, column=1, value="LEDGERPRO ERP - EMPLOYEE CUMULATIVE SALARY & COMPENSATION STATEMENT").font = title_font
    ws.cell(row=2, column=1, value=f"Generated on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')} | Confidential Payroll Record").font = sec_font

    # 2. Employee Profile Info Block
    ws.cell(row=4, column=1, value="Employee ID:").font = bold_font
    ws.cell(row=4, column=2, value=emp.employee_id).font = val_font
    ws.cell(row=4, column=4, value="Full Name:").font = bold_font
    ws.cell(row=4, column=5, value=emp.name).font = val_font

    ws.cell(row=5, column=1, value="Department:").font = bold_font
    ws.cell(row=5, column=2, value=emp.department).font = val_font
    ws.cell(row=5, column=4, value="Designation:").font = bold_font
    ws.cell(row=5, column=5, value=emp.designation).font = val_font

    ws.cell(row=6, column=1, value="Joining Date:").font = bold_font
    ws.cell(row=6, column=2, value=emp.joining_date.strftime('%d-%b-%Y') if emp.joining_date else "N/A").font = val_font
    ws.cell(row=6, column=4, value="Total Periods Recorded:").font = bold_font
    ws.cell(row=6, column=5, value=f"{len(histories)} Months").font = val_font

    # 3. Headers
    headers = [
        "Month", "Payment Date", "Basic (৳)", "HR Allowance (৳)", "Medical (৳)", "Conveyance (৳)",
        "Gross Pay (৳)", "PF Rate", "PF Employer (৳)", "PF Employee (৳)", "PF Total (৳)",
        "Salary (৳)", "Govt. Holiday (৳)", "Holiday & Extra Duty (৳)", "Loyalty Allowance (৳)",
        "Additional Duty (৳)", "Office Order (৳)", "Eid Bonus (৳)", "Other Additions (৳)",
        "Work Days", "Absent", "Late", "Absent Amount (৳)", "Late Amount (৳)", "Other Deductions (৳)",
        "Total Deductions (৳)", "Arrear (৳)", "Net Pay (৳)", "TDS (৳)", "Net Pay After TDS (৳)",
        "COMPANY TOTAL PAID (৳)"
    ]

    start_r = 8
    ws.row_dimensions[start_r].height = 25
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_r, column=c_idx, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # 4. Data Rows
    for idx, h in enumerate(histories, start_r + 1):
        ws.row_dimensions[idx].height = 18
        p_empr = h.pf_employer or h.pf_deduction or 0.0
        p_emp = h.pf_employee or h.pf_deduction or 0.0
        p_tot = h.pf_total or (p_empr + p_emp)
        net_after = h.net_pay_after_tds if (h.net_pay_after_tds is not None and h.net_pay_after_tds > 0) else (h.net_salary or 0.0)
        co_paid = h.display_company_paid

        row_vals = [
            h.salary_month,
            h.payment_date.strftime('%d-%b-%Y') if h.payment_date else '',
            h.basic_salary or 0.0,
            h.house_rent or 0.0,
            h.medical or 0.0,
            h.conveyance or 0.0,
            h.gross_salary or 0.0,
            h.pf_rate or 0.0,
            p_empr,
            p_emp,
            p_tot,
            h.salary_payable or 0.0,
            h.govt_holiday or 0.0,
            h.holiday_extra_duty or 0.0,
            h.loyalty_allowance or 0.0,
            h.additional_duty or 0.0,
            h.office_order_allowance or 0.0,
            h.eid_bonus or 0.0,
            h.other_addition or h.other_allowance or 0.0,
            h.working_days or 26.0,
            h.absent_days or 0.0,
            h.late_days or 0.0,
            h.absenteeism_amount or 0.0,
            h.lateness_amount or 0.0,
            h.other_deduction or 0.0,
            h.total_deductions or 0.0,
            h.arrear or 0.0,
            h.net_salary or 0.0,
            h.tds or h.tax_deduction or 0.0,
            net_after,
            co_paid
        ]

        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.border = thin_border
            if isinstance(val, (int, float)):
                if c_idx in [8, 20, 21, 22]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if c_idx == 8:
                        cell.number_format = '0.00%'
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00'
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if c_idx == len(headers):
                cell.fill = co_paid_fill
                cell.font = co_paid_font

    # 5. Grand Totals Row
    tot_row = start_r + len(histories) + 1
    ws.row_dimensions[tot_row].height = 22
    ws.cell(row=tot_row, column=1, value="CUMULATIVE TOTALS:").font = bold_font
    ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=tot_row, column=1).fill = tot_fill

    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=tot_row, column=c_idx)
        cell.border = thin_border
        cell.fill = tot_fill
        if c_idx >= 3 and c_idx not in [8, 20, 21, 22]:
            col_letter = get_column_letter(c_idx)
            cell.value = f"=SUM({col_letter}{start_r + 1}:{col_letter}{tot_row - 1})"
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = bold_font
            if c_idx == len(headers):
                cell.fill = co_paid_fill
                cell.font = co_paid_font

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"Salary_Report_{emp.employee_id}_{emp.name.replace(' ', '_')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==============================================================================
# OFFICIAL CERTIFICATES & SLIPS GENERATOR MODULE
# ==============================================================================

@app.route('/employees/documents')
def documents_hub():
    emp_id = request.args.get('employee_id', type=int)
    employees = Employee.query.order_by(Employee.employee_id.asc()).all()
    selected_emp = Employee.query.get(emp_id) if emp_id else (employees[0] if employees else None)

    return render_template(
        'employees/documents_hub.html',
        employees=employees,
        selected_employee=selected_emp
    )


@app.route('/employees/<int:emp_id>/noc')
def generate_noc(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    purpose = request.args.get('purpose', 'Visa Application and International Travel')
    country = request.args.get('country', 'Overseas')
    issue_date_str = request.args.get('issue_date')
    issue_date = parse_accounting_date(issue_date_str) or date.today()
    ref_no = request.args.get('ref_no') or f"LPRO/HR/NOC/{issue_date.year}/{employee.employee_id}"

    return render_template(
        'certificates/noc_certificate.html',
        employee=employee,
        purpose=purpose,
        country=country,
        issue_date=issue_date,
        ref_no=ref_no
    )


@app.route('/employees/<int:emp_id>/salary-certificate')
def generate_salary_certificate(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    purpose = request.args.get('purpose', 'Bank Loan & Credit Card Facility Application')
    issue_date_str = request.args.get('issue_date')
    issue_date = parse_accounting_date(issue_date_str) or date.today()
    ref_no = request.args.get('ref_no') or f"LPRO/FIN/SC/{issue_date.year}/{employee.employee_id}"

    b_sal = employee.basic_salary
    h_rent = employee.house_rent if employee.house_rent > 0 else round(b_sal * 0.4, 2)
    med = employee.medical_allowance if employee.medical_allowance > 0 else round(b_sal * 0.1, 2)
    conv = employee.conveyance if employee.conveyance > 0 else round(b_sal * 0.05, 2)
    gross = b_sal + h_rent + med + conv + employee.allowances
    pf = employee.provident_fund if employee.provident_fund > 0 else round(b_sal * 0.08, 2)
    tot_ded = pf + employee.deductions
    net = gross - tot_ded

    net_words = num_to_words(net)

    return render_template(
        'certificates/salary_certificate.html',
        employee=employee,
        purpose=purpose,
        issue_date=issue_date,
        ref_no=ref_no,
        basic=b_sal,
        house_rent=h_rent,
        medical=med,
        conveyance=conv,
        gross=gross,
        pf_deduction=pf,
        total_deductions=tot_ded,
        net_salary=net,
        net_words=net_words
    )


@app.route('/employees/<int:emp_id>/experience-certificate')
def generate_experience_certificate(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    issue_date_str = request.args.get('issue_date')
    issue_date = parse_accounting_date(issue_date_str) or date.today()
    ref_no = request.args.get('ref_no') or f"LPRO/HR/EXP/{issue_date.year}/{employee.employee_id}"

    j_date = employee.joining_date
    end_date = employee.release_date or date.today()

    return render_template(
        'certificates/experience_certificate.html',
        employee=employee,
        issue_date=issue_date,
        ref_no=ref_no,
        joining_date=j_date,
        end_date=end_date
    )


@app.route('/employees/<int:emp_id>/payslip-generator')
def universal_payslip_generator(emp_id):
    employee = Employee.query.get_or_404(emp_id)
    month = request.args.get('month', '2026-01')

    # 1. Check if there is an official payroll run item
    payroll_item = PayrollItem.query.join(PayrollRecord).filter(
        PayrollItem.employee_id == emp_id,
        PayrollRecord.payroll_month == month
    ).first()

    # 2. Check if there is a salary history record
    salary_hist = EmployeeSalaryHistory.query.filter_by(employee_id=emp_id, salary_month=month).first()

    # Available months for this employee
    hist_months = [h.salary_month for h in EmployeeSalaryHistory.query.filter_by(employee_id=emp_id).all()]
    run_months = [r.payroll_month for r in PayrollRecord.query.all()]
    all_months = sorted(list(set(hist_months + run_months + ['2026-01', '2026-02', '2026-03'])), reverse=True)

    b_sal = employee.basic_salary
    h_rent = employee.house_rent if employee.house_rent > 0 else round(b_sal * 0.4, 2)
    med = employee.medical_allowance if employee.medical_allowance > 0 else round(b_sal * 0.1, 2)
    conv = employee.conveyance if employee.conveyance > 0 else round(b_sal * 0.05, 2)
    pf = employee.provident_fund if employee.provident_fund > 0 else round(b_sal * 0.08, 2)

    if payroll_item:
        b_sal = payroll_item.basic_salary
        allw = payroll_item.allowances
        ded = payroll_item.deductions
        net = payroll_item.net_salary
        p_date = payroll_item.payroll_record.processed_date
    elif salary_hist:
        b_sal = salary_hist.basic_salary
        h_rent = salary_hist.house_rent
        med = salary_hist.medical
        conv = salary_hist.conveyance
        allw = salary_hist.other_allowance + h_rent + med + conv
        ded = salary_hist.total_deductions
        pf = salary_hist.pf_deduction
        net = salary_hist.net_salary
        p_date = salary_hist.payment_date
    else:
        allw = h_rent + med + conv
        ded = pf + employee.deductions
        net = (b_sal + allw) - ded
        p_date = date.today()

    gross = b_sal + allw
    net_words = num_to_words(net)

    return render_template(
        'certificates/universal_payslip.html',
        employee=employee,
        month=month,
        all_months=all_months,
        basic=b_sal,
        house_rent=h_rent,
        medical=med,
        conveyance=conv,
        allowances=allw,
        gross=gross,
        pf_deduction=pf,
        deductions=ded,
        net_salary=net,
        net_words=net_words,
        payment_date=p_date
    )


# ==============================================================================
# ROUTES: CONFIGURATION, MULTI-COMPANY & DATA TRANSFER
# ==============================================================================

@app.route('/configuration')
@login_required
@admin_required
def configuration_hub():
    all_companies = CompanyManager.get_all()
    active_cid = session.get('active_company_id', 1)
    active_company = CompanyManager.get_by_id(active_cid)
    smtp_cfg = get_smtp_config()
    return render_template(
        'configuration/index.html',
        all_companies=all_companies,
        active_company=active_company,
        smtp_cfg=smtp_cfg
    )

@app.route('/configuration/companies/new', methods=['POST'])
@login_required
@admin_required
def create_company_profile():
    name = request.form.get('name', '').strip()
    code = request.form.get('code', '').strip().upper()
    description = request.form.get('description', '').strip()
    is_locked = request.form.get('is_locked') == 'yes'
    password = request.form.get('password', '').strip()
    init_setup = request.form.get('init_setup', 'blank')
    clone_source_id = request.form.get('clone_source_id')

    if not name or not code:
        flash('Company Name and Company Code are required.', 'danger')
        return redirect(url_for('configuration_hub'))

    clone_from = int(clone_source_id) if (init_setup == 'clone' and clone_source_id) else None

    try:
        new_comp = CompanyManager.create_company(
            name=name,
            code=code,
            description=description,
            is_locked=is_locked,
            password=password,
            metadata=db.metadata,
            clone_from_id=clone_from,
            clone_mode='full'
        )
        lock_msg = "🔒 Password Protection Active" if is_locked else "🔓 Open Access"
        flash(f'Company Profile "{new_comp.name}" ({new_comp.code}) created successfully! [{lock_msg}]', 'success')
    except Exception as e:
        flash(f'Error creating company profile: {str(e)}', 'danger')

    return redirect(url_for('configuration_hub'))

@app.route('/configuration/companies/<int:cid>/edit', methods=['POST'])
@login_required
@admin_required
def edit_company(cid):
    name = request.form.get('name', '').strip()
    code = request.form.get('code', '').strip().upper()
    description = request.form.get('description', '').strip()
    is_locked = request.form.get('is_locked') == 'yes'
    password = request.form.get('password', '').strip()

    if not name or not code:
        flash('Company Name and Company Code are required.', 'danger')
        return redirect(url_for('configuration_hub'))

    try:
        updated = CompanyManager.update_company(
            cid=cid,
            name=name,
            code=code,
            description=description,
            is_locked=is_locked,
            new_password=password
        )
        lock_msg = "🔒 Password Protection Enabled" if is_locked else "🔓 Open Access"
        flash(f'Company Profile "{updated.name}" updated successfully. [{lock_msg}]', 'success')
    except Exception as e:
        flash(f'Error updating company profile: {str(e)}', 'danger')

    return redirect(url_for('configuration_hub'))

@app.route('/configuration/companies/<int:cid>/delete', methods=['POST'])
@login_required
@admin_required
def delete_company(cid):
    try:
        comp = CompanyManager.get_by_id(cid)
        comp_name = comp.name if comp else f'ID {cid}'
        if CompanyManager.delete_company(cid):
            if session.get('active_company_id') == cid:
                session['active_company_id'] = 1
            flash(f'Company Profile "{comp_name}" and its isolated database were deleted.', 'info')
        else:
            flash('Company profile could not be found or deleted.', 'danger')
    except Exception as e:
        flash(f'Error deleting company: {str(e)}', 'danger')

    return redirect(url_for('configuration_hub'))

@app.route('/configuration/companies/<int:cid>/switch')
@login_required
def switch_company(cid):
    comp = CompanyManager.get_by_id(cid)
    if not comp:
        flash('Selected company profile does not exist.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))

    # Check password lock
    unlocked = session.get('unlocked_companies', [])
    if comp.is_locked and comp.id not in unlocked:
        # Prompt for profile password
        return redirect(url_for('unlock_company', cid=comp.id))

    session['active_company_id'] = comp.id
    flash(f'Active company switched to "{comp.name}" ({comp.code}). All ledgers and payroll now route to this entity.', 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/configuration/companies/<int:cid>/unlock', methods=['GET', 'POST'])
@login_required
def unlock_company(cid):
    comp = CompanyManager.get_by_id(cid)
    if not comp:
        flash('Selected company profile does not exist.', 'danger')
        return redirect(url_for('configuration_hub'))

    if not comp.is_locked:
        session['active_company_id'] = comp.id
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        entered_password = request.form.get('company_password', '')
        if comp.check_password(entered_password):
            unlocked = session.setdefault('unlocked_companies', [])
            if comp.id not in unlocked:
                unlocked.append(comp.id)
            session['active_company_id'] = comp.id
            flash(f'Company Profile "{comp.name}" unlocked successfully.', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Incorrect profile password. Access denied.', 'danger')

    return render_template('configuration/unlock_company.html', company=comp)

@app.route('/configuration/transfer-data', methods=['POST'])
@login_required
@admin_required
def transfer_company_data():
    source_cid = request.form.get('source_cid', type=int)
    dest_cid = request.form.get('dest_cid', type=int)
    mode = request.form.get('transfer_mode', 'full')
    selected_modules = request.form.getlist('modules')
    strategy = request.form.get('strategy', 'replace')

    if not source_cid or not dest_cid:
        flash('Please select both source and destination companies.', 'danger')
        return redirect(url_for('configuration_hub'))

    if source_cid == dest_cid:
        flash('Source and destination companies must be different.', 'danger')
        return redirect(url_for('configuration_hub'))

    try:
        result = CompanyManager.transfer_data(
            source_cid=source_cid,
            dest_cid=dest_cid,
            mode=mode,
            selected_modules=selected_modules,
            strategy=strategy,
            metadata=db.metadata
        )
        if result['success']:
            transferred_summary = ", ".join(f"{tbl}: {cnt} rows" for tbl, cnt in result['row_counts'].items() if cnt > 0)
            flash(f"Data transfer successful from '{result['source']}' to '{result['destination']}'! [{transferred_summary or 'Completed'}]", 'success')
        else:
            flash(result.get('message', 'Transfer failed.'), 'warning')
    except Exception as e:
        flash(f'Data transfer error: {str(e)}', 'danger')

    return redirect(url_for('configuration_hub'))

@app.route('/configuration/smtp', methods=['POST'])
@login_required
@admin_required
def save_smtp_settings():
    action = request.form.get('action', 'save')
    smtp_server = request.form.get('smtp_server', 'smtp.gmail.com').strip()
    smtp_port = int(request.form.get('smtp_port') or 587)
    smtp_user = request.form.get('smtp_user', ADMIN_EMAIL).strip()
    smtp_password = request.form.get('smtp_password', '').strip()
    sender_name = request.form.get('sender_name', 'ZR-ERP Security System').strip()
    use_tls = request.form.get('use_tls') == 'yes'

    cfg = get_smtp_config()
    cfg['smtp_server'] = smtp_server
    cfg['smtp_port'] = smtp_port
    cfg['smtp_user'] = smtp_user
    if smtp_password:
        cfg['smtp_password'] = smtp_password
    cfg['sender_name'] = sender_name
    cfg['use_tls'] = use_tls
    save_smtp_config(cfg)

    if action == 'test':
        test_otp = generate_otp()
        res = send_otp_email(ADMIN_EMAIL, test_otp)
        if res.get('smtp_dispatched'):
            flash(f'Test OTP ({test_otp}) sent successfully via Gmail SMTP to {ADMIN_EMAIL}!', 'success')
        else:
            flash(f'SMTP Test Result: Code generated ({test_otp}) for {ADMIN_EMAIL}. Note: Set Gmail App Password to enable direct cloud dispatch. Error details: {res.get("error") or "No password set"}', 'info')
    else:
        flash('SMTP configuration saved successfully.', 'success')

    return redirect(url_for('configuration_hub'))


# ==============================================================================
# BUDGET AND EXPENSE MANAGEMENT MODULE
# ==============================================================================

import json

@app.route('/budgets')
@login_required
def budget_dashboard():
    # 1. Determine requested month
    req_month = request.args.get('month', '').strip()
    
    # All available budgets
    all_budgets = MonthlyBudget.query.order_by(MonthlyBudget.month_year.desc()).all()
    available_months = [b.month_year for b in all_budgets]
    
    current_month_str = datetime.utcnow().strftime('%Y-%m')
    if not req_month:
        if current_month_str in available_months:
            req_month = current_month_str
        elif available_months:
            req_month = available_months[0]
        else:
            req_month = current_month_str
            
    # Active budget for this month
    budget = MonthlyBudget.query.filter_by(month_year=req_month).order_by(MonthlyBudget.id.desc()).first()
    
    # 2. Query all expense accounts
    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.category, Account.code).all()
    
    # 3. Query actual expenses incurred in req_month
    actual_rows = db.session.query(
        TransactionLine.account_id,
        db.func.sum(TransactionLine.debit - TransactionLine.credit).label('spent')
    ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
    ).filter(
        db.func.strftime('%Y-%m', JournalEntry.date) == req_month
    ).group_by(TransactionLine.account_id).all()
    
    actual_spent_map = {r.account_id: float(r.spent or 0.0) for r in actual_rows}
    budget_item_map = {item.account_id: item.allocated_amount for item in budget.items} if budget else {}
    
    categories = ['Operating Expense', 'Administrative Expense', 'Financial Expense']
    items_by_cat = {cat: [] for cat in categories}
    items_by_cat['Other Expense'] = []
    
    total_budgeted = 0.0
    total_actual = 0.0
    count_safe = 0
    count_warning = 0
    count_danger = 0
    
    comparison_list = []
    
    for acc in expense_accounts:
        b_amt = float(budget_item_map.get(acc.id, 0.0))
        a_amt = float(actual_spent_map.get(acc.id, 0.0))
        
        variance = b_amt - a_amt  # Positive = Saved / Remaining, Negative = Over budget
        
        if b_amt > 0:
            utilization = round((a_amt / b_amt) * 100, 1)
        elif a_amt > 0:
            utilization = 999.0
        else:
            utilization = 0.0
            
        if b_amt == 0 and a_amt == 0:
            status = 'inactive'
            status_label = 'Not Budgeted'
            status_class = 'badge-secondary'
        elif (a_amt > b_amt and b_amt > 0) or (b_amt == 0 and a_amt > 0):
            status = 'danger'
            status_label = 'Over Budget' if b_amt > 0 else 'Unbudgeted'
            status_class = 'badge-danger'
            count_danger += 1
        elif utilization >= 85.0:
            status = 'warning'
            status_label = 'Near Limit'
            status_class = 'badge-warning'
            count_warning += 1
        else:
            status = 'safe'
            status_label = 'On Track'
            status_class = 'badge-success'
            count_safe += 1
            
        total_budgeted += b_amt
        total_actual += a_amt
        
        item_obj = {
            'account_id': acc.id,
            'code': acc.code,
            'name': acc.name,
            'category': acc.category,
            'budgeted': b_amt,
            'actual': a_amt,
            'variance': variance,
            'utilization': min(utilization, 999.0),
            'status': status,
            'status_label': status_label,
            'status_class': status_class
        }
        
        target_cat = acc.category if acc.category in items_by_cat else 'Other Expense'
        items_by_cat[target_cat].append(item_obj)
        comparison_list.append(item_obj)
        
    total_variance = total_budgeted - total_actual
    overall_utilization = round((total_actual / total_budgeted * 100), 1) if total_budgeted > 0 else (100.0 if total_actual > 0 else 0.0)
    
    # 5. Category Rollups
    cat_summaries = []
    for cat_name, cat_items in items_by_cat.items():
        if not cat_items:
            continue
        c_bud = sum(i['budgeted'] for i in cat_items)
        c_act = sum(i['actual'] for i in cat_items)
        c_var = c_bud - c_act
        c_util = round((c_act / c_bud * 100), 1) if c_bud > 0 else 0.0
        cat_summaries.append({
            'name': cat_name,
            'budgeted': c_bud,
            'actual': c_act,
            'variance': c_var,
            'utilization': c_util,
            'item_count': len(cat_items)
        })
        
    # 6. Chart.js datasets:
    # A. Top 8 cost centers comparison (Budget vs Actual)
    active_items = [i for i in comparison_list if i['budgeted'] > 0 or i['actual'] > 0]
    sorted_by_impact = sorted(active_items, key=lambda x: max(x['budgeted'], x['actual']), reverse=True)[:8]
    bar_chart_labels = [i['name'][:20] for i in sorted_by_impact]
    bar_chart_budgets = [i['budgeted'] for i in sorted_by_impact]
    bar_chart_actuals = [i['actual'] for i in sorted_by_impact]
    
    # B. Multi-Month Trendline (up to last 8 months)
    months_query = db.session.query(
        db.func.strftime('%Y-%m', JournalEntry.date).label('m')
    ).group_by('m').order_by('m').all()
    
    trend_months = [r.m for r in months_query if r.m]
    if req_month not in trend_months:
        trend_months.append(req_month)
    trend_months = sorted(list(set(trend_months)))[-8:]
    
    trend_labels = []
    trend_budget_vals = []
    trend_actual_vals = []
    
    for tm in trend_months:
        try:
            p = tm.split('-')
            trend_labels.append(datetime(int(p[0]), int(p[1]), 1).strftime('%b %y'))
        except Exception:
            trend_labels.append(tm)
            
        mb = MonthlyBudget.query.filter_by(month_year=tm).first()
        b_val = mb.total_budgeted if mb else 0.0
        
        act_val = db.session.query(
            db.func.sum(TransactionLine.debit - TransactionLine.credit)
        ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
        ).join(Account, TransactionLine.account_id == Account.id
        ).filter(
            Account.type == 'Expense',
            db.func.strftime('%Y-%m', JournalEntry.date) == tm
        ).scalar() or 0.0
        
        trend_budget_vals.append(float(b_val))
        trend_actual_vals.append(float(act_val))
        
    # Month display name
    try:
        p = req_month.split('-')
        month_name = datetime(int(p[0]), int(p[1]), 1).strftime('%B %Y')
    except Exception:
        month_name = req_month

    # Parse structured corporate budget sheet if available
    budget_sheet = None
    if budget and budget.budget_data_json:
        try:
            budget_sheet = json.loads(budget.budget_data_json)
        except Exception:
            budget_sheet = None

    return render_template(
        'budget/dashboard.html',
        month=req_month,
        month_name=month_name,
        budget=budget,
        budget_sheet=budget_sheet,
        all_budgets=all_budgets,
        available_months=available_months,
        items_by_cat=items_by_cat,
        comparison_list=comparison_list,
        cat_summaries=cat_summaries,
        total_budgeted=total_budgeted,
        total_actual=total_actual,
        total_variance=total_variance,
        overall_utilization=overall_utilization,
        count_safe=count_safe,
        count_warning=count_warning,
        count_danger=count_danger,
        bar_chart_labels=json.dumps(bar_chart_labels),
        bar_chart_budgets=json.dumps(bar_chart_budgets),
        bar_chart_actuals=json.dumps(bar_chart_actuals),
        trend_labels=json.dumps(trend_labels),
        trend_budget_vals=json.dumps(trend_budget_vals),
        trend_actual_vals=json.dumps(trend_actual_vals)
    )

@app.route('/budgets/issue', methods=['GET', 'POST'])
@login_required
def issue_budget():
    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.category, Account.code).all()
    
    today = date.today()
    default_month = request.args.get('month', today.strftime('%Y-%m'))
    
    clone_from_id = request.args.get('clone_from', type=int)
    cloned_allocations = {}
    if clone_from_id:
        c_budget = MonthlyBudget.query.get(clone_from_id)
        if c_budget:
            cloned_allocations = {item.account_id: item.allocated_amount for item in c_budget.items}
            
    # Calculate benchmark from previous month's actual expenses
    try:
        p = default_month.split('-')
        cur_dt = datetime(int(p[0]), int(p[1]), 1)
        prev_dt = cur_dt - timedelta(days=1)
        prev_month_str = prev_dt.strftime('%Y-%m')
    except Exception:
        prev_month_str = '2026-08'
        
    bench_rows = db.session.query(
        TransactionLine.account_id,
        db.func.sum(TransactionLine.debit - TransactionLine.credit).label('spent')
    ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
    ).filter(
        db.func.strftime('%Y-%m', JournalEntry.date) == prev_month_str
    ).group_by(TransactionLine.account_id).all()
    benchmark_map = {r.account_id: float(r.spent or 0.0) for r in bench_rows}
    
    if request.method == 'POST':
        month_year = request.form.get('month_year', '').strip()
        title = request.form.get('title', '').strip()
        fiscal_year = request.form.get('fiscal_year', '').strip()
        status = request.form.get('status', 'Issued')
        notes = request.form.get('notes', '').strip()
        
        if not month_year:
            flash('Please specify a valid Budget Month (YYYY-MM).', 'danger')
            return redirect(url_for('issue_budget'))
            
        if not title:
            try:
                p = month_year.split('-')
                m_name = datetime(int(p[0]), int(p[1]), 1).strftime('%B %Y')
                title = f'Monthly Operating Budget - {m_name}'
            except Exception:
                title = f'Monthly Operating Budget - {month_year}'
                
        user_obj = User.query.get(session.get('user_id')) if 'user_id' in session else None
        issuer_name = user_obj.username if user_obj else 'Admin'

        existing = MonthlyBudget.query.filter_by(month_year=month_year).first()
        if existing:
            budget = existing
            budget.title = title
            budget.fiscal_year = fiscal_year
            budget.status = status
            budget.notes = notes
            budget.issued_by = issuer_name
            budget.updated_at = datetime.utcnow()
            BudgetItem.query.filter_by(budget_id=budget.id).delete()
        else:
            budget = MonthlyBudget(
                month_year=month_year,
                title=title,
                fiscal_year=fiscal_year,
                status=status,
                notes=notes,
                issued_by=issuer_name,
                total_budgeted=0.0
            )
            db.session.add(budget)
            db.session.flush()
            
        total_budgeted = 0.0
        for acc in expense_accounts:
            val_str = request.form.get(f'allocated_{acc.id}', '0').strip().replace(',', '')
            try:
                alloc_val = float(val_str)
            except ValueError:
                alloc_val = 0.0
                
            if alloc_val > 0:
                item = BudgetItem(
                    budget_id=budget.id,
                    account_id=acc.id,
                    allocated_amount=alloc_val,
                    notes=request.form.get(f'notes_{acc.id}', '')
                )
                db.session.add(item)
                total_budgeted += alloc_val
                
        budget.total_budgeted = total_budgeted
        db.session.commit()
        
        flash(f'Monthly Budget for {budget.month_display} issued successfully with total allocation of ৳{total_budgeted:,.2f}!', 'success')
        return redirect(url_for('budget_dashboard', month=month_year))
        
    return render_template(
        'budget/issue.html',
        expense_accounts=expense_accounts,
        default_month=default_month,
        cloned_allocations=cloned_allocations,
        benchmark_map=benchmark_map,
        prev_month_str=prev_month_str,
        existing_budgets=MonthlyBudget.query.order_by(MonthlyBudget.month_year.desc()).all()
    )

@app.route('/budgets/<int:budget_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_budget(budget_id):
    budget = MonthlyBudget.query.get_or_404(budget_id)
    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.category, Account.code).all()
    existing_items = {item.account_id: item.allocated_amount for item in budget.items}
    
    if request.method == 'POST':
        budget.title = request.form.get('title', budget.title).strip()
        budget.status = request.form.get('status', budget.status)
        budget.notes = request.form.get('notes', budget.notes).strip()
        budget.fiscal_year = request.form.get('fiscal_year', budget.fiscal_year).strip()
        
        BudgetItem.query.filter_by(budget_id=budget.id).delete()
        total_budgeted = 0.0
        for acc in expense_accounts:
            val_str = request.form.get(f'allocated_{acc.id}', '0').strip().replace(',', '')
            try:
                alloc_val = float(val_str)
            except ValueError:
                alloc_val = 0.0
            if alloc_val > 0:
                item = BudgetItem(
                    budget_id=budget.id,
                    account_id=acc.id,
                    allocated_amount=alloc_val
                )
                db.session.add(item)
                total_budgeted += alloc_val
                
        budget.total_budgeted = total_budgeted
        budget.updated_at = datetime.utcnow()
        db.session.commit()
        
        flash(f'Monthly Budget "{budget.title}" updated successfully!', 'success')
        return redirect(url_for('budget_dashboard', month=budget.month_year))
        
    return render_template(
        'budget/edit.html',
        budget=budget,
        expense_accounts=expense_accounts,
        existing_items=existing_items
    )

@app.route('/budgets/<int:budget_id>/delete', methods=['POST'])
@login_required
def delete_budget(budget_id):
    budget = MonthlyBudget.query.get_or_404(budget_id)
    month_name = budget.month_display
    db.session.delete(budget)
    db.session.commit()
    flash(f'Monthly Budget for {month_name} was deleted successfully.', 'success')
    return redirect(url_for('budget_list'))

@app.route('/budgets/list')
@login_required
def budget_list():
    budgets = MonthlyBudget.query.order_by(MonthlyBudget.month_year.desc()).all()
    budget_summaries = []
    for b in budgets:
        act_spent = db.session.query(
            db.func.sum(TransactionLine.debit - TransactionLine.credit)
        ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
        ).join(Account, TransactionLine.account_id == Account.id
        ).filter(
            Account.type == 'Expense',
            db.func.strftime('%Y-%m', JournalEntry.date) == b.month_year
        ).scalar() or 0.0
        
        var = b.total_budgeted - act_spent
        util = round((act_spent / b.total_budgeted * 100), 1) if b.total_budgeted > 0 else 0.0
        budget_summaries.append({
            'budget': b,
            'actual_spent': act_spent,
            'variance': var,
            'utilization': util
        })
    return render_template('budget/list.html', budget_summaries=budget_summaries)

@app.route('/budgets/report')
@login_required
def budget_report():
    req_month = request.args.get('month', '').strip()
    if not req_month:
        latest = MonthlyBudget.query.order_by(MonthlyBudget.month_year.desc()).first()
        req_month = latest.month_year if latest else datetime.utcnow().strftime('%Y-%m')
        
    budget = MonthlyBudget.query.filter_by(month_year=req_month).order_by(MonthlyBudget.id.desc()).first()
    expense_accounts = Account.query.filter_by(type='Expense').order_by(Account.category, Account.code).all()
    
    actual_rows = db.session.query(
        TransactionLine.account_id,
        db.func.sum(TransactionLine.debit - TransactionLine.credit).label('spent')
    ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
    ).filter(
        db.func.strftime('%Y-%m', JournalEntry.date) == req_month
    ).group_by(TransactionLine.account_id).all()
    actual_spent_map = {r.account_id: float(r.spent or 0.0) for r in actual_rows}
    budget_item_map = {item.account_id: item.allocated_amount for item in budget.items} if budget else {}
    
    categories = ['Operating Expense', 'Administrative Expense', 'Financial Expense']
    items_by_cat = {cat: [] for cat in categories}
    items_by_cat['Other Expense'] = []
    
    total_budgeted = 0.0
    total_actual = 0.0
    
    for acc in expense_accounts:
        b_amt = float(budget_item_map.get(acc.id, 0.0))
        a_amt = float(actual_spent_map.get(acc.id, 0.0))
        if b_amt == 0 and a_amt == 0:
            continue
        variance = b_amt - a_amt
        util = round((a_amt / b_amt * 100), 1) if b_amt > 0 else 0.0
        
        target_cat = acc.category if acc.category in items_by_cat else 'Other Expense'
        items_by_cat[target_cat].append({
            'code': acc.code,
            'name': acc.name,
            'budgeted': b_amt,
            'actual': a_amt,
            'variance': variance,
            'utilization': util
        })
        total_budgeted += b_amt
        total_actual += a_amt
        
    total_variance = total_budgeted - total_actual
    overall_utilization = round((total_actual / total_budgeted * 100), 1) if total_budgeted > 0 else 0.0
    
    try:
        p = req_month.split('-')
        month_name = datetime(int(p[0]), int(p[1]), 1).strftime('%B %Y')
    except Exception:
        month_name = req_month
        
    budget_sheet = None
    if budget and budget.budget_data_json:
        try:
            budget_sheet = json.loads(budget.budget_data_json)
        except Exception:
            budget_sheet = None

    return render_template(
        'budget/report.html',
        month=req_month,
        month_name=month_name,
        budget=budget,
        budget_sheet=budget_sheet,
        items_by_cat=items_by_cat,
        total_budgeted=total_budgeted,
        total_actual=total_actual,
        total_variance=total_variance,
        overall_utilization=overall_utilization,
        generated_at=datetime.utcnow()
    )

@app.route('/budgets/api/prefill')
@login_required
def api_budget_prefill():
    mode = request.args.get('mode', 'prev_actual')
    ref_month = request.args.get('ref_month', '2026-08')
    
    allocations = {}
    if mode == 'budget':
        mb = MonthlyBudget.query.filter_by(month_year=ref_month).first()
        if mb:
            allocations = {item.account_id: item.allocated_amount for item in mb.items}
    else:
        rows = db.session.query(
            TransactionLine.account_id,
            db.func.sum(TransactionLine.debit - TransactionLine.credit).label('spent')
        ).join(JournalEntry, TransactionLine.journal_entry_id == JournalEntry.id
        ).filter(
            db.func.strftime('%Y-%m', JournalEntry.date) == ref_month
        ).group_by(TransactionLine.account_id).all()
        allocations = {r.account_id: float(r.spent or 0.0) for r in rows if (r.spent or 0) > 0}
        
    return json.dumps({'success': True, 'allocations': allocations})


# ==============================================================================
# REMITTANCE, INCENTIVE & BUDGET STATEMENT MODULE
# ==============================================================================

@app.route('/remittances')
@login_required
def remittance_dashboard():
    # 1. Query records ordered chronologically
    records = RemittanceRecord.query.order_by(RemittanceRecord.date.asc(), RemittanceRecord.id.asc()).all()
    
    # 2. Recalculate dynamic running balance
    running_bal = 0.0
    for r in records:
        if r.budget_amount > 0:
            running_bal -= r.budget_amount
        if r.received_bdt > 0:
            running_bal += r.received_bdt
        r.calculated_balance = running_bal

    # 3. Macro KPI aggregations
    total_budget_issued = sum(r.budget_amount for r in records)
    total_received_bdt = sum(r.received_bdt for r in records)
    total_remittance_usd = sum(r.remittance_usd for r in records)
    total_incentive_claimed = sum(r.incentive_10_bdt for r in records)
    total_incentive_audited = sum(r.incentive_10_audited_bdt for r in records)
    total_benefit_received = sum(r.benefit_received for r in records)
    total_benefit_receivable = sum(r.benefit_receivable for r in records)
    pending_pipeline = sum(r.incentive_10_bdt for r in records if r.status != 'Clear & Expense in Dec-2023' and r.status != 'Declined')
    
    # Checkpoint completion stats
    foreign_remittances = [r for r in records if r.remittance_usd > 0]
    total_foreign_count = len(foreign_remittances)
    completed_bb_count = len([r for r in foreign_remittances if r.send_to_bangladesh_bank])
    cleared_count = len([r for r in foreign_remittances if 'Clear' in (r.status or '')])

    return render_template(
        'remittance/dashboard.html',
        records=records,
        total_budget_issued=total_budget_issued,
        total_received_bdt=total_received_bdt,
        total_remittance_usd=total_remittance_usd,
        total_incentive_claimed=total_incentive_claimed,
        total_incentive_audited=total_incentive_audited,
        total_benefit_received=total_benefit_received,
        total_benefit_receivable=total_benefit_receivable,
        pending_pipeline=pending_pipeline,
        total_foreign_count=total_foreign_count,
        completed_bb_count=completed_bb_count,
        cleared_count=cleared_count
    )

@app.route('/remittances/create', methods=['GET', 'POST'])
@login_required
def create_remittance_entry():
    if request.method == 'POST':
        date_str = request.form.get('date', '')
        entry_date = datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else date.today()
        
        particulars = request.form.get('particulars', '').strip()
        deal_no = request.form.get('deal_no', '').strip()
        invoice_no = request.form.get('invoice_no', '').strip()
        sent_by = request.form.get('sent_by', '').strip()
        
        budget_amt = float(request.form.get('budget_amount', '0').replace(',', '') or 0.0)
        received_bdt = float(request.form.get('received_bdt', '0').replace(',', '') or 0.0)
        usd_rate = float(request.form.get('usd_rate', '0').replace(',', '') or 0.0)
        remittance_usd = float(request.form.get('remittance_usd', '0').replace(',', '') or 0.0)
        
        if remittance_usd > 0 and usd_rate > 0 and received_bdt == 0:
            received_bdt = round(remittance_usd * usd_rate, 2)
            
        rem_date_str = request.form.get('remittance_date', '')
        remittance_date = datetime.strptime(rem_date_str, '%Y-%m-%d').date() if rem_date_str else None
        
        incentive_10 = float(request.form.get('incentive_10_bdt', '0').replace(',', '') or 0.0)
        if incentive_10 == 0 and received_bdt > 0 and remittance_usd > 0:
            incentive_10 = round(received_bdt * 0.10, 2)
            
        audited_10 = float(request.form.get('incentive_10_audited_bdt', '0').replace(',', '') or 0.0)
        
        tt_cert = request.form.get('tt_certificate') == 'on'
        encash_cert = request.form.get('encashment_certificate') == 'on'
        ad_cert = request.form.get('authorized_dealer_cert') == 'on'
        cash_inc_cert = request.form.get('cash_incentive_cert') == 'on'
        send_to_bank = request.form.get('send_to_bank') == 'on'
        
        bank_date_str = request.form.get('date_send_to_bank', '')
        date_send_to_bank = datetime.strptime(bank_date_str, '%Y-%m-%d').date() if bank_date_str else None
        
        send_to_head_office = request.form.get('send_to_head_office') == 'on'
        send_to_audit_office = request.form.get('send_to_audit_office') == 'on'
        
        audit_date_str = request.form.get('audit_certificate_date', '')
        audit_cert_date = datetime.strptime(audit_date_str, '%Y-%m-%d').date() if audit_date_str else None
        
        send_to_intl_div = request.form.get('send_to_intl_division') == 'on'
        send_to_bb = request.form.get('send_to_bangladesh_bank') == 'on'
        
        benefit_receivable = float(request.form.get('benefit_receivable', '0').replace(',', '') or 0.0)
        benefit_received = float(request.form.get('benefit_received', '0').replace(',', '') or 0.0)
        
        recv_date_str = request.form.get('benefit_received_date', '')
        benefit_recv_date = datetime.strptime(recv_date_str, '%Y-%m-%d').date() if recv_date_str else None
        
        status = request.form.get('status', 'Pending')
        notes = request.form.get('notes', '').strip()
        
        rec = RemittanceRecord(
            date=entry_date,
            particulars=particulars,
            budget_amount=budget_amt,
            received_bdt=received_bdt,
            deal_no=deal_no,
            invoice_no=invoice_no,
            usd_rate=usd_rate,
            remittance_usd=remittance_usd,
            sent_by=sent_by,
            remittance_date=remittance_date,
            incentive_10_bdt=incentive_10,
            incentive_10_audited_bdt=audited_10,
            tt_certificate=tt_cert,
            encashment_certificate=encash_cert,
            authorized_dealer_cert=ad_cert,
            cash_incentive_cert=cash_inc_cert,
            send_to_bank=send_to_bank,
            date_send_to_bank=date_send_to_bank,
            send_to_head_office=send_to_head_office,
            send_to_audit_office=send_to_audit_office,
            audit_certificate_date=audit_cert_date,
            send_to_intl_division=send_to_intl_div,
            send_to_bangladesh_bank=send_to_bb,
            benefit_receivable=benefit_receivable,
            benefit_received=benefit_received,
            benefit_received_date=benefit_recv_date,
            status=status,
            notes=notes
        )
        db.session.add(rec)
        db.session.commit()
        
        flash(f'Record for "{particulars}" logged successfully in Remittance & Incentive Statement!', 'success')
        return redirect(url_for('remittance_dashboard'))
        
    return render_template('remittance/create.html', today=date.today())

@app.route('/remittances/<int:rec_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_remittance_entry(rec_id):
    rec = RemittanceRecord.query.get_or_404(rec_id)
    if request.method == 'POST':
        date_str = request.form.get('date', '')
        if date_str:
            rec.date = datetime.strptime(date_str, '%Y-%m-%d').date()
            
        rec.particulars = request.form.get('particulars', rec.particulars).strip()
        rec.deal_no = request.form.get('deal_no', rec.deal_no).strip()
        rec.invoice_no = request.form.get('invoice_no', rec.invoice_no).strip()
        rec.sent_by = request.form.get('sent_by', rec.sent_by).strip()
        
        rec.budget_amount = float(request.form.get('budget_amount', '0').replace(',', '') or 0.0)
        rec.received_bdt = float(request.form.get('received_bdt', '0').replace(',', '') or 0.0)
        rec.usd_rate = float(request.form.get('usd_rate', '0').replace(',', '') or 0.0)
        rec.remittance_usd = float(request.form.get('remittance_usd', '0').replace(',', '') or 0.0)
        
        rem_date_str = request.form.get('remittance_date', '')
        rec.remittance_date = datetime.strptime(rem_date_str, '%Y-%m-%d').date() if rem_date_str else None
        
        rec.incentive_10_bdt = float(request.form.get('incentive_10_bdt', '0').replace(',', '') or 0.0)
        rec.incentive_10_audited_bdt = float(request.form.get('incentive_10_audited_bdt', '0').replace(',', '') or 0.0)
        
        rec.tt_certificate = request.form.get('tt_certificate') == 'on'
        rec.encashment_certificate = request.form.get('encashment_certificate') == 'on'
        rec.authorized_dealer_cert = request.form.get('authorized_dealer_cert') == 'on'
        rec.cash_incentive_cert = request.form.get('cash_incentive_cert') == 'on'
        rec.send_to_bank = request.form.get('send_to_bank') == 'on'
        
        bank_date_str = request.form.get('date_send_to_bank', '')
        rec.date_send_to_bank = datetime.strptime(bank_date_str, '%Y-%m-%d').date() if bank_date_str else None
        
        rec.send_to_head_office = request.form.get('send_to_head_office') == 'on'
        rec.send_to_audit_office = request.form.get('send_to_audit_office') == 'on'
        
        audit_date_str = request.form.get('audit_certificate_date', '')
        rec.audit_certificate_date = datetime.strptime(audit_date_str, '%Y-%m-%d').date() if audit_date_str else None
        
        rec.send_to_intl_division = request.form.get('send_to_intl_division') == 'on'
        rec.send_to_bangladesh_bank = request.form.get('send_to_bangladesh_bank') == 'on'
        
        rec.benefit_receivable = float(request.form.get('benefit_receivable', '0').replace(',', '') or 0.0)
        rec.benefit_received = float(request.form.get('benefit_received', '0').replace(',', '') or 0.0)
        
        recv_date_str = request.form.get('benefit_received_date', '')
        rec.benefit_received_date = datetime.strptime(recv_date_str, '%Y-%m-%d').date() if recv_date_str else None
        
        rec.status = request.form.get('status', rec.status)
        rec.notes = request.form.get('notes', rec.notes).strip()
        
        db.session.commit()
        flash(f'Record #{rec.id} updated successfully!', 'success')
        return redirect(url_for('remittance_dashboard'))
        
    return render_template('remittance/edit.html', record=rec)

@app.route('/remittances/<int:rec_id>/delete', methods=['POST'])
@login_required
def delete_remittance_entry(rec_id):
    rec = RemittanceRecord.query.get_or_404(rec_id)
    db.session.delete(rec)
    db.session.commit()
    flash('Remittance entry deleted successfully.', 'success')
    return redirect(url_for('remittance_dashboard'))

@app.route('/remittances/statement')
@login_required
def remittance_statement_report():
    records = RemittanceRecord.query.order_by(RemittanceRecord.date.asc(), RemittanceRecord.id.asc()).all()
    running_bal = 0.0
    for r in records:
        if r.budget_amount > 0:
            running_bal -= r.budget_amount
        if r.received_bdt > 0:
            running_bal += r.received_bdt
        r.calculated_balance = running_bal

    total_budget_issued = sum(r.budget_amount for r in records)
    total_received_bdt = sum(r.received_bdt for r in records)
    total_remittance_usd = sum(r.remittance_usd for r in records)
    total_incentive_claimed = sum(r.incentive_10_bdt for r in records)
    total_benefit_received = sum(r.benefit_received for r in records)

    return render_template(
        'remittance/statement.html',
        records=records,
        total_budget_issued=total_budget_issued,
        total_received_bdt=total_received_bdt,
        total_remittance_usd=total_remittance_usd,
        total_incentive_claimed=total_incentive_claimed,
        total_benefit_received=total_benefit_received,
        generated_at=datetime.utcnow()
    )

@app.route('/remittances/api/toggle-checkpoint', methods=['POST'])
@login_required
def api_toggle_remittance_checkpoint():
    rec_id = request.form.get('id', type=int)
    field = request.form.get('field', '')
    val = request.form.get('val') == 'true'
    
    rec = RemittanceRecord.query.get(rec_id)
    if not rec:
        return json.dumps({'success': False, 'error': 'Not found'}), 404
        
    valid_fields = [
        'tt_certificate', 'encashment_certificate', 'authorized_dealer_cert',
        'cash_incentive_cert', 'send_to_bank', 'send_to_head_office',
        'send_to_audit_office', 'send_to_intl_division', 'send_to_bangladesh_bank'
    ]
    if field in valid_fields:
        setattr(rec, field, val)
        db.session.commit()
        return json.dumps({'success': True, 'field': field, 'val': val})
    return json.dumps({'success': False, 'error': 'Invalid field'}), 400


# ==============================================================================
# SAMPLE SEED DATA GENERATOR
# ==============================================================================

def seed_sample_accounting_data():
    """Initializes realistic Chart of Accounts, Subledgers, Users, and Employees from real corporate data."""
    if Account.query.count() == 0:
        try:
            from import_real_data import run_import
            run_import()
        except Exception as e:
            print(f"Error seeding real data: {e}")
    seed_sample_budgets()
    seed_sample_remittances()

def seed_sample_remittances():
    try:
        from seed_remittance_sample import seed_remittance_records
        seed_remittance_records()
    except Exception as e:
        print(f"Error seeding sample remittances: {e}")

def seed_sample_budgets():
    """Initializes realistic monthly budgets for past and active months."""
    try:
        if MonthlyBudget.query.count() == 0:
            expense_accounts = Account.query.filter_by(type='Expense').all()
            acc_map = {acc.name: acc for acc in expense_accounts}
            
            # --- 1. August 2026 Issued Budget ---
            b_aug = MonthlyBudget(
                month_year='2026-08',
                title='Corporate Operational Budget - August 2026',
                fiscal_year='2026-2027',
                status='Closed',
                notes='Full monthly operational allocation for all administrative, ISP, and staff expenses.',
                issued_by='Admin'
            )
            db.session.add(b_aug)
            db.session.flush()
            
            aug_allocs = {
                'Salary': 3000000.0,
                'Electricity Bill': 350000.0,
                'Incentive for Sale': 350000.0,
                'ISP-Summit Communications Ltd.': 310000.0,
                'House Rent': 275000.0,
                'Transportation (Car rent)': 105000.0,
                'ISP-SKYTEL Communications Ltd.': 100000.0,
                'Fuel for Vehicle': 75000.0,
                'Director Expenses': 60000.0,
                'ISP Optimax': 40000.0,
                'Entertainment': 25000.0,
                'Entertainment-Breakfast': 10000.0,
                "Entertainment-Night Snack's": 20000.0,
                'Stationary': 15000.0,
                'Office Expencess': 25000.0,
                'Repair & Maintenance': 30000.0,
                'Mobile Bill': 18000.0,
                'Internet bill': 10000.0,
                'Water bill': 15000.0,
                'Legel fee': 30000.0,
                'Medical Expenses': 5000.0,
                'Bank Charge': 5000.0,
                'Cleaner wages': 8000.0
            }
            
            tot_aug = 0.0
            for name, amt in aug_allocs.items():
                if name in acc_map:
                    item = BudgetItem(budget_id=b_aug.id, account_id=acc_map[name].id, allocated_amount=amt)
                    db.session.add(item)
                    tot_aug += amt
            b_aug.total_budgeted = tot_aug

            # --- 2. September 2026 Active Issued Budget ---
            b_sep = MonthlyBudget(
                month_year='2026-09',
                title='Executive Operating Budget - September 2026',
                fiscal_year='2026-2027',
                status='Issued',
                notes='Current active monthly budget allocation. Target burn rate < 95% across departments.',
                issued_by='Admin'
            )
            db.session.add(b_sep)
            db.session.flush()
            
            sep_allocs = {
                'Salary': 3100000.0,
                'Electricity Bill': 360000.0,
                'Incentive for Sale': 360000.0,
                'ISP-Summit Communications Ltd.': 320000.0,
                'House Rent': 280000.0,
                'Transportation (Car rent)': 110000.0,
                'ISP-SKYTEL Communications Ltd.': 105000.0,
                'Fuel for Vehicle': 80000.0,
                'Director Expenses': 65000.0,
                'ISP Optimax': 42000.0,
                'Entertainment': 25000.0,
                'Entertainment-Breakfast': 12000.0,
                "Entertainment-Night Snack's": 22000.0,
                'Stationary': 15000.0,
                'Office Expencess': 25000.0,
                'Repair & Maintenance': 35000.0,
                'Mobile Bill': 18000.0,
                'Internet bill': 10000.0,
                'Water bill': 15000.0,
                'Legel fee': 30000.0,
                'Medical Expenses': 5000.0,
                'Bank Charge': 5000.0,
                'Cleaner wages': 8000.0
            }
            
            tot_sep = 0.0
            for name, amt in sep_allocs.items():
                if name in acc_map:
                    item = BudgetItem(budget_id=b_sep.id, account_id=acc_map[name].id, allocated_amount=amt)
                    db.session.add(item)
                    tot_sep += amt
            b_sep.total_budgeted = tot_sep

            # Also seed active September 2026 vouchers if none exist, so that September has live actuals
            sep_entries_count = db.session.query(db.func.count(JournalEntry.id)).filter(
                db.func.strftime('%Y-%m', JournalEntry.date) == '2026-09'
            ).scalar() or 0
            
            if sep_entries_count == 0:
                cash_acc = Account.query.filter_by(is_cash_or_bank=True).first()
                if cash_acc:
                    vouchers_data = [
                        ('PV-2026-09-001', 'Payment', date(2026, 9, 3), 'Office Rent for September 2026', 280000.0, 'House Rent'),
                        ('PV-2026-09-002', 'Payment', date(2026, 9, 7), 'DESCO Electricity Utility Bill - Sep', 320000.0, 'Electricity Bill'),
                        ('PV-2026-09-003', 'Payment', date(2026, 9, 10), 'Corporate Vehicle Fuel & Maintenance', 64500.0, 'Fuel for Vehicle'),
                        ('PV-2026-09-004', 'Payment', date(2026, 9, 15), 'Mid-Month Salary Advance & Staff Allowances', 2150000.0, 'Salary'),
                        ('PV-2026-09-005', 'Payment', date(2026, 9, 18), 'ISP Bandwidth & Summit Comm Dues', 305000.0, 'ISP-Summit Communications Ltd.'),
                        ('PV-2026-09-006', 'Payment', date(2026, 9, 22), 'Office Stationery, Refreshments & Supplies', 38500.0, 'Office Expencess'),
                    ]
                    for v_no, v_type, v_date, narr, amt, exp_acc_name in vouchers_data:
                        exp_a = acc_map.get(exp_acc_name)
                        if exp_a:
                            je = JournalEntry(
                                voucher_no=v_no,
                                voucher_type=v_type,
                                date=v_date,
                                reference='SEP-EXP',
                                narration=narr,
                                total_amount=amt
                            )
                            db.session.add(je)
                            db.session.flush()
                            line_dr = TransactionLine(
                                journal_entry_id=je.id,
                                account_id=exp_a.id,
                                debit=amt,
                                credit=0.0,
                                particulars=narr
                            )
                            line_cr = TransactionLine(
                                journal_entry_id=je.id,
                                account_id=cash_acc.id,
                                debit=0.0,
                                credit=amt,
                                particulars=narr
                            )
                            db.session.add_all([line_dr, line_cr])
                            
            db.session.commit()
            print(" * Sample monthly budgets and September actual vouchers seeded successfully.")
    except Exception as e:
        db.session.rollback()
        print(f"Error seeding sample budgets: {e}")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        seed_sample_accounting_data()
    print(" * ZR-ERP Enterprise Running on http://127.0.0.1:5000")
    app.run(debug=True, host='127.0.0.1', port=5000)