import os
import sys
import re
from datetime import datetime, date

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from app import app, db, User, Account, JournalEntry, TransactionLine, Employee, PayrollRecord, PayrollItem, CashBookRow, recalculate_account_balances

def clean_num(s):
    if not s or s.strip() in ("-", "", "#N/A"):
        return 0.0
    s = s.strip().replace(",", "")
    try:
        return float(s)
    except:
        return 0.0

def parse_date(date_str):
    if not date_str or date_str in ("#N/A", "-"):
        return None
    date_str = date_str.strip()
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except:
            pass
    return None

def run_import():
    print("==================================================")
    print("STARTING FULL GENERAL LEDGER IMPORT FROM CASH BOOK")
    print("==================================================")

    with app.app_context():
        # 1. Reset / recreate clean database schema
        db.drop_all()
        db.create_all()
        print(" [✓] Clean database schema created.")

        # 2. Seed Users
        admin_user = User(username="admin", password="admin123", full_name="Md. Zillur Rahman (Finance)", role="Admin")
        zillur_user = User(username="zillur", password="zillur123", full_name="Md. Zillur Rahman", role="Admin")
        shaukat_user = User(username="shaukat", password="user123", full_name="Mir Shaukat Ali", role="General User")
        staff_user = User(username="accountant", password="user123", full_name="Accounts Executive", role="General User")
        db.session.add_all([admin_user, zillur_user, shaukat_user, staff_user])
        db.session.commit()
        print(" [✓] Multi-users created: 'admin', 'zillur', 'shaukat', 'accountant'.")

        # 3. Seed Chart of Accounts
        # Define master list of 65 accounts
        account_definitions = [
            # Liquid Assets (Cash & Bank)
            {"code": "BS-001", "name": "Cash in Hand", "type": "Asset", "category": "Current Asset", "is_cash_or_bank": True, "open": 1676799.00},
            {"code": "BS-002", "name": "Premier Bank", "type": "Asset", "category": "Current Asset", "is_cash_or_bank": True, "open": -279306.00},
            {"code": "BS-003", "name": "Cash withdraw", "type": "Contra", "category": "Contra Account", "is_cash_or_bank": False, "open": 0.0},
            
            # Equity & Capital
            {"code": "BS-099", "name": "Opening Balance Equity", "type": "Equity", "category": "Equity", "is_cash_or_bank": False, "open": 1397493.00},

            # Operating & Other Revenue
            {"code": "BS-004", "name": "Remittance Received (Sales)", "type": "Revenue", "category": "Operating Revenue", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-035", "name": "Accounts-TBN-25", "type": "Asset", "category": "Current Asset", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-090", "name": "Incentive from bangladesh bank", "type": "Revenue", "category": "Other Income", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-108", "name": "Income form ISP", "type": "Revenue", "category": "Operating Revenue", "is_cash_or_bank": False, "open": 0.0},

            # Balance Sheet - Non-Current Assets
            {"code": "BS-010", "name": "Furniture & Fixture", "type": "Asset", "category": "Non-Current Asset", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-013", "name": "Computer Accessories", "type": "Asset", "category": "Current Asset", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-014", "name": "Office Equipments", "type": "Asset", "category": "Non-Current Asset", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-032", "name": "Accessories", "type": "Asset", "category": "Current Asset", "is_cash_or_bank": False, "open": 0.0},

            # Balance Sheet - Liabilities
            {"code": "BS-006", "name": "Internal Loan (Director)", "type": "Liability", "category": "Current Liability", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-022", "name": "PF employee contribution", "type": "Liability", "category": "Current Liability", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-023", "name": "PF employer contribution", "type": "Liability", "category": "Current Liability", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-040", "name": "Accounts CSF -17", "type": "Liability", "category": "Current Liability", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-070", "name": "PF Loan Return Diposit", "type": "Liability", "category": "Current Liability", "is_cash_or_bank": False, "open": 0.0},

            # Equity / Drawings
            {"code": "BS-034", "name": "Director Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},

            # Operating Expenses
            {"code": "IS-000", "name": "Salary", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "BS-029", "name": "Eid bonus", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-011", "name": "Transportation (Car rent)", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-012", "name": "Allowance & Gift", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-014", "name": "Conveyance", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-015", "name": "Tour Expenses", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-022", "name": "Programatic Cost", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-024", "name": "Fuel for Vehicle", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-029", "name": "ISP Optimax", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-032", "name": "Fuel for Generator", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-037", "name": "Makeup Expenses", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-039", "name": "Advartisement", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-041", "name": "Development expenses", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-046", "name": "Incentive for Sale", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-066", "name": "Cleaner wages", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-072", "name": "ISP-Amra network", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-086", "name": "ISP-Summit Communications Ltd.", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-087", "name": "ISP-SKYTEL Communications Ltd.", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-092", "name": "Festival Expense", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-106", "name": "BSCL Expenses", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-111", "name": "Device Recharge", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-116", "name": "EGP Tendar", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-117", "name": "Election Expenses (BD)", "type": "Expense", "category": "Operating Expense", "is_cash_or_bank": False, "open": 0.0},

            # Administrative Expenses
            {"code": "IS-001", "name": "House Rent", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-003", "name": "Electricity Bill", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-005", "name": "Gas Bill", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-006", "name": "Entertainment", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-101", "name": "Entertainment-Breakfast", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-102", "name": "Entertainment-Night Snack's", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-008", "name": "Stationary", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-016", "name": "Office Expencess", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-017", "name": "Repair & Maintenance", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-018", "name": "Postage and Courier", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-019", "name": "Printing cost", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-021", "name": "Miscellaneous Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-023", "name": "Mobile Bill", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-025", "name": "Internet bill", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-031", "name": "Water bill", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-043", "name": "Legel fee", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-044", "name": "Iftar and sahri", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-062", "name": "Medical Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-064", "name": "Vat Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-075", "name": "TDS expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-079", "name": "BASIS Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-083", "name": "Stamp Duty", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-088", "name": "Tax office -Return submission", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-098", "name": "Donation Expenses", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-109", "name": "Vat Miscellenous", "type": "Expense", "category": "Administrative Expense", "is_cash_or_bank": False, "open": 0.0},

            # Financial Expenses
            {"code": "IS-020", "name": "Bank Charge", "type": "Expense", "category": "Financial Expense", "is_cash_or_bank": False, "open": 0.0},
            {"code": "IS-095", "name": "Cash Incentive Processing Fee", "type": "Expense", "category": "Financial Expense", "is_cash_or_bank": False, "open": 0.0},
        ]

        acc_obj_map = {}
        for item in account_definitions:
            acc = Account(
                code=item["code"],
                name=item["name"],
                type=item["type"],
                category=item["category"],
                is_cash_or_bank=item["is_cash_or_bank"],
                opening_balance=item["open"]
            )
            db.session.add(acc)
            acc_obj_map[item["code"]] = acc

        db.session.commit()

        # Link Subledgers
        ent_acc = Account.query.filter_by(code="IS-006").first()
        if ent_acc:
            b_fast = Account.query.filter_by(code="IS-101").first()
            if b_fast: b_fast.parent_id = ent_acc.id
            n_snack = Account.query.filter_by(code="IS-102").first()
            if n_snack: n_snack.parent_id = ent_acc.id
            db.session.commit()

        print(f" [✓] Chart of Accounts initialized with {Account.query.count()} accounts & subledgers.")

        # 4. Seed Corporate Employees (from actual salary disbursements)
        real_employees = [
            ("EMP-001", "Md. Zillur Rahman", "Senior Accounts Executive", "Finance & Accounts", 61176.00, 5000.00, 0.00, "Premier Bank 108-992"),
            ("EMP-002", "Mir Shaukat Ali", "Senior Operations Manager", "Administration", 74241.00, 6000.00, 0.00, "Premier Bank 108-993"),
            ("EMP-003", "Md. Rahat Pathan", "Administrative Officer", "Administration", 41287.00, 4000.00, 0.00, "Premier Bank 108-994"),
            ("EMP-004", "Kamrul Islam Saleh", "Senior Producer", "News Department", 31894.00, 3000.00, 0.00, "Premier Bank 108-995"),
            ("EMP-005", "Mostafaye Abeer", "Production Executive", "News Department", 15613.00, 2000.00, 0.00, "Premier Bank 108-996"),
            ("EMP-006", "Shahanur Shahin", "Broadcast Engineer", "Technical & IT", 25205.00, 2500.00, 0.00, "Premier Bank 108-997"),
            ("EMP-007", "Basu Debnath", "Broadcast Journalist", "Editorial", 18000.00, 2000.00, 0.00, "Premier Bank 108-998"),
            ("EMP-008", "Tasnim Jahan Tazin", "Customer Relations Executive", "CRD", 16010.00, 1500.00, 0.00, "Premier Bank 108-999"),
            ("EMP-009", "Afroza Khatun", "Accounts Officer", "Finance & Accounts", 18000.00, 1500.00, 0.00, "Premier Bank 108-100"),
            ("EMP-010", "Mahmudul Hasan", "IT Support Specialist", "Technical & IT", 15000.00, 1500.00, 0.00, "Premier Bank 108-101"),
            ("EMP-011", "Md. Sohel", "Office Attendant", "General Services", 15500.00, 1000.00, 0.00, "Cash Disbursement"),
            ("EMP-012", "Md. Karim", "Office Attendant", "General Services", 12000.00, 1000.00, 0.00, "Cash Disbursement"),
            ("EMP-013", "Md. Din Islam", "Office Attendant", "General Services", 12000.00, 1000.00, 0.00, "Cash Disbursement"),
            ("EMP-014", "Mizan Driver", "Senior Driver", "Logistics & Transport", 25000.00, 2000.00, 0.00, "Cash Disbursement"),
            ("EMP-015", "Nannu Driver", "Driver", "Logistics & Transport", 24000.00, 2000.00, 0.00, "Cash Disbursement"),
        ]
        for emp_tuple in real_employees:
            e = Employee(
                employee_id=emp_tuple[0],
                name=emp_tuple[1],
                designation=emp_tuple[2],
                department=emp_tuple[3],
                basic_salary=emp_tuple[4],
                allowances=emp_tuple[5],
                deductions=emp_tuple[6],
                bank_account_no=emp_tuple[7],
                status="Active"
            )
            db.session.add(e)
        db.session.commit()
        print(f" [✓] {Employee.query.count()} corporate employee profiles loaded for payroll.")

        # Cache account IDs
        acc_id_map = {a.code: a.id for a in Account.query.all()}
        cash_id = acc_id_map["BS-001"]
        prem_id = acc_id_map["BS-002"]

        # 5. Parse and Import All Journal Entries, Vouchers & Cash Book Rows
        tsv_path = "client_cash_book_data.tsv"
        if not os.path.exists(tsv_path):
            tsv_path = "raw_user_input.txt"

        with open(tsv_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        rv_counter = 1
        cv_counter = 1
        pv_gen_counter = 1
        voucher_count = 0
        cash_book_row_count = 0
        used_voucher_numbers = set()

        def get_unique_vno(vno_candidate, vtype):
            nonlocal rv_counter, cv_counter, pv_gen_counter
            if vno_candidate and vno_candidate not in used_voucher_numbers:
                used_voucher_numbers.add(vno_candidate)
                return vno_candidate
            if vtype == "Contra":
                candidate = f"CV-2026-{cv_counter:04d}"
                cv_counter += 1
            elif vtype == "Receipt":
                candidate = f"RV-2026-{rv_counter:04d}"
                rv_counter += 1
            else:
                candidate = f"PV-2026-{pv_gen_counter:04d}"
                pv_gen_counter += 1
            while candidate in used_voucher_numbers:
                if vtype == "Contra":
                    candidate = f"CV-2026-{cv_counter:04d}"
                    cv_counter += 1
                elif vtype == "Receipt":
                    candidate = f"RV-2026-{rv_counter:04d}"
                    rv_counter += 1
                else:
                    candidate = f"PV-2026-{pv_gen_counter:04d}"
                    pv_gen_counter += 1
            used_voucher_numbers.add(candidate)
            return candidate

        batch_entries = []
        batch_lines = []

        for line_idx, line in enumerate(lines):
            if not line.strip() or "Receipts" in line or "Voucher No." in line:
                continue
            parts = line.split("\t")
            if len(parts) < 20:
                continue

            remarks = ""
            for k in range(23, len(parts)):
                if parts[k].strip():
                    remarks = parts[k].strip()
                    break

            r_code = parts[0].strip() if len(parts) > 0 else ""
            r_date_str = parts[2].strip() if len(parts) > 2 and parts[2].strip() else (parts[1].strip() if len(parts) > 1 else "")
            r_voucher = parts[3].strip() if len(parts) > 3 else ""
            r_desc = parts[4].strip() if len(parts) > 4 else ""
            r_ref = parts[5].strip() if len(parts) > 5 else ""
            r_cash = clean_num(parts[6]) if len(parts) > 6 else 0.0
            r_prem = clean_num(parts[7]) if len(parts) > 7 else 0.0
            r_bank = clean_num(parts[8]) if len(parts) > 8 else 0.0
            r_total = clean_num(parts[9]) if (len(parts) > 9 and parts[9].strip() and parts[9].strip() != "-") else (r_cash + r_prem)

            p_date_str = parts[11].strip() if len(parts) > 11 else ""
            p_vno = parts[12].strip() if len(parts) > 12 else ""
            p_code = parts[13].strip() if len(parts) > 13 else ""
            p_desc = parts[14].strip() if len(parts) > 14 else ""
            p_ref = parts[15].strip() if len(parts) > 15 else ""
            p_cash = clean_num(parts[16]) if len(parts) > 16 else 0.0
            p_bank = clean_num(parts[17]) if len(parts) > 17 else 0.0
            p_prem = clean_num(parts[18]) if len(parts) > 18 else 0.0
            p_total = clean_num(parts[19]) if (len(parts) > 19 and parts[19].strip() and parts[19].strip() != "-") else (p_cash + p_prem + p_bank)

            bal_cash = clean_num(parts[20]) if len(parts) > 20 else 0.0
            bal_bank = clean_num(parts[21]) if len(parts) > 21 else 0.0
            bal_prem = clean_num(parts[22]) if len(parts) > 22 else 0.0

            # Create Cash Book Row
            cb_row = CashBookRow(
                row_order=line_idx,
                receipt_code=r_code if r_code != "#N/A" else "",
                receipt_date=parse_date(r_date_str),
                receipt_voucher=r_voucher,
                receipt_desc=r_desc if r_desc != "#N/A" else "",
                receipt_ref=r_ref,
                receipt_cash=r_cash,
                receipt_premier=r_prem,
                receipt_bank=r_bank,
                receipt_total=r_total,
                payment_date=parse_date(p_date_str),
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
            cash_book_row_count += 1

            if "Opening Balance" in r_desc or "Opeining Balance" in r_desc:
                continue

            # CASE 1: Contra Transaction (BS-003 Cash withdraw / deposit)
            if r_code == "BS-003" or p_code == "BS-003":
                # Line 1535 is the paired payment side of Line 1533's withdrawal
                if line_idx == 1535:
                    continue

                c_date = parse_date(r_date_str) or parse_date(p_date_str) or date(2026, 1, 1)
                cv_num = get_unique_vno(None, "Contra")

                if (r_cash > 0 and (p_prem > 0 or p_total > 0)) or (r_cash == 221720):
                    amt = r_cash
                    cv_entry = JournalEntry(
                        voucher_no=cv_num,
                        voucher_type="Contra",
                        date=c_date,
                        reference="Cash Withdraw",
                        narration=remarks or "Cash withdrawal from Premier Bank to Cash in Hand",
                        total_amount=round(amt, 2)
                    )
                    db.session.add(cv_entry)
                    db.session.flush()

                    db.session.add(TransactionLine(journal_entry_id=cv_entry.id, account_id=cash_id, debit=round(amt, 2), credit=0.00, particulars="Cash received"))
                    db.session.add(TransactionLine(journal_entry_id=cv_entry.id, account_id=prem_id, debit=0.00, credit=round(amt, 2), particulars="Premier Bank withdrawal payment"))
                    voucher_count += 1

                elif r_prem > 0 and (p_cash > 0 or p_total > 0):
                    amt = r_prem
                    cv_entry = JournalEntry(
                        voucher_no=cv_num,
                        voucher_type="Contra",
                        date=c_date,
                        reference="Cash Deposit",
                        narration=remarks or "Cash deposit to Premier Bank from Cash in Hand",
                        total_amount=round(amt, 2)
                    )
                    db.session.add(cv_entry)
                    db.session.flush()

                    db.session.add(TransactionLine(journal_entry_id=cv_entry.id, account_id=prem_id, debit=round(amt, 2), credit=0.00, particulars="Premier Bank deposit"))
                    db.session.add(TransactionLine(journal_entry_id=cv_entry.id, account_id=cash_id, debit=0.00, credit=round(amt, 2), particulars="Cash in Hand disbursed"))
                    voucher_count += 1

                # If there was ALSO a payment on that line (e.g. Line 1533 Conveyance 140)
                if p_code and p_code != "BS-003" and p_total > 0:
                    p_acc_id = acc_id_map.get(p_code)
                    if p_acc_id:
                        pv_num = get_unique_vno(p_vno, "Payment")
                        pv_date = parse_date(p_date_str) or c_date
                        pv_entry = JournalEntry(
                            voucher_no=pv_num,
                            voucher_type="Payment",
                            date=pv_date,
                            reference=p_ref or None,
                            narration=remarks or f"Payment for {p_desc} ({p_code})",
                            total_amount=round(p_total, 2)
                        )
                        db.session.add(pv_entry)
                        db.session.flush()

                        db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=p_acc_id, debit=round(p_total, 2), credit=0.00, particulars=p_desc))
                        if p_cash > 0:
                            db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=cash_id, debit=0.00, credit=round(p_cash, 2), particulars="Paid in Cash"))
                        if p_prem > 0:
                            db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=prem_id, debit=0.00, credit=round(p_prem, 2), particulars="Paid via Premier Bank"))
                        voucher_count += 1
                continue

            # CASE 2: Receipt Voucher (RV)
            if r_total > 0:
                r_acc_id = acc_id_map.get(r_code)
                if not r_acc_id:
                    # Default to Accounts-TBN-25 if unknown
                    r_acc_id = acc_id_map["BS-035"]

                r_date = parse_date(r_date_str) or parse_date(p_date_str) or date(2026, 1, 1)
                rv_num = get_unique_vno(r_voucher, "Receipt")

                rv_entry = JournalEntry(
                    voucher_no=rv_num,
                    voucher_type="Receipt",
                    date=r_date,
                    reference=r_ref or None,
                    narration=remarks if (remarks and p_total == 0) else f"Receipt from {r_desc} ({r_code})",
                    total_amount=round(r_total, 2)
                )
                db.session.add(rv_entry)
                db.session.flush()

                if r_cash > 0:
                    db.session.add(TransactionLine(journal_entry_id=rv_entry.id, account_id=cash_id, debit=round(r_cash, 2), credit=0.00, particulars="Cash received"))
                if r_prem > 0:
                    db.session.add(TransactionLine(journal_entry_id=rv_entry.id, account_id=prem_id, debit=round(r_prem, 2), credit=0.00, particulars="Premier Bank received"))

                db.session.add(TransactionLine(journal_entry_id=rv_entry.id, account_id=r_acc_id, debit=0.00, credit=round(r_total, 2), particulars=r_desc or "Receipt Credit"))
                voucher_count += 1

            # CASE 3: Payment Voucher (PV)
            if p_total > 0:
                p_acc_id = acc_id_map.get(p_code)
                if not p_acc_id:
                    p_acc_id = acc_id_map["IS-016"] # Office Expencess default

                p_date = parse_date(p_date_str) or parse_date(r_date_str) or date(2026, 1, 1)
                pv_num = get_unique_vno(p_vno, "Payment")

                pv_entry = JournalEntry(
                    voucher_no=pv_num,
                    voucher_type="Payment",
                    date=p_date,
                    reference=p_ref or None,
                    narration=remarks or f"Payment for {p_desc} ({p_code})",
                    total_amount=round(p_total, 2)
                )
                db.session.add(pv_entry)
                db.session.flush()

                db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=p_acc_id, debit=round(p_total, 2), credit=0.00, particulars=p_desc))
                if p_cash > 0:
                    db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=cash_id, debit=0.00, credit=round(p_cash, 2), particulars="Paid in Cash"))
                if p_prem > 0:
                    db.session.add(TransactionLine(journal_entry_id=pv_entry.id, account_id=prem_id, debit=0.00, credit=round(p_prem, 2), particulars="Paid via Premier Bank"))
                voucher_count += 1

        db.session.commit()
        print(f" [✓] Total {voucher_count} General Ledger Vouchers, {TransactionLine.query.count()} lines, and {CashBookRow.query.count()} Cash Book Rows inserted.")

        # 6. Recalculate all Ledger Balances
        recalculate_account_balances()
        print(" [✓] Account balances recalculated successfully.")

        # 7. Verification Audit
        cash_acc = Account.query.filter_by(code="BS-001").first()
        prem_acc = Account.query.filter_by(code="BS-002").first()

        print("\n==================================================")
        print("IMPORT AUDIT & VERIFICATION METRICS")
        print("==================================================")
        print(f" Cash in Hand Ending Balance:  ৳{cash_acc.balance:,.2f}")
        print(f" Premier Bank Ending Balance:  ৳{prem_acc.balance:,.2f}")
        print(f" Combined Liquid Balance:      ৳{(cash_acc.balance + prem_acc.balance):,.2f}")

        # Check total dr vs cr across all lines
        dr_sum = db.session.query(db.func.sum(TransactionLine.debit)).scalar() or 0.0
        cr_sum = db.session.query(db.func.sum(TransactionLine.credit)).scalar() or 0.0
        print(f" Total General Ledger Debits:  ৳{dr_sum:,.2f}")
        print(f" Total General Ledger Credits: ৳{cr_sum:,.2f}")
        print(f" Journal Difference:           ৳{abs(dr_sum - cr_sum):,.2f}")

        if abs(dr_sum - cr_sum) < 0.01:
            print(" >>> AUDIT PASSED: The General Ledger is 100% Balanced! <<<")
        else:
            print(" >>> AUDIT WARNING: General Ledger debits and credits do not match! <<<")

if __name__ == '__main__':
    run_import()
