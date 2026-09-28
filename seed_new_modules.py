import io
from datetime import date, datetime
from app import app, db, Account, Employee, BankStatement, BankStatementLine, BankReconciliation, EmployeeFundContribution, EmployeeSalaryHistory, TransactionLine, JournalEntry

with app.app_context():
    print("Seeding initial data for new modules...")

    # 1. PREMIER BANK STATEMENT & RECONCILIATION SEEDING
    premier_acc = Account.query.filter_by(code='BS-002').first()
    if not premier_acc:
        premier_acc = Account.query.filter(Account.name.ilike('%Premier Bank%')).first()

    if premier_acc:
        existing_stmt = BankStatement.query.filter_by(bank_account_id=premier_acc.id).first()
        if not existing_stmt:
            jan_lines = TransactionLine.query.join(JournalEntry).filter(
                TransactionLine.account_id == premier_acc.id,
                JournalEntry.date >= date(2026, 1, 1),
                JournalEntry.date <= date(2026, 1, 31)
            ).order_by(JournalEntry.date.asc(), JournalEntry.id.asc()).all()

            stmt = BankStatement(
                bank_account_id=premier_acc.id,
                statement_title="Premier Bank Corporate Statement - Jan 2026",
                filename="Premier_Bank_Stmt_Jan2026.pdf",
                from_date=date(2026, 1, 1),
                to_date=date(2026, 1, 31),
                opening_balance=1397493.00,
                closing_balance=1397493.00,
                total_withdrawals=0.00,
                total_deposits=0.00
            )
            db.session.add(stmt)
            db.session.flush()

            running_bal = 1397493.00
            tot_w = 0.0
            tot_d = 0.0

            op_line = BankStatementLine(
                statement_id=stmt.id,
                bank_account_id=premier_acc.id,
                txn_date=date(2026, 1, 1),
                value_date=date(2026, 1, 1),
                particulars="OPENING BALANCE BROUGHT FORWARD",
                cheque_ref="B/F",
                withdrawal_dr=0.00,
                deposit_cr=0.00,
                balance=running_bal,
                is_reconciled=True,
                reconciled_date=date(2026, 1, 1)
            )
            db.session.add(op_line)

            for idx, tl in enumerate(jan_lines[:30]):
                v = tl.journal_entry
                is_deposit = tl.debit > 0
                amount = tl.debit if is_deposit else tl.credit
                if amount <= 0:
                    continue

                if is_deposit:
                    w_dr = 0.0
                    d_cr = amount
                    running_bal += amount
                    tot_d += amount
                else:
                    w_dr = amount
                    d_cr = 0.0
                    running_bal -= amount
                    tot_w += amount

                is_recon = idx < 28
                b_line = BankStatementLine(
                    statement_id=stmt.id,
                    bank_account_id=premier_acc.id,
                    txn_date=v.date,
                    value_date=v.date,
                    particulars=f"{v.narration[:80]}",
                    cheque_ref=v.reference or f"CHQ-{10000+idx}",
                    withdrawal_dr=w_dr,
                    deposit_cr=d_cr,
                    balance=running_bal,
                    is_reconciled=is_recon,
                    reconciled_date=v.date if is_recon else None,
                    matched_voucher_id=v.id if is_recon else None,
                    matched_voucher_no=v.voucher_no if is_recon else None
                )
                db.session.add(b_line)

            running_bal += 15250.00
            tot_d += 15250.00
            db.session.add(BankStatementLine(
                statement_id=stmt.id,
                bank_account_id=premier_acc.id,
                txn_date=date(2026, 1, 31),
                value_date=date(2026, 1, 31),
                particulars="CREDIT INTEREST FOR JAN-2026",
                cheque_ref="INT-JAN26",
                withdrawal_dr=0.00,
                deposit_cr=15250.00,
                balance=running_bal,
                is_reconciled=False
            ))

            running_bal -= 1150.00
            tot_w += 1150.00
            db.session.add(BankStatementLine(
                statement_id=stmt.id,
                bank_account_id=premier_acc.id,
                txn_date=date(2026, 1, 31),
                value_date=date(2026, 1, 31),
                particulars="EXCISE DUTY & HALF YEARLY LEDGER FEE",
                cheque_ref="CHG-JAN26",
                withdrawal_dr=1150.00,
                deposit_cr=0.00,
                balance=running_bal,
                is_reconciled=False
            ))

            stmt.closing_balance = running_bal
            stmt.total_withdrawals = tot_w
            stmt.total_deposits = tot_d

            recon = BankReconciliation(
                bank_account_id=premier_acc.id,
                recon_month="2026-01",
                statement_date=date(2026, 1, 31),
                statement_balance=running_bal,
                book_balance=running_bal - 15250.00 + 1150.00,
                uncredited_deposits=0.00,
                unpresented_cheques=0.00,
                bank_charges=1150.00,
                direct_credits=15250.00,
                adjusted_bank_balance=running_bal,
                adjusted_book_balance=running_bal,
                variance=0.00,
                status="Reconciled",
                notes="January 2026 Month-End Reconciliation verified against Premier Bank Official e-Statement.",
                reconciled_by="Md. Zillur Rahman (Senior Accounts Exec)"
            )
            db.session.add(recon)
            print(" [OK] Seeded Premier Bank Statement and Bank Reconciliation.")

    # 2. SEED EMPLOYEE CONTRIBUTION FUND
    employees = Employee.query.order_by(Employee.employee_id).all()
    existing_funds = EmployeeFundContribution.query.count()
    if existing_funds == 0:
        months = ["2025-11", "2025-12", "2026-01", "2026-02"]
        for emp in employees:
            cumul = 0.0
            emp_share = round(emp.basic_salary * 0.08, 2)
            comp_share = round(emp.basic_salary * 0.08, 2)
            
            for m_idx, m_str in enumerate(months):
                y, m = map(int, m_str.split('-'))
                c_date = date(y, m, 28)
                welfare = 200.00
                profit = round(cumul * 0.005, 2) if cumul > 0 else 0.0
                monthly_tot = emp_share + comp_share + welfare + profit
                cumul += monthly_tot
                
                cf = EmployeeFundContribution(
                    employee_id=emp.id,
                    contribution_month=m_str,
                    contribution_date=c_date,
                    employee_share=emp_share,
                    employer_share=comp_share,
                    welfare_fund=welfare,
                    profit_interest=profit,
                    total_monthly=monthly_tot,
                    cumulative_balance=cumul,
                    remarks=f"Regular monthly provident & welfare fund contribution ({m_str})"
                )
                db.session.add(cf)
        print(f" [OK] Seeded Employee Contribution Fund for {len(employees)} employees across 4 months.")

    # 3. SEED PREVIOUS SALARY HISTORY
    existing_hist = EmployeeSalaryHistory.query.count()
    if existing_hist == 0:
        hist_months = [
            ("2025-10", date(2025, 10, 31)),
            ("2025-11", date(2025, 11, 30)),
            ("2025-12", date(2025, 12, 31)),
            ("2026-01", date(2026, 1, 31))
        ]
        for emp in employees:
            b_sal = emp.basic_salary
            h_rent = round(b_sal * 0.40, 2)
            med = round(b_sal * 0.10, 2)
            conv = round(b_sal * 0.05, 2)
            oth_all = round(emp.allowances - (h_rent + med + conv), 2)
            if oth_all < 0:
                oth_all = 0.0
            gross = b_sal + h_rent + med + conv + oth_all

            pf = round(b_sal * 0.08, 2)
            tax = round(b_sal * 0.03, 2) if gross > 30000 else 0.0
            oth_ded = round(emp.deductions - (pf + tax), 2)
            if oth_ded < 0:
                oth_ded = 0.0
            tot_ded = pf + tax + oth_ded
            net = gross - tot_ded

            for m_str, p_date in hist_months:
                sh = EmployeeSalaryHistory(
                    employee_id=emp.id,
                    salary_month=m_str,
                    payment_date=p_date,
                    basic_salary=b_sal,
                    house_rent=h_rent,
                    medical=med,
                    conveyance=conv,
                    other_allowance=oth_all,
                    gross_salary=gross,
                    pf_deduction=pf,
                    tax_deduction=tax,
                    other_deduction=oth_ded,
                    total_deductions=tot_ded,
                    net_salary=net,
                    payment_method="Bank Transfer (BEFTN/NPSB)",
                    remarks=f"Salary disbursed via Premier Bank Corporate Transfer for {m_str}"
                )
                db.session.add(sh)
        print(f" [OK] Seeded Salary History for {len(employees)} employees across 4 past months.")

    db.session.commit()
    print("ALL SEEDING COMPLETED SUCCESSFULLY!")
