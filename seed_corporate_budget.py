"""
Seeds the exact corporate monthly budget for September 2026 as per user specification.
TBN (104 headcount), TIS (42 headcount), ISP (6 headcount).
Total Budget: ৳ 5,214,068
Gross Total: ৳ 5,306,518
Net Budget: ৳ 5,119,518 ($42,049)
"""

import json
from datetime import datetime
from app import app, db, MonthlyBudget, BudgetItem, Account

SEPTEMBER_2026_BUDGET_DATA = {
    "month_year": "2026-09",
    "title": "Corporate Budget for September 2026",
    "fiscal_year": "2026-2027",
    "status": "Issued",
    "headcounts": {
        "tbn": 104,
        "tis": 42,
        "isp": 6,
        "total": 152
    },
    "direct_expenses": [
        {
            "group": "Salary & Benefit (with company CSF)",
            "particulars": "Salary & Benefit (with company CSF)",
            "account_code": "IS-000",
            "tbn": 1940616.0,
            "tis": 1127325.0,
            "isp": 177462.0,
            "total": 3245403.0
        },
        {
            "group": "Office Rent",
            "particulars": "Rent",
            "account_code": "IS-001",
            "tbn": 162000.0,
            "tis": 60000.0,
            "isp": 35000.0,
            "total": 257000.0
        },
        {
            "group": "Office Rent",
            "particulars": "TDS on Rent",
            "account_code": "IS-075",
            "tbn": 16200.0,
            "tis": 6000.0,
            "isp": 3500.0,
            "total": 25700.0
        },
        {
            "group": "Utilities",
            "particulars": "Gas",
            "account_code": "IS-005",
            "tbn": 2160.0,
            "tis": 2160.0,
            "isp": 1080.0,
            "total": 5400.0
        },
        {
            "group": "Utilities",
            "particulars": "Electricity",
            "account_code": "IS-003",
            "tbn": 134755.0,
            "tis": 106994.0,
            "isp": 47424.0,
            "total": 289173.0
        },
        {
            "group": "Utilities",
            "particulars": "Water",
            "account_code": "IS-031",
            "tbn": 10733.0,
            "tis": 13578.0,
            "isp": 2010.0,
            "total": 26321.0
        },
        {
            "group": "Utilities",
            "particulars": "Service charge",
            "account_code": "IS-016",
            "tbn": 6000.0,
            "tis": 3000.0,
            "isp": 3000.0,
            "total": 12000.0
        },
        {
            "group": "Program Expenses (Islamic)",
            "particulars": "Program Expenses (Islamic)",
            "account_code": "IS-022",
            "tbn": 35000.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 35000.0
        },
        {
            "group": "Program Expenses (Islamic)",
            "particulars": "Makeup & Costume ex",
            "account_code": "IS-037",
            "tbn": 8000.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 8000.0
        },
        {
            "group": "Program Expenses (Islamic)",
            "particulars": "Entertainment-Breakfast",
            "account_code": "IS-101",
            "tbn": 11000.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 11000.0
        },
        {
            "group": "Program Expenses (Islamic)",
            "particulars": "Gift and Allowance (Mita)",
            "account_code": "IS-012",
            "tbn": 5000.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 5000.0
        },
        {
            "group": "Cash Incentive Expenses",
            "particulars": "Cash incentive-Basis Expenses",
            "account_code": "IS-079",
            "tbn": 13277.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 13277.0
        },
        {
            "group": "Cash Incentive Expenses",
            "particulars": "Cash Incentive- audit misc",
            "account_code": "IS-021",
            "tbn": 3500.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 3500.0
        },
        {
            "group": "Cash Incentive Expenses",
            "particulars": "Miscellinious Expenses-vat",
            "account_code": "IS-109",
            "tbn": 6750.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 6750.0
        },
        {
            "group": "Annual Selected Expenses",
            "particulars": "Eid bonus (2 Eid)",
            "account_code": "BS-029",
            "tbn": 170110.0,
            "tis": 108730.0,
            "isp": 21160.0,
            "total": 300000.0
        },
        {
            "group": "Licence & legal fee",
            "particulars": "Tax Return submission",
            "account_code": "IS-088",
            "tbn": 25000.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 25000.0
        },
        {
            "group": "Licence & legal fee",
            "particulars": "Audit fee",
            "account_code": "IS-043",
            "tbn": 10833.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 10833.0
        },
        {
            "group": "Licence & legal fee",
            "particulars": "ISP Licence Renew",
            "account_code": "IS-041",
            "tbn": 0.0,
            "tis": 0.0,
            "isp": 4792.0,
            "total": 4792.0
        },
        {
            "group": "Licence & legal fee",
            "particulars": "RJSC Expenses",
            "account_code": "IS-043",
            "tbn": 2083.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 2083.0
        },
        {
            "group": "Licence & legal fee",
            "particulars": "Trade Licence",
            "account_code": "IS-083",
            "tbn": 4333.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 4333.0
        }
    ],
    "direct_subtotal": {
        "tbn": 2567351.0,
        "tis": 1427787.0,
        "isp": 295428.0,
        "total": 4290566.0
    },
    "proportionate_expenses": [
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Office Expencess",
            "account_code": "IS-016",
            "tbn": 26000.0,
            "tis": 10500.0,
            "isp": 1500.0,
            "total": 38000.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Cleaner wages",
            "account_code": "IS-066",
            "tbn": 4550.0,
            "tis": 1838.0,
            "isp": 263.0,
            "total": 6650.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Advertisement",
            "account_code": "IS-039",
            "tbn": 5850.0,
            "tis": 2363.0,
            "isp": 338.0,
            "total": 8550.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Printing Expnepses",
            "account_code": "IS-019",
            "tbn": 1950.0,
            "tis": 788.0,
            "isp": 113.0,
            "total": 2850.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Bank Charge",
            "account_code": "IS-020",
            "tbn": 650.0,
            "tis": 263.0,
            "isp": 38.0,
            "total": 950.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Electric Accessories",
            "account_code": "IS-017",
            "tbn": 3900.0,
            "tis": 1575.0,
            "isp": 225.0,
            "total": 5700.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Printer Tonner",
            "account_code": "IS-008",
            "tbn": 3250.0,
            "tis": 1313.0,
            "isp": 188.0,
            "total": 4750.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Entertainment-office",
            "account_code": "IS-006",
            "tbn": 22750.0,
            "tis": 9188.0,
            "isp": 1313.0,
            "total": 33250.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Entertainment-Night",
            "account_code": "IS-102",
            "tbn": 14950.0,
            "tis": 8038.0,
            "isp": 863.0,
            "total": 23850.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Postage and currier",
            "account_code": "IS-018",
            "tbn": 405.0,
            "tis": 164.0,
            "isp": 23.0,
            "total": 592.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Stamp Duty",
            "account_code": "IS-083",
            "tbn": 1209.0,
            "tis": 488.0,
            "isp": 70.0,
            "total": 1767.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Stationary",
            "account_code": "IS-008",
            "tbn": 3400.0,
            "tis": 1575.0,
            "isp": 225.0,
            "total": 5200.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Conveyance",
            "account_code": "IS-014",
            "tbn": 3250.0,
            "tis": 1313.0,
            "isp": 188.0,
            "total": 4750.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Mobile Bill",
            "account_code": "IS-023",
            "tbn": 3250.0,
            "tis": 1313.0,
            "isp": 188.0,
            "total": 4750.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Device Recharge",
            "account_code": "IS-111",
            "tbn": 1979.0,
            "tis": 799.0,
            "isp": 114.0,
            "total": 2893.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Internet Expenses",
            "account_code": "IS-025",
            "tbn": 2925.0,
            "tis": 1181.0,
            "isp": 169.0,
            "total": 4275.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Fuel for Generator",
            "account_code": "IS-032",
            "tbn": 8450.0,
            "tis": 3413.0,
            "isp": 488.0,
            "total": 12350.0
        },
        {
            "group": "Office & Administrative Expenses",
            "particulars": "Computer Accessories",
            "account_code": "IS-041",
            "tbn": 7750.0,
            "tis": 3938.0,
            "isp": 563.0,
            "total": 12250.0
        },
        {
            "group": "TDS VDS",
            "particulars": "TDS Payment",
            "account_code": "IS-075",
            "tbn": 5978.0,
            "tis": 2414.0,
            "isp": 345.0,
            "total": 8737.0
        },
        {
            "group": "TDS VDS",
            "particulars": "Vat Expenses",
            "account_code": "IS-064",
            "tbn": 7970.0,
            "tis": 3219.0,
            "isp": 460.0,
            "total": 11649.0
        },
        {
            "group": "Transportation Expenses",
            "particulars": "Transport(Car rent)",
            "account_code": "IS-011",
            "tbn": 49600.0,
            "tis": 49600.0,
            "isp": 0.0,
            "total": 99200.0
        },
        {
            "group": "Transportation Expenses",
            "particulars": "Fuel for Vehicle",
            "account_code": "IS-024",
            "tbn": 56950.0,
            "tis": 21788.0,
            "isp": 3113.0,
            "total": 81850.0
        },
        {
            "group": "Transportation Expenses",
            "particulars": "TDS & VDS",
            "account_code": "IS-083",
            "tbn": 4960.0,
            "tis": 0.0,
            "isp": 0.0,
            "total": 4960.0
        },
        {
            "group": "Repair & Maintenance",
            "particulars": "Repair & Maintenance",
            "account_code": "IS-017",
            "tbn": 13250.0,
            "tis": 6563.0,
            "isp": 938.0,
            "total": 20750.0
        },
        {
            "group": "Repair & Maintenance",
            "particulars": "Repair & Maintenance (AC)",
            "account_code": "IS-017",
            "tbn": 7750.0,
            "tis": 3938.0,
            "isp": 563.0,
            "total": 12250.0
        },
        {
            "group": "ISP Expenses",
            "particulars": "ISP Expenses",
            "account_code": "IS-086",
            "tbn": 69516.0,
            "tis": 301238.0,
            "isp": 139975.0,
            "total": 510729.0
        }
    ],
    "proportionate_subtotal": {
        "tbn": 332443.0,
        "tis": 438803.0,
        "isp": 152256.0,
        "total": 923502.0
    },
    "grand_total": {
        "tbn": 2899794.0,
        "tis": 1866590.0,
        "isp": 447684.0,
        "total": 5214068.0
    },
    "additional_expenses": [
        {
            "particulars": "Incentive for Sale (July)",
            "account_code": "IS-046",
            "tbn": 0.0,
            "tis": 92450.0,
            "isp": 0.0,
            "total": 92450.0
        }
    ],
    "additional_subtotal": 92450.0,
    "gross_total": 5306518.0,
    "bill_adjustment": 187000.0,
    "bill_adjustment_note": "Adjusted with ISP Bill Received",
    "net_budget": 5119518.0,
    "usd_rate": 121.75,
    "usd_amount": 42049.0
}

def seed_september_exact_budget():
    with app.app_context():
        # Map codes to account objects
        accounts_by_code = {acc.code: acc for acc in Account.query.all()}
        
        # Check or retrieve September budget
        mb = MonthlyBudget.query.filter_by(month_year='2026-09').first()
        if not mb:
            mb = MonthlyBudget(
                month_year='2026-09',
                title='Corporate Budget for September 2026',
                fiscal_year='2026-2027',
                status='Issued',
                issued_by='Admin'
            )
            db.session.add(mb)
            db.session.flush()

        mb.title = SEPTEMBER_2026_BUDGET_DATA["title"]
        mb.total_budgeted = SEPTEMBER_2026_BUDGET_DATA["grand_total"]["total"]
        mb.tbn_budget = SEPTEMBER_2026_BUDGET_DATA["grand_total"]["tbn"]
        mb.tis_budget = SEPTEMBER_2026_BUDGET_DATA["grand_total"]["tis"]
        mb.isp_budget = SEPTEMBER_2026_BUDGET_DATA["grand_total"]["isp"]
        mb.tbn_headcount = SEPTEMBER_2026_BUDGET_DATA["headcounts"]["tbn"]
        mb.tis_headcount = SEPTEMBER_2026_BUDGET_DATA["headcounts"]["tis"]
        mb.isp_headcount = SEPTEMBER_2026_BUDGET_DATA["headcounts"]["isp"]
        mb.additional_expenses = SEPTEMBER_2026_BUDGET_DATA["additional_subtotal"]
        mb.additional_expenses_note = "Incentive for Sale (July)"
        mb.gross_total = SEPTEMBER_2026_BUDGET_DATA["gross_total"]
        mb.bill_adjustment = SEPTEMBER_2026_BUDGET_DATA["bill_adjustment"]
        mb.bill_adjustment_note = SEPTEMBER_2026_BUDGET_DATA["bill_adjustment_note"]
        mb.net_budget = SEPTEMBER_2026_BUDGET_DATA["net_budget"]
        mb.usd_rate = SEPTEMBER_2026_BUDGET_DATA["usd_rate"]
        mb.usd_amount = SEPTEMBER_2026_BUDGET_DATA["usd_amount"]
        mb.budget_data_json = json.dumps(SEPTEMBER_2026_BUDGET_DATA)

        # Remove existing items for this budget to rebuild with exact multi-division fields
        BudgetItem.query.filter_by(budget_id=mb.id).delete()

        # Add Direct Expenses
        for item in SEPTEMBER_2026_BUDGET_DATA["direct_expenses"]:
            acc = accounts_by_code.get(item["account_code"])
            bi = BudgetItem(
                budget_id=mb.id,
                account_id=acc.id if acc else None,
                particulars=item["particulars"],
                section="direct",
                category_group=item["group"],
                tbn_amount=item["tbn"],
                tis_amount=item["tis"],
                isp_amount=item["isp"],
                allocated_amount=item["total"],
                is_proportionate=False
            )
            db.session.add(bi)

        # Add Proportionate Expenses
        for item in SEPTEMBER_2026_BUDGET_DATA["proportionate_expenses"]:
            acc = accounts_by_code.get(item["account_code"])
            bi = BudgetItem(
                budget_id=mb.id,
                account_id=acc.id if acc else None,
                particulars=item["particulars"],
                section="proportionate",
                category_group=item["group"],
                tbn_amount=item["tbn"],
                tis_amount=item["tis"],
                isp_amount=item["isp"],
                allocated_amount=item["total"],
                is_proportionate=True
            )
            db.session.add(bi)

        # Add Additional Expenses
        for item in SEPTEMBER_2026_BUDGET_DATA["additional_expenses"]:
            acc = accounts_by_code.get(item["account_code"])
            bi = BudgetItem(
                budget_id=mb.id,
                account_id=acc.id if acc else None,
                particulars=item["particulars"],
                section="additional",
                category_group="Additional Expenses",
                tbn_amount=item["tbn"],
                tis_amount=item["tis"],
                isp_amount=item["isp"],
                allocated_amount=item["total"],
                is_proportionate=False
            )
            db.session.add(bi)

        db.session.commit()
        print(" * Successfully seeded September 2026 corporate budget format!")

if __name__ == '__main__':
    seed_september_exact_budget()
