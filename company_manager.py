import os
import json
import sqlite3
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import create_engine

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
COMPANIES_DIR = os.path.join(BASE_DIR, 'companies')
PROFILES_FILE = os.path.join(BASE_DIR, 'company_profiles.json')

# Table groupings for selective data transfer
MODULE_TABLES = {
    'accounts': ['account'],
    'vouchers': ['journal_entry', 'transaction_line', 'cash_book_row'],
    'employees': ['employee', 'payroll_record', 'payroll_item', 'employee_fund_contribution', 'employee_fund_loan', 'employee_fund_settlement', 'employee_salary_history', 'employee_noc_record', 'employee_confirmation_letter', 'employee_offer_letter', 'employee_salary_certificate_record', 'employee_experience_certificate_record', 'employee_tax_challan', 'employee_clearance_record'],
    'banking': ['bank_statement', 'bank_statement_line', 'bank_reconciliation'],
    'budgets': ['monthly_budget', 'budget_item'],
    'remittances': ['remittance_record'],
    'invoices': ['service_invoice'],
    'users': ['user', 'email_otp']
}

# Ordered tables for proper foreign-key insertion order
ALL_TABLES_ORDERED = [
    'user',
    'email_otp',
    'account',
    'employee',
    'bank_statement',
    'bank_statement_line',
    'bank_reconciliation',
    'monthly_budget',
    'budget_item',
    'service_invoice',
    'remittance_record',
    'journal_entry',
    'transaction_line',
    'cash_book_row',
    'payroll_record',
    'payroll_item',
    'employee_fund_contribution',
    'employee_fund_loan',
    'employee_fund_settlement',
    'employee_salary_history',
    'employee_noc_record',
    'employee_confirmation_letter',
    'employee_offer_letter',
    'employee_salary_certificate_record',
    'employee_experience_certificate_record',
    'employee_tax_challan',
    'employee_clearance_record'
]

class CompanyProfile:
    def __init__(self, id, name, code, db_file, is_locked=False, password_hash=None, description="", created_at=None):
        self.id = id
        self.name = name
        self.code = code
        self.db_file = db_file
        self.is_locked = is_locked
        self.password_hash = password_hash
        self.description = description or ""
        self.created_at = created_at or datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')

    @property
    def db_path(self):
        if os.path.isabs(self.db_file):
            return self.db_file
        return os.path.join(BASE_DIR, self.db_file)

    def check_password(self, password):
        if not self.is_locked or not self.password_hash:
            return True
        return check_password_hash(self.password_hash, password)

    def set_password(self, password):
        if password:
            self.password_hash = generate_password_hash(password)
            self.is_locked = True
        else:
            self.password_hash = None
            self.is_locked = False

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'code': self.code,
            'db_file': self.db_file,
            'is_locked': self.is_locked,
            'password_hash': self.password_hash,
            'description': self.description,
            'created_at': self.created_at
        }


class CompanyManager:
    _engines = {}

    @classmethod
    def ensure_storage(cls):
        os.makedirs(COMPANIES_DIR, exist_ok=True)
        if not os.path.exists(PROFILES_FILE):
            default_profile = {
                'id': 1,
                'name': 'ZR-ERP Main Enterprise',
                'code': 'ZR-01',
                'db_file': 'accounting.db',
                'is_locked': False,
                'password_hash': None,
                'description': 'Primary Corporate Financial & Accounting Ledger',
                'created_at': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(PROFILES_FILE, 'w', encoding='utf-8') as f:
                json.dump([default_profile], f, indent=2)

    @classmethod
    def load_profiles(cls):
        cls.ensure_storage()
        try:
            with open(PROFILES_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return [CompanyProfile(**item) for item in data]
        except Exception as e:
            print(f"Error loading company profiles: {e}")
            return [CompanyProfile(1, 'ZR-ERP Main Enterprise', 'ZR-01', 'accounting.db')]

    @classmethod
    def save_profiles(cls, profiles):
        cls.ensure_storage()
        data = [p.to_dict() for p in profiles]
        with open(PROFILES_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)

    @classmethod
    def get_all(cls):
        return cls.load_profiles()

    @classmethod
    def get_by_id(cls, cid):
        profiles = cls.load_profiles()
        for p in profiles:
            if p.id == int(cid):
                return p
        return profiles[0] if profiles else None

    @classmethod
    def get_engine(cls, cid, metadata=None):
        return cls.get_engine_for_company(cid, metadata=metadata)

    @classmethod
    def get_engine_for_company(cls, cid, metadata=None):
        company = cls.get_by_id(cid)
        if not company:
            company = cls.get_by_id(1)

        db_path = company.db_path
        if cid not in cls._engines:
            engine = create_engine(
                f'sqlite:///{db_path}',
                connect_args={'timeout': 60, 'check_same_thread': False}
            )
            # Enable WAL mode & high-performance SQLite PRAGMAs
            with engine.connect() as conn:
                conn.exec_driver_sql("PRAGMA journal_mode=WAL;")
                conn.exec_driver_sql("PRAGMA busy_timeout=60000;")
                conn.exec_driver_sql("PRAGMA synchronous=NORMAL;")
                conn.exec_driver_sql("PRAGMA cache_size=-64000;")
                conn.exec_driver_sql("PRAGMA temp_store=MEMORY;")
                conn.exec_driver_sql("PRAGMA mmap_size=268435456;")
            cls._engines[cid] = engine

        engine = cls._engines[cid]
        if metadata is not None:
            # Ensure schema tables exist
            with engine.begin() as conn:
                metadata.create_all(conn)
        cls._ensure_schema_columns(engine)
        return engine

    @classmethod
    def _ensure_schema_columns(cls, engine):
        try:
            with engine.begin() as conn:
                # Check employee table columns
                cur = conn.exec_driver_sql("PRAGMA table_info(employee);").fetchall()
                if cur:
                    emp_cols = {r[1] for r in cur}
                    needed_emp = {
                        'bank_name': "VARCHAR(100) DEFAULT NULL",
                        'bank_branch': "VARCHAR(100) DEFAULT NULL",
                        'routing_no': "VARCHAR(20) DEFAULT NULL",
                        'payment_method': "VARCHAR(20) DEFAULT 'Cash'",
                        'pf_status': "VARCHAR(10) DEFAULT 'P'",
                        'tds_amount': "FLOAT DEFAULT 0.0",
                        'employment_type': "VARCHAR(20) DEFAULT 'Full-Time'",
                        'duty_category': "VARCHAR(50) DEFAULT NULL"
                    }
                    for col, col_def in needed_emp.items():
                        if col not in emp_cols:
                            conn.exec_driver_sql(f"ALTER TABLE employee ADD COLUMN {col} {col_def};")

                # Check monthly_salary_item table columns
                cur = conn.exec_driver_sql("PRAGMA table_info(monthly_salary_item);").fetchall()
                if cur:
                    msi_cols = {r[1] for r in cur}
                    needed_msi = {
                        'status': "VARCHAR(20) DEFAULT 'Running'",
                        'employment_type': "VARCHAR(20) DEFAULT 'Full-Time'",
                        'duty_category': "VARCHAR(50) DEFAULT NULL",
                        'shift_eve_duty': "FLOAT DEFAULT 0.0",
                        'shift_eve_rate': "FLOAT DEFAULT 800.0",
                        'shift_night_duty': "FLOAT DEFAULT 0.0",
                        'shift_night_rate': "FLOAT DEFAULT 1400.0",
                        'shift_morn_duty': "FLOAT DEFAULT 0.0",
                        'shift_morn_rate': "FLOAT DEFAULT 1500.0",
                        'shift_aft_duty': "FLOAT DEFAULT 0.0",
                        'shift_aft_rate': "FLOAT DEFAULT 700.0",
                        'total_duty': "FLOAT DEFAULT 0.0"
                    }
                    for col, col_def in needed_msi.items():
                        if col not in msi_cols:
                            conn.exec_driver_sql(f"ALTER TABLE monthly_salary_item ADD COLUMN {col} {col_def};")

                # Check journal_entry table columns
                cur = conn.exec_driver_sql("PRAGMA table_info(journal_entry);").fetchall()
                if cur:
                    je_cols = {r[1] for r in cur}
                    needed_je = {
                        'attachment': "VARCHAR(255) DEFAULT NULL",
                        'original_filename': "VARCHAR(255) DEFAULT NULL",
                        'drive_file_id': "VARCHAR(100) DEFAULT NULL",
                        'drive_view_link': "TEXT DEFAULT NULL"
                    }
                    for col, col_def in needed_je.items():
                        if col not in je_cols:
                            conn.exec_driver_sql(f"ALTER TABLE journal_entry ADD COLUMN {col} {col_def};")

                # Check employee_document table columns
                cur = conn.exec_driver_sql("PRAGMA table_info(employee_document);").fetchall()
                if cur:
                    ed_cols = {r[1] for r in cur}
                    needed_ed = {
                        'emp_code': "VARCHAR(50) DEFAULT NULL",
                        'file_type': "VARCHAR(50) DEFAULT NULL",
                        'file_path': "VARCHAR(255) DEFAULT NULL",
                        'upload_date': "TIMESTAMP DEFAULT NULL",
                        'drive_file_id': "VARCHAR(100) DEFAULT NULL",
                        'drive_view_link': "TEXT DEFAULT NULL"
                    }
                    for col, col_def in needed_ed.items():
                        if col not in ed_cols:
                            conn.exec_driver_sql(f"ALTER TABLE employee_document ADD COLUMN {col} {col_def};")

                # Check user table columns (email and phone)
                cur = conn.exec_driver_sql("PRAGMA table_info(user);").fetchall()
                if cur:
                    user_cols = {r[1] for r in cur}
                    if 'email' not in user_cols:
                        conn.exec_driver_sql("ALTER TABLE user ADD COLUMN email VARCHAR(120);")
                    if 'phone' not in user_cols:
                        conn.exec_driver_sql("ALTER TABLE user ADD COLUMN phone VARCHAR(30);")
        except Exception:
            pass

    @classmethod
    def create_company(cls, name, code, description="", is_locked=False, password="", metadata=None, clone_from_id=None, clone_mode="full", selected_modules=None):
        profiles = cls.load_profiles()
        new_id = max((p.id for p in profiles), default=0) + 1
        safe_code = code.strip().replace(" ", "_").upper()
        rel_db_file = os.path.join('companies', f'company_{new_id}_{safe_code}.db').replace('\\', '/')
        abs_db_path = os.path.join(BASE_DIR, rel_db_file)

        new_company = CompanyProfile(
            id=new_id,
            name=name.strip(),
            code=safe_code,
            db_file=rel_db_file,
            is_locked=is_locked,
            description=description.strip()
        )
        if is_locked and password:
            new_company.set_password(password)

        profiles.append(new_company)
        cls.save_profiles(profiles)

        # Initialize schema in new database
        engine = cls.get_engine_for_company(new_id, metadata=metadata)

        # If cloning requested from an existing company
        if clone_from_id:
            cls.transfer_data(
                source_cid=clone_from_id,
                dest_cid=new_id,
                mode=clone_mode,
                selected_modules=selected_modules or [],
                strategy='replace',
                metadata=metadata
            )
        else:
            # Seed users table from company 1 by default so administrator accounts work seamlessly
            try:
                cls.transfer_data(
                    source_cid=1,
                    dest_cid=new_id,
                    mode='selected',
                    selected_modules=['users'],
                    strategy='replace',
                    metadata=metadata
                )
            except Exception as e:
                print(f"Note: Could not copy initial users to company {new_id}: {e}")

        return new_company

    @classmethod
    def update_company(cls, cid, name, code, description="", is_locked=False, new_password=""):
        profiles = cls.load_profiles()
        target = None
        for p in profiles:
            if p.id == int(cid):
                target = p
                break

        if not target:
            raise ValueError(f"Company ID {cid} not found.")

        target.name = name.strip()
        target.code = code.strip().upper()
        target.description = description.strip()

        if is_locked:
            target.is_locked = True
            if new_password:
                target.set_password(new_password)
        else:
            target.is_locked = False
            target.password_hash = None

        cls.save_profiles(profiles)
        return target

    @classmethod
    def delete_company(cls, cid):
        cid = int(cid)
        if cid == 1:
            raise ValueError("The primary company profile cannot be deleted.")

        profiles = cls.load_profiles()
        target = next((p for p in profiles if p.id == cid), None)
        if not target:
            return False

        # Remove engine from cache if open
        if cid in cls._engines:
            cls._engines[cid].dispose()
            del cls._engines[cid]

        # Delete database file
        if os.path.exists(target.db_path):
            try:
                os.remove(target.db_path)
            except Exception as e:
                print(f"Could not remove DB file {target.db_path}: {e}")

        profiles = [p for p in profiles if p.id != cid]
        cls.save_profiles(profiles)
        return True

    @classmethod
    def transfer_data(cls, source_cid, dest_cid, mode="full", selected_modules=None, strategy="replace", metadata=None):
        """
        Transfers full or selected data from source company to destination company.
        mode: 'full' or 'selected'
        selected_modules: list of keys in MODULE_TABLES (e.g. ['accounts', 'vouchers', ...])
        strategy: 'replace' (clears dest tables first) or 'merge' (INSERT OR IGNORE)
        """
        source_comp = cls.get_by_id(source_cid)
        dest_comp = cls.get_by_id(dest_cid)

        if not source_comp or not dest_comp:
            raise ValueError("Invalid source or destination company ID.")
        if source_comp.id == dest_comp.id:
            raise ValueError("Source and destination company profiles must be different.")

        # Ensure destination tables exist
        cls.get_engine_for_company(dest_cid, metadata=metadata)

        # Determine tables to transfer
        if mode == 'full':
            tables_to_transfer = [t for t in ALL_TABLES_ORDERED]
        else:
            tables_set = set()
            for mod in (selected_modules or []):
                for t in MODULE_TABLES.get(mod, []):
                    tables_set.add(t)
            # Maintain proper foreign key order
            tables_to_transfer = [t for t in ALL_TABLES_ORDERED if t in tables_set]

        if not tables_to_transfer:
            return {'success': False, 'message': 'No modules selected for transfer.', 'row_counts': {}}

        source_db_path = source_comp.db_path
        dest_db_path = dest_comp.db_path

        # Connect directly to destination SQLite DB
        conn = sqlite3.connect(dest_db_path, timeout=60)
        cursor = conn.cursor()
        cursor.execute("PRAGMA foreign_keys = OFF;")
        cursor.execute("ATTACH DATABASE ? AS src;", (source_db_path,))

        transferred_counts = {}

        try:
            for table in tables_to_transfer:
                # Check if table exists in source
                cursor.execute(f"SELECT count(*) FROM src.sqlite_master WHERE type='table' AND name=?;", (table,))
                if cursor.fetchone()[0] == 0:
                    continue

                # Query matching column names between src and destination to ensure exact mapping
                cursor.execute(f"PRAGMA src.table_info({table});")
                src_cols = [r[1] for r in cursor.fetchall()]
                cursor.execute(f"PRAGMA main.table_info({table});")
                dest_cols = [r[1] for r in cursor.fetchall()]
                common_cols = [c for c in src_cols if c in dest_cols]
                if not common_cols:
                    continue

                cols_clause = ", ".join(f'"{c}"' for c in common_cols)

                if strategy == 'replace':
                    cursor.execute(f"DELETE FROM main.{table};")
                    sql = f"INSERT OR REPLACE INTO main.{table} ({cols_clause}) SELECT {cols_clause} FROM src.{table};"
                else:
                    sql = f"INSERT OR IGNORE INTO main.{table} ({cols_clause}) SELECT {cols_clause} FROM src.{table};"

                cursor.execute(sql)
                # Count affected rows in destination
                cursor.execute(f"SELECT count(*) FROM main.{table};")
                transferred_counts[table] = cursor.fetchone()[0]

            conn.commit()
            cursor.execute("DETACH DATABASE src;")
            cursor.execute("PRAGMA foreign_keys = ON;")
            conn.close()

            return {
                'success': True,
                'source': source_comp.name,
                'destination': dest_comp.name,
                'tables': tables_to_transfer,
                'row_counts': transferred_counts,
                'message': f"Data successfully transferred from '{source_comp.name}' to '{dest_comp.name}'."
            }

        except Exception as e:
            conn.rollback()
            try:
                cursor.execute("DETACH DATABASE src;")
            except Exception:
                pass
            conn.close()
            raise RuntimeError(f"Data transfer failed: {str(e)}")
