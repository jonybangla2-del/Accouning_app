"""
Employee Document Management Engine
Module: employee_doc_manager.py

Handles:
1. Folder scanning & Employee ID detection from folder structures
2. Supported formats: .pdf, .jpg, .jpeg, .png, .docx, .doc, .xlsx, .xls
3. Pre-upload Confirmation & selective attachment
4. Secure central storage under app_storage/employee_docs/{emp_id}/
5. Streaming downloads & browser views
6. 1-Click Bulk Delete of all attached documents
"""

import os
import re
import shutil
import io
import uuid
import logging
from datetime import datetime
from werkzeug.utils import secure_filename
from PIL import Image, ImageOps

logger = logging.getLogger('EmployeeDocOptimizer')
if not logger.handlers:
    logging.basicConfig(level=logging.INFO)

def format_bytes_human(num_bytes):
    if num_bytes is None:
        return "0 B"
    b = float(num_bytes)
    if b < 1024:
        return f"{int(b)} B"
    elif b < 1024 * 1024:
        return f"{b / 1024:.1f} KB"
    else:
        return f"{b / (1024 * 1024):.2f} MB"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
STORAGE_ROOT = os.path.join(BASE_DIR, 'app_storage', 'employee_docs')


ALLOWED_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png', '.docx', '.doc', '.xlsx', '.xls'}

EXTENSION_CATEGORIES = {
    '.pdf': 'PDF',
    '.jpg': 'IMAGE',
    '.jpeg': 'IMAGE',
    '.png': 'IMAGE',
    '.docx': 'DOCX',
    '.doc': 'DOCX',
    '.xlsx': 'EXCEL',
    '.xls': 'EXCEL'
}

CATEGORY_ICONS = {
    'PDF': '📕',
    'IMAGE': '🖼️',
    'DOCX': '📘',
    'EXCEL': '📗',
    'OTHER': '📄'
}

PRESET_CONFIGS = {
    'balanced': {
        'name': 'Standard / Balanced',
        'badge': '1080p Balanced (<1MB)',
        'max_width': 1080,
        'quality': 75,
        'pdf_dpi': 150,
        'target_pdf_kb': 1024
    },
    'high': {
        'name': 'High Quality',
        'badge': '1920p High Quality (<2MB)',
        'max_width': 1920,
        'quality': 85,
        'pdf_dpi': 200,
        'target_pdf_kb': 2048
    },
    'small': {
        'name': 'Small Size',
        'badge': '720p Small Size (<500KB)',
        'max_width': 720,
        'quality': 60,
        'pdf_dpi': 100,
        'target_pdf_kb': 500
    },
    'original': {
        'name': 'Original (No Compression)',
        'badge': 'Original File',
        'max_width': None,
        'quality': None,
        'pdf_dpi': None,
        'target_pdf_kb': None
    }
}


class EmployeeDocManager:
    def __init__(self, storage_dir=None):
        self.storage_dir = storage_dir or STORAGE_ROOT
        os.makedirs(self.storage_dir, exist_ok=True)

    def is_allowed_file(self, filename):
        if not filename or '.' not in filename:
            return False
        ext = os.path.splitext(filename)[1].lower()
        return ext in ALLOWED_EXTENSIONS

    def get_file_category(self, filename_or_ext):
        ext = filename_or_ext.lower()
        if not ext.startswith('.'):
            ext = os.path.splitext(ext)[1]
        return EXTENSION_CATEGORIES.get(ext, 'OTHER')

    def get_file_icon(self, category):
        return CATEGORY_ICONS.get(category, '📄')

    GENERIC_FOLDER_NAMES = {
        'EMPLOYEE', 'EMPLOYEES', 'ALL', 'ALL_EMPLOYEES', 'ALL EMPLOYEES',
        'DOCS', 'DOCUMENTS', 'EMPLOYEE_DOCS', 'EMPLOYEE DOCUMENTS', 'FILES',
        'SCANS', 'SCAN', 'UPLOADS', 'ATTACHMENTS', 'DATA', 'DESKTOP',
        'DOWNLOADS', 'TEMP', 'NEW FOLDER', 'FOLDER', 'BACKUP', 'RECORDS',
        'MYDOCS', 'MY_DOCS', 'STAFF', 'PERSONNEL', 'ROOT'
    }

    DOC_NAME_KEYWORDS = {
        'offer', 'letter', 'appointment', 'confirmation', 'agreement', 'contract',
        'resume', 'cv', 'biodata', 'nid', 'passport', 'photo', 'certificate',
        'authorized', 'chaque', 'cheque', 'book', 'bank', 'slip', 'copy', 'scan',
        'tbn24', 'tbn', 'ltd', 'doc', 'docx', 'pdf', 'jpg', 'png', 'image',
        'for', 'of', 'and', 'the', 'to', 'in', 'at', 'on', 'by', 'with',
        'clearance', 'resignation', 'salary', 'experience', 'noc', 'tax',
        'challan', 'tin', 'ssc', 'hsc', 'degree', 'masters', 'bba', 'mba',
        'transcript', 'marksheet', 'training', 'official', 'confidential',
        'document', 'documents', 'file', 'files', 'page', 'signed', 'final',
        'new', 'old', 'updated', 'profile', 'employee', 'id'
    }

    def extract_candidate_name_from_filenames(self, filenames):
        """
        Parses filenames in an employee's folder to extract candidate full name.
        Example: 'Farhana Rahman Mitu - Offer Letter.docx' -> 'Farhana Rahman Mitu'
        """
        if not filenames:
            return None
        for fn in filenames:
            if not fn:
                continue
            base = os.path.basename(str(fn)).rsplit('.', 1)[0]
            # Split by hyphens, underscores, en/em-dashes, or brackets
            parts = re.split(r'[-–—_]+', base)
            for p in parts:
                cleaned = re.sub(r'[\(\)\[\]]', ' ', p).strip()
                words = cleaned.split()
                valid_words = [w for w in words if w.lower() not in self.DOC_NAME_KEYWORDS and not w.isdigit()]
                if len(valid_words) >= 2:
                    candidate = ' '.join(valid_words).title()
                    # Sanity check: reasonable name length
                    if 3 <= len(candidate) <= 45:
                        return candidate
        return None

    def extract_employee_metadata_from_files(self, files_data, is_local_disk=False, source_root=None):
        """
        Reads files (filenames + document text content) to extract structured employee details:
        name, designation, department, salary, phone, email, joining_date.
        """
        extracted = {
            'name': None,
            'designation': None,
            'department': None,
            'salary': 0.0,
            'phone': None,
            'email': None,
            'joining_date': None
        }

        # 1. First pass: extract candidate name from filenames
        fnames = []
        if is_local_disk:
            fnames = [f.get('filename') or os.path.basename(f.get('rel_path', '')) for f in files_data if isinstance(f, dict)]
        else:
            for item in files_data:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    fnames.append(item[1])
                elif isinstance(item, dict):
                    fnames.append(item.get('filename', ''))
                elif hasattr(item, 'filename'):
                    fnames.append(item.filename)

        cand_name = self.extract_candidate_name_from_filenames(fnames)
        if cand_name:
            extracted['name'] = cand_name

        # 2. Second pass: deep document parsing for docx, pdf, txt, doc
        try:
            from document_parser import DocumentParser
            doc_parser = DocumentParser(max_file_size_mb=25)
        except Exception:
            doc_parser = None

        if doc_parser:
            for item in files_data:
                orig_name = ''
                file_source = None

                if is_local_disk and isinstance(item, dict):
                    orig_name = item.get('filename') or os.path.basename(item.get('rel_path', ''))
                    rel_p = item.get('rel_path') or orig_name
                    abs_p = item.get('abs_path') or (os.path.join(source_root, rel_p) if source_root else None)
                    if abs_p and os.path.exists(abs_p):
                        file_source = abs_p
                elif isinstance(item, (list, tuple)) and len(item) >= 2:
                    file_obj = item[0]
                    orig_name = item[1]
                    if hasattr(file_obj, 'read'):
                        try:
                            file_obj.seek(0)
                            file_source = file_obj.read()
                            file_obj.seek(0)
                        except Exception:
                            file_source = None

                if not file_source or not orig_name:
                    continue

                ext = orig_name.rsplit('.', 1)[-1].lower() if '.' in orig_name else ''
                if ext in ('docx', 'doc', 'pdf', 'txt'):
                    try:
                        parse_res = doc_parser.parse_file(file_source, orig_name)
                        text = parse_res.get('text', '').strip()
                        if text:
                            lines = [l.strip() for l in text.split('\n') if l.strip()]

                            # Check 'To,' format for recipient name
                            if not extracted['name']:
                                for idx, l in enumerate(lines):
                                    if l.lower() in ('to,', 'to:', 'to') and idx + 1 < len(lines):
                                        c_line = lines[idx + 1].strip(' ,.-')
                                        words = c_line.split()
                                        if 2 <= len(words) <= 5 and not any(w in c_line.lower() for w in ['ltd', 'limited', 'company', 'channel', 'television', 'hr']):
                                            extracted['name'] = c_line.title()
                                            break

                            # Check 'Dear Name' format
                            if not extracted['name']:
                                m_dear = re.search(r'Dear\s+([A-Z][a-zA-Z\.\s]{2,35})[,:]', text)
                                if m_dear:
                                    extracted['name'] = m_dear.group(1).strip().title()

                            # Check Name: format
                            if not extracted['name']:
                                m_name = re.search(r'(?:Employee Name|Candidate Name|Name)\s*[:\-–—]\s*([A-Z][a-zA-Z\.\s]{2,35})', text)
                                if m_name:
                                    extracted['name'] = m_name.group(1).strip().title()

                            # Designation
                            if not extracted['designation']:
                                m_desig = re.search(r'(?:position of|Designation\s*[:\-–—]|Position\s*[:\-–—]|appointed as)\s+([A-Za-z\s&/\-]{3,35})(?:\s+in|\s+at|,|\.|\n)', text, re.I)
                                if m_desig:
                                    extracted['designation'] = m_desig.group(1).strip(' ,.-').title()

                            # Department
                            if not extracted['department']:
                                m_dept = re.search(r'(?:in our|Department\s*[:\-–—]|Dept\s*[:\-–—])\s+([A-Za-z\s&/\-]{3,40})(?:\s+Department|\s+Dept|,|\.|\n)', text, re.I)
                                if m_dept:
                                    d_cand = m_dept.group(1).strip(' ,.-').title()
                                    if 'Department' not in d_cand:
                                        d_cand += ' Department'
                                    extracted['department'] = d_cand

                            # Salary
                            if not extracted['salary'] or extracted['salary'] == 0.0:
                                m_sal = re.search(r'(?:gross salary|basic salary|monthly salary|remuneration)\s*(?:will be|is|[:\-–—])?\s*(?:BDT|Tk\.?|৳)?\s*([0-9,]+)', text, re.I)
                                if m_sal:
                                    try:
                                        extracted['salary'] = float(m_sal.group(1).replace(',', ''))
                                    except Exception:
                                        pass

                            # Phone
                            if not extracted['phone']:
                                m_ph = re.search(r'(?:\+?880\s?|0)1[3-9]\d{8}', text)
                                if m_ph:
                                    extracted['phone'] = m_ph.group(0).strip()

                            # Email
                            if not extracted['email']:
                                m_em = re.search(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', text)
                                if m_em:
                                    extracted['email'] = m_em.group(0).strip()
                    except Exception as parse_err:
                        logger.warning(f"[DocOptimizer] Content parsing note for {orig_name}: {parse_err}")

        return extracted

    def get_or_create_employee_from_files(self, db_session, emp_code, files_data, is_local_disk=False, source_root=None):
        """
        Locates existing Employee by ID or creates a new one from extracted file data.
        Guarantees 100% duplicate protection: never duplicates an employee ID.
        Returns: (Employee_obj, is_new_bool)
        """
        from app import Employee
        from sqlalchemy import func

        clean_code = str(emp_code).strip()
        clean_upper = clean_code.upper()

        # 1. Lookup existing employee
        existing = Employee.query.filter(
            func.lower(func.trim(Employee.employee_id)) == clean_upper.lower()
        ).first()

        if not existing and clean_code.isdigit():
            num = int(clean_code)
            existing = Employee.query.filter(
                func.trim(Employee.employee_id) == str(num)
            ).first()
            if not existing:
                for pad in (2, 3, 4, 5):
                    existing = Employee.query.filter(
                        func.trim(Employee.employee_id) == f"{num:0{pad}d}"
                    ).first()
                    if existing:
                        break
            if not existing:
                for prefix in ("EMP-", "EMP_", "EMP", "ID-", "ID"):
                    existing = Employee.query.filter(
                        func.lower(func.trim(Employee.employee_id)) == f"{prefix}{clean_code}".lower()
                    ).first()
                    if existing:
                        break

        # 2. Extract metadata from files
        extracted = self.extract_employee_metadata_from_files(files_data, is_local_disk=is_local_disk, source_root=source_root)

        if existing:
            # DUPLICATE PROTECTION: Existing employee found! Do not create duplicate!
            # If existing employee had a placeholder name, upgrade it with parsed real name
            if extracted.get('name'):
                c_name = extracted['name'].strip()
                if (existing.name.lower().startswith('employee ') or existing.name.strip() in ('—', '-', '')) and c_name:
                    existing.name = c_name
                if (existing.designation in ('—', '-', 'Employee', '')) and extracted.get('designation'):
                    existing.designation = extracted['designation']
                if (existing.department in ('—', '-', 'General', '')) and extracted.get('department'):
                    existing.department = extracted['department']
                if not existing.phone and extracted.get('phone'):
                    existing.phone = extracted['phone']
                if not existing.email and extracted.get('email'):
                    existing.email = extracted['email']
                if (not existing.basic_salary or existing.basic_salary == 0) and extracted.get('salary'):
                    existing.basic_salary = extracted['salary']
            return existing, False

        # 3. Not existing: Create new Employee record in database
        final_name = extracted.get('name') or f"Employee {clean_code}"
        final_desig = extracted.get('designation') or "Employee"
        final_dept = extracted.get('department') or "General"

        new_emp = Employee(
            employee_id=clean_code,
            name=final_name,
            designation=final_desig,
            department=final_dept,
            status='Active',
            phone=extracted.get('phone'),
            email=extracted.get('email'),
            basic_salary=extracted.get('salary') or 0.0
        )
        db_session.add(new_emp)
        db_session.flush()
        logger.info(f"[DocOptimizer] Auto-created new Employee ID {clean_code} - Name: {final_name}")
        return new_emp, True

    def build_known_employee_map(self, employees):
        """
        Builds lookup dictionaries to match folder names and file names to Employee records.
        """
        code_to_emp = {}
        norm_code_to_emp = {}
        name_to_emp = {}
        for emp in employees:
            raw_code = str(emp.employee_id).strip()
            code_upper = raw_code.upper()
            code_to_emp[code_upper] = emp

            # Normalized numeric variations (leading zero stripped & zero-padded)
            if raw_code.isdigit():
                num = int(raw_code)
                clean_num = str(num)
                norm_code_to_emp[clean_num] = emp
                for pad in (2, 3, 4, 5):
                    padded = f"{num:0{pad}d}"
                    norm_code_to_emp[padded] = emp
                    norm_code_to_emp[f"EMP{padded}"] = emp
                    norm_code_to_emp[f"EMP-{padded}"] = emp
                    norm_code_to_emp[f"EMP_{padded}"] = emp
                    norm_code_to_emp[f"ID{padded}"] = emp
                    norm_code_to_emp[f"ID-{padded}"] = emp
                    norm_code_to_emp[f"ID_{padded}"] = emp
                norm_code_to_emp[f"EMP{code_upper}"] = emp
                norm_code_to_emp[f"EMP-{code_upper}"] = emp
                norm_code_to_emp[f"EMP_{code_upper}"] = emp
                norm_code_to_emp[f"ID{code_upper}"] = emp
                norm_code_to_emp[f"ID-{code_upper}"] = emp
                norm_code_to_emp[f"ID_{code_upper}"] = emp
            else:
                norm_code_to_emp[f"EMP{code_upper}"] = emp
                norm_code_to_emp[f"EMP-{code_upper}"] = emp
                norm_code_to_emp[f"EMP_{code_upper}"] = emp
                norm_code_to_emp[f"ID{code_upper}"] = emp
                norm_code_to_emp[f"ID-{code_upper}"] = emp
                norm_code_to_emp[f"ID_{code_upper}"] = emp

            clean_name = re.sub(r'[^a-zA-Z0-9\s]', '', (emp.name or '').lower()).strip()
            if clean_name:
                name_to_emp[clean_name] = emp

        return code_to_emp, name_to_emp, norm_code_to_emp

    def detect_emp_id_from_path(self, rel_path, code_to_emp=None, name_to_emp=None, norm_code_to_emp=None):
        """
        Scans folder names AND file names in the relative path to identify Employee ID.
        Checks both subfolder hierarchy (deepest first) and filename for ID patterns.
        Returns: (detected_code, matched_employee_obj_or_None)
        """
        norm_path = os.path.normpath(rel_path).replace('\\', '/')
        parts = [p.strip() for p in norm_path.split('/') if p.strip()]
        if not parts:
            return "UNKNOWN", None

        filename = parts[-1]
        base_name, _ = os.path.splitext(filename)
        folder_parts = parts[:-1]

        candidates = []
        for f in reversed(folder_parts):
            candidates.append(('folder', f))
        if base_name:
            candidates.append(('file', base_name))

        def get_emp_code(emp, default):
            if emp and hasattr(emp, 'employee_id'):
                return str(emp.employee_id).upper()
            return str(default).upper()

        delim_start = r'(?:^|[\s_\-\.\(\)\[\]\/\#\:\,])'
        delim_end = r'(?:[\s_\-\.\(\)\[\]\/\#\:\,]|$)'

        # PASS 1: Check candidates for known employee codes / names
        for kind, text in candidates:
            text_strip = text.strip()
            text_upper = text_strip.upper()

            if kind == 'folder' and text_upper in self.GENERIC_FOLDER_NAMES:
                continue

            # Exact match with code_to_emp or norm_code_to_emp
            if code_to_emp and text_upper in code_to_emp:
                emp = code_to_emp[text_upper]
                return get_emp_code(emp, text_upper), emp

            if norm_code_to_emp and text_upper in norm_code_to_emp:
                emp = norm_code_to_emp[text_upper]
                return get_emp_code(emp, text_upper), emp

            # Known employee code inside text with delimiter boundaries
            if code_to_emp:
                for k, emp in sorted(code_to_emp.items(), key=lambda x: len(x[0]), reverse=True):
                    if len(k) == 1:
                        if re.search(rf'(?:^|[_\-\(\[\#\:])\s*{re.escape(k)}\s*(?:[_\-\)\]\#\:]|$)', text_upper):
                            return get_emp_code(emp, k), emp
                    else:
                        if re.search(f'{delim_start}{re.escape(k)}{delim_end}', text_upper):
                            return get_emp_code(emp, k), emp

            # Check numeric matches with / without leading zeros
            if code_to_emp or norm_code_to_emp:
                m_nums = re.findall(r'(?:EMP[-_ ]?|ID[-_ ]?|EMP#|ID#)?(\d{1,6})', text, re.I)
                for num in m_nums:
                    clean_num = num.lstrip('0') or '0'
                    if code_to_emp and num in code_to_emp:
                        emp = code_to_emp[num]
                        return get_emp_code(emp, num), emp
                    if norm_code_to_emp and num in norm_code_to_emp:
                        emp = norm_code_to_emp[num]
                        return get_emp_code(emp, num), emp
                    if norm_code_to_emp and clean_num in norm_code_to_emp:
                        emp = norm_code_to_emp[clean_num]
                        return get_emp_code(emp, clean_num), emp

            # Check known employee name inside text
            if name_to_emp:
                clean_t = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower()).strip()
                for name_key, emp in sorted(name_to_emp.items(), key=lambda x: len(x[0]), reverse=True):
                    if name_key and len(name_key) > 3 and name_key in clean_t:
                        return get_emp_code(emp, emp.employee_id), emp

        # PASS 2: Not in database, but has clear ID pattern
        for kind, text in candidates:
            text_upper = text.strip().upper()
            if kind == 'folder' and text_upper in self.GENERIC_FOLDER_NAMES:
                continue

            m_prefix = re.search(r'\b((?:EMP|ID)[-_ ]?\d+)\b', text, re.I)
            if m_prefix:
                code = m_prefix.group(1).upper().replace('_', '-').replace(' ', '-')
                return code, None

            m_num = re.search(r'(?:^|[_\-\s\(\[\#])(\d{1,6})(?:[_\-\s\.\)\]\#]|$)', text)
            if m_num:
                code = m_num.group(1)
                return code, None

        # PASS 3: Fallback - use non-generic folder or filename base
        for kind, text in candidates:
            text_upper = text.strip().upper()
            if kind == 'folder' and text_upper in self.GENERIC_FOLDER_NAMES:
                continue
            clean_code = re.sub(r'[^a-zA-Z0-9\-_]', '_', text)[:30].strip('_')
            if clean_code:
                return clean_code, None

        return "UNKNOWN", None

    def group_files_for_preview(self, file_entries, employees):
        """
        Groups client-submitted file list metadata into a structured preview by Employee ID.
        file_entries: list of dicts with keys: rel_path, filename, size, (optional file_token/index)
        """
        maps = self.build_known_employee_map(employees)
        if len(maps) == 3:
            code_to_emp, name_to_emp, norm_code_to_emp = maps
        else:
            code_to_emp, name_to_emp = maps
            norm_code_to_emp = {}

        grouped = {}
        unsupported = []
        total_valid = 0

        for idx, entry in enumerate(file_entries):
            orig_name = entry.get('filename') or os.path.basename(entry.get('rel_path', ''))
            rel_path = entry.get('rel_path') or orig_name
            size = int(entry.get('size') or 0)

            ext = os.path.splitext(orig_name)[1].lower()
            if ext not in ALLOWED_EXTENSIONS:
                unsupported.append({
                    'filename': orig_name,
                    'rel_path': rel_path,
                    'ext': ext or '(no extension)',
                    'reason': 'Unsupported format'
                })
                continue

            emp_code, emp_obj = self.detect_emp_id_from_path(rel_path, code_to_emp, name_to_emp, norm_code_to_emp)
            emp_code = emp_code or 'UNKNOWN'

            category = self.get_file_category(ext)
            icon = self.get_file_icon(category)

            if emp_code not in grouped:
                grouped[emp_code] = {
                    'emp_code': emp_code,
                    'matched': bool(emp_obj),
                    'db_employee_id': emp_obj.id if emp_obj else None,
                    'emp_name': emp_obj.name if emp_obj else f"Employee {emp_code}",
                    'designation': emp_obj.designation if emp_obj else '—',
                    'department': emp_obj.department if emp_obj else '—',
                    'photo': emp_obj.photo if emp_obj else None,
                    'files': [],
                    'total_size_bytes': 0
                }

            grouped[emp_code]['files'].append({
                'file_index': idx,
                'filename': orig_name,
                'rel_path': rel_path,
                'file_size': size,
                'file_size_kb': round(size / 1024, 1),
                'file_ext': ext,
                'category': category,
                'icon': icon
            })
            grouped[emp_code]['total_size_bytes'] += size
            total_valid += 1

        # Format list sorted with matched employees first
        emp_list = list(grouped.values())
        emp_list.sort(key=lambda x: (not x['matched'], x['emp_code']))

        for item in emp_list:
            item['total_size_kb'] = round(item['total_size_bytes'] / 1024, 1)
            item['file_count'] = len(item['files'])
            if not item['matched']:
                cand_name = self.extract_candidate_name_from_filenames([f['filename'] for f in item['files']])
                if cand_name:
                    item['emp_name'] = cand_name
                    item['candidate_name'] = cand_name
                    item['is_auto_create'] = True
                    item['auto_create_hint'] = f"New Employee '{cand_name}' will be auto-created in database"
                else:
                    item['is_auto_create'] = True
                    item['auto_create_hint'] = f"New Employee (ID: {item['emp_code']}) will be auto-created in database"
            else:
                item['is_auto_create'] = False

        return {
            'success': True,
            'total_files_submitted': len(file_entries),
            'valid_files_count': total_valid,
            'unsupported_count': len(unsupported),
            'unsupported_files': unsupported[:20],
            'employee_count': len(emp_list),
            'employees': emp_list
        }

    def scan_local_folder_on_disk(self, root_path, employees):
        """
        Scans a local Windows folder on disk and returns the preview structure.
        """
        if not root_path or not os.path.exists(root_path):
            return {
                'success': False,
                'message': f"Path not found: {root_path}"
            }

        file_entries = []
        for root, dirs, files in os.walk(root_path):
            for f in files:
                abs_path = os.path.join(root, f)
                rel_path = os.path.relpath(abs_path, root_path)
                try:
                    size = os.path.getsize(abs_path)
                except Exception:
                    size = 0
                file_entries.append({
                    'filename': f,
                    'rel_path': rel_path,
                    'abs_path': abs_path,
                    'size': size
                })

        res = self.group_files_for_preview(file_entries, employees)
        res['source_root'] = root_path
        return res

    def optimize_and_store_file(
        self,
        source,
        target_abs,
        ext,
        emp_code='UNKNOWN',
        orig_name='document',
        preset='balanced',
        custom_max_width=None,
        custom_quality=None,
        custom_pdf_target=None
    ):
        """
        Processes and optimizes the document before saving according to preset:
        - 'balanced' (Standard): Images max 1080p, Quality 75%, PDF target < 1MB, DPI 150
        - 'high' (High Quality): Images max 1920p, Quality 85%, PDF target < 2MB, DPI 200
        - 'small' (Small Size): Images max 720p, Quality 60%, PDF target < 500KB, DPI 100
        - 'original' (No Compression): Direct byte copy without alterations
        - 'custom': User-specified max width, quality, and PDF target size
        """
        # Read raw stream or disk file content
        if hasattr(source, 'read'):
            source.seek(0, os.SEEK_END)
            orig_size = source.tell()
            source.seek(0)
            raw_content = source.read()
        else:
            orig_size = os.path.getsize(source)
            with open(source, 'rb') as f:
                raw_content = f.read()

        ext = ext.lower().strip()
        final_size = orig_size
        optimized_success = False

        # --- 0. Original preset: bypass all compression ---
        if preset == 'original':
            with open(target_abs, 'wb') as dst:
                dst.write(raw_content)
            final_size = orig_size
            status_str = f"{format_bytes_human(final_size)} (Original - No compression)"
            logger.info(f"[DocOptimizer] Preset:original | ID:{emp_code} | {orig_name} | {status_str}")
            print(f"[DocOptimizer] Preset:original | ID:{emp_code} | {orig_name} | {status_str}")
            return orig_size, final_size, status_str

        # Resolve parameters based on preset
        if preset == 'custom':
            try:
                max_width = int(custom_max_width) if custom_max_width else 1280
            except Exception:
                max_width = 1280
            try:
                quality = max(20, min(100, int(custom_quality))) if custom_quality else 75
            except Exception:
                quality = 75
            try:
                target_pdf_kb = int(custom_pdf_target) if custom_pdf_target else 1024
            except Exception:
                target_pdf_kb = 1024
            pdf_dpi = 140
            preset_label = f"Custom ({max_width}px, {quality}%)"
        else:
            cfg = PRESET_CONFIGS.get(preset, PRESET_CONFIGS['balanced'])
            max_width = cfg['max_width']
            quality = cfg['quality']
            pdf_dpi = cfg['pdf_dpi']
            target_pdf_kb = cfg['target_pdf_kb']
            preset_label = cfg['name']

        # --- 1. Images (.jpg, .jpeg, .png) ---
        if ext in ('.jpg', '.jpeg', '.png'):
            try:
                img = Image.open(io.BytesIO(raw_content))

                # EXIF rotation correction
                try:
                    img = ImageOps.exif_transpose(img)
                except Exception:
                    pass

                # Resizing: if max_width is set and width > max_width, downscale keeping aspect ratio
                w, h = img.size
                if max_width and w > max_width:
                    ratio = max_width / float(w)
                    new_w = max_width
                    new_h = max(1, int(round(h * ratio)))
                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

                out_buf = io.BytesIO()
                if ext in ('.jpg', '.jpeg'):
                    if img.mode in ('RGBA', 'LA', 'P'):
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        if img.mode == 'RGBA':
                            bg.paste(img, mask=img.split()[-1])
                        else:
                            bg.paste(img)
                        img = bg
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                    img.save(out_buf, format='JPEG', quality=quality, optimize=True)
                elif ext == '.png':
                    img.save(out_buf, format='PNG', optimize=True, compress_level=9)
                else:
                    img.save(out_buf, format=img.format or 'JPEG', quality=quality, optimize=True)

                comp_bytes = out_buf.getvalue()
                if len(comp_bytes) > 0 and (orig_size == 0 or len(comp_bytes) < orig_size):
                    with open(target_abs, 'wb') as dst:
                        dst.write(comp_bytes)
                    final_size = len(comp_bytes)
                    optimized_success = True
                else:
                    with open(target_abs, 'wb') as dst:
                        dst.write(raw_content)
                    final_size = orig_size
            except Exception as img_err:
                logger.warning(f"Image optimization warning for {orig_name}: {img_err}")
                with open(target_abs, 'wb') as dst:
                    dst.write(raw_content)
                final_size = orig_size

        # --- 2. PDFs (.pdf) ---
        elif ext == '.pdf':
            try:
                opt_bytes = None
                # First try pymupdf
                try:
                    import pymupdf
                    doc = pymupdf.open(stream=raw_content, filetype='pdf')
                    # Rewrite images inside PDF if preset specifies DPI / quality target
                    if pdf_dpi and quality:
                        try:
                            doc.rewrite_images(dpi_threshold=pdf_dpi, dpi_target=pdf_dpi, quality=quality)
                        except Exception as rw_err:
                            logger.debug(f"rewrite_images skipped: {rw_err}")

                    opt_bytes = doc.tobytes(garbage=4, deflate=True, clean=True)
                    doc.close()
                except Exception as mu_err:
                    logger.warning(f"pymupdf error on {orig_name}: {mu_err}")

                # Fallback to pypdf
                if not opt_bytes:
                    try:
                        import pypdf
                        reader = pypdf.PdfReader(io.BytesIO(raw_content))
                        writer = pypdf.PdfWriter()
                        for p in reader.pages:
                            p.compress_content_streams()
                            writer.add_page(p)
                        writer.compress_identical_objects()
                        pdf_buf = io.BytesIO()
                        writer.write(pdf_buf)
                        opt_bytes = pdf_buf.getvalue()
                    except Exception as pypdf_err:
                        logger.warning(f"pypdf fallback error on {orig_name}: {pypdf_err}")

                if opt_bytes and len(opt_bytes) > 0 and (orig_size == 0 or len(opt_bytes) < orig_size):
                    with open(target_abs, 'wb') as dst:
                        dst.write(opt_bytes)
                    final_size = len(opt_bytes)
                    optimized_success = True
                else:
                    with open(target_abs, 'wb') as dst:
                        dst.write(raw_content)
                    final_size = orig_size
            except Exception as pdf_err:
                logger.warning(f"PDF optimization warning on {orig_name}: {pdf_err}")
                with open(target_abs, 'wb') as dst:
                    dst.write(raw_content)
                final_size = orig_size

        # --- 3. Word & Excel (.docx, .doc, .xlsx, .xls) ---
        else:
            with open(target_abs, 'wb') as dst:
                dst.write(raw_content)
            final_size = orig_size

        # Construct status feedback & logging
        orig_formatted = format_bytes_human(orig_size)
        final_formatted = format_bytes_human(final_size)

        if optimized_success and orig_size > final_size and orig_size > 0:
            saved = orig_size - final_size
            pct = round((saved / float(orig_size)) * 100, 1)
            status_str = f"[{preset_label}] Orig: {orig_formatted} -> Comp: {final_formatted} (-{pct}%)"
        elif ext in ('.docx', '.doc', '.xlsx', '.xls'):
            status_str = f"{final_formatted} (Original file copied unaltered)"
        else:
            status_str = f"{final_formatted} (Optimal / Preserved)"

        logger.info(f"[DocOptimizer] Preset:{preset} | ID:{emp_code} | {orig_name} | {status_str}")
        print(f"[DocOptimizer] Preset:{preset} | ID:{emp_code} | {orig_name} | {status_str}")

        return orig_size, final_size, status_str

    def save_file_for_employee(
        self,
        file_obj_or_stream,
        emp_code,
        orig_filename=None,
        preset='balanced',
        custom_max_width=None,
        custom_quality=None,
        custom_pdf_target=None
    ):
        """
        Saves uploaded file into app_storage/employee_docs/{emp_code}/{safe_filename}
        with preset-based image/PDF compression applied in-flight.
        Returns metadata dict for EmployeeDocument record.
        """
        emp_code_clean = re.sub(r'[^a-zA-Z0-9\-_]', '_', str(emp_code)).strip() or 'UNKNOWN'
        target_dir = os.path.join(self.storage_dir, emp_code_clean)
        os.makedirs(target_dir, exist_ok=True)

        orig_name = orig_filename or getattr(file_obj_or_stream, 'filename', 'document')
        base, ext = os.path.splitext(orig_name)
        ext = ext.lower()
        if ext not in ALLOWED_EXTENSIONS:
            return None

        # Clean base name
        safe_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', base)[:40]
        ts = datetime.utcnow().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
        unique_name = f"{safe_base}_{ts}{ext}"
        target_abs = os.path.join(target_dir, unique_name)

        orig_size, final_size, comp_status = self.optimize_and_store_file(
            source=file_obj_or_stream,
            target_abs=target_abs,
            ext=ext,
            emp_code=emp_code_clean,
            orig_name=orig_name,
            preset=preset,
            custom_max_width=custom_max_width,
            custom_quality=custom_quality,
            custom_pdf_target=custom_pdf_target
        )

        rel_stored_path = f"app_storage/employee_docs/{emp_code_clean}/{unique_name}"
        category = self.get_file_category(ext)

        return {
            'filename': unique_name,
            'original_filename': orig_name,
            'file_size': final_size,
            'original_size': orig_size,
            'compression_status': comp_status,
            'file_ext': ext,
            'file_type': category,
            'file_path': rel_stored_path,
            'abs_path': target_abs,
            'emp_code': emp_code_clean
        }

    def copy_file_for_employee(
        self,
        source_abs_path,
        emp_code,
        orig_filename=None,
        preset='balanced',
        custom_max_width=None,
        custom_quality=None,
        custom_pdf_target=None
    ):
        """
        Copies a local file on disk into app_storage/employee_docs/{emp_code}/
        with preset-based image/PDF compression applied in-flight.
        """
        if not os.path.exists(source_abs_path):
            return None

        emp_code_clean = re.sub(r'[^a-zA-Z0-9\-_]', '_', str(emp_code)).strip() or 'UNKNOWN'
        target_dir = os.path.join(self.storage_dir, emp_code_clean)
        os.makedirs(target_dir, exist_ok=True)

        orig_name = orig_filename or os.path.basename(source_abs_path)
        base, ext = os.path.splitext(orig_name)
        ext = ext.lower()
        if ext not in ALLOWED_EXTENSIONS:
            return None

        safe_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', base)[:40]
        ts = datetime.utcnow().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
        unique_name = f"{safe_base}_{ts}{ext}"
        target_abs = os.path.join(target_dir, unique_name)

        orig_size, final_size, comp_status = self.optimize_and_store_file(
            source=source_abs_path,
            target_abs=target_abs,
            ext=ext,
            emp_code=emp_code_clean,
            orig_name=orig_name,
            preset=preset,
            custom_max_width=custom_max_width,
            custom_quality=custom_quality,
            custom_pdf_target=custom_pdf_target
        )

        rel_stored_path = f"app_storage/employee_docs/{emp_code_clean}/{unique_name}"
        category = self.get_file_category(ext)

        return {
            'filename': unique_name,
            'original_filename': orig_name,
            'file_size': final_size,
            'original_size': orig_size,
            'compression_status': comp_status,
            'file_ext': ext,
            'file_type': category,
            'file_path': rel_stored_path,
            'abs_path': target_abs,
            'emp_code': emp_code_clean
        }

    def resolve_physical_path(self, doc):
        """
        Resolves the absolute path for an EmployeeDocument.
        Checks new app_storage path, emp_code subfolder, or legacy static folder.
        """
        # 1. From stored file_path if valid
        if doc.file_path:
            candidate = os.path.join(BASE_DIR, doc.file_path)
            if os.path.exists(candidate):
                return candidate

        # 2. In app_storage/employee_docs/{emp_code}/{filename}
        if doc.emp_code:
            candidate = os.path.join(self.storage_dir, str(doc.emp_code), doc.filename)
            if os.path.exists(candidate):
                return candidate

        # 3. Direct in storage_dir
        candidate = os.path.join(self.storage_dir, doc.filename)
        if os.path.exists(candidate):
            return candidate

        # 4. Legacy static uploads folder
        legacy = os.path.join(BASE_DIR, 'static', 'uploads', 'employee_docs', doc.filename)
        if os.path.exists(legacy):
            return legacy

        return None

    def delete_single_document(self, doc, db_session):
        """
        Deletes a single document record and its physical file.
        """
        phys_path = self.resolve_physical_path(doc)
        if phys_path and os.path.exists(phys_path):
            try:
                os.remove(phys_path)
                # Remove parent directory if empty
                parent = os.path.dirname(phys_path)
                if parent != self.storage_dir and os.path.exists(parent) and not os.listdir(parent):
                    os.rmdir(parent)
            except Exception:
                pass

        db_session.delete(doc)
        db_session.commit()
        return True

    def delete_all_attached_documents(self, db_session, EmployeeDocumentClass):
        """
        Permanently deletes ALL attached documents:
        1. Deletes all EmployeeDocument records from database
        2. Deletes all physical files & directories inside app_storage/employee_docs/
        """
        all_docs = EmployeeDocumentClass.query.all()
        doc_count = len(all_docs)

        # 1. Delete DB records
        EmployeeDocumentClass.query.delete()
        db_session.commit()

        # 2. Delete physical storage folder
        files_deleted = 0
        if os.path.exists(self.storage_dir):
            import gc
            gc.collect()
            try:
                for item in os.listdir(self.storage_dir):
                    item_path = os.path.join(self.storage_dir, item)
                    if os.path.isdir(item_path):
                        for root, dirs, files in os.walk(item_path):
                            files_deleted += len(files)
                        shutil.rmtree(item_path, ignore_errors=True)
                    elif os.path.isfile(item_path):
                        os.remove(item_path)
                        files_deleted += 1
            except Exception as e:
                print(f"Error purging document storage: {e}")

        # Ensure directory is ready for next uploads
        os.makedirs(self.storage_dir, exist_ok=True)

        return {
            'success': True,
            'records_deleted': doc_count,
            'files_deleted': files_deleted,
            'message': f"All {doc_count} attached document records and physical files were permanently deleted."
        }

    def get_storage_stats(self, db_session, EmployeeDocumentClass):
        """
        Returns high-level statistics for document dashboard including compression savings.
        """
        total_docs = EmployeeDocumentClass.query.count()
        total_employees_with_docs = db_session.query(EmployeeDocumentClass.emp_code).distinct().count()

        total_bytes = 0
        total_orig_bytes = 0
        type_breakdown = {'PDF': 0, 'IMAGE': 0, 'DOCX': 0, 'EXCEL': 0, 'OTHER': 0}

        docs = EmployeeDocumentClass.query.all()
        for d in docs:
            cur_sz = d.file_size or 0
            orig_sz = getattr(d, 'original_size', None) or cur_sz
            total_bytes += cur_sz
            total_orig_bytes += orig_sz
            cat = d.file_type or self.get_file_category(d.file_ext or '')
            type_breakdown[cat] = type_breakdown.get(cat, 0) + 1

        total_mb = round(total_bytes / (1024 * 1024), 2)
        total_kb = round(total_bytes / 1024, 1)
        total_saved_bytes = max(0, total_orig_bytes - total_bytes)
        total_saved_mb = round(total_saved_bytes / (1024 * 1024), 2)
        saved_pct = round((total_saved_bytes / total_orig_bytes * 100), 1) if total_orig_bytes > 0 else 0

        return {
            'total_docs': total_docs,
            'total_employees_with_docs': total_employees_with_docs,
            'total_size_mb': total_mb,
            'total_size_kb': total_kb,
            'total_saved_mb': total_saved_mb,
            'saved_pct': saved_pct,
            'type_breakdown': type_breakdown
        }

    delete_all_documents = delete_all_attached_documents


# Singleton instance
employee_doc_manager = EmployeeDocManager()
