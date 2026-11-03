"""
Govt Holiday Manager & Attendance Intelligence Engine
Parses Government Holiday calendars from PDF, Excel (.xlsx, .xls), and CSV files.
Provides built-in Bangladesh Government Holiday Gazette calendars and integrates
with employee biometric attendance sheets, duty registers, and monthly payroll matrices.
"""

import os
import re
import io
import logging
from datetime import datetime, date, timedelta
import openpyxl

logger = logging.getLogger(__name__)

# Month name mapping (English and Bengali)
MONTH_MAP = {
    'january': 1, 'jan': 1, 'জানুয়ারি': 1, 'জানুয়ারি': 1,
    'february': 2, 'feb': 2, 'ফেব্রুয়ারি': 2, 'ফেব্রুয়ারি': 2,
    'march': 3, 'mar': 3, 'মার্চ': 3,
    'april': 4, 'apr': 4, 'এপ্রিল': 4,
    'may': 5, 'মে': 5,
    'june': 6, 'jun': 6, 'জুন': 6,
    'july': 7, 'jul': 7, 'জুলাই': 7,
    'august': 8, 'aug': 8, 'আগস্ট': 8, 'আগষ্ট': 8,
    'september': 9, 'sep': 9, 'sept': 9, 'সেপ্টেম্বর': 9,
    'october': 10, 'oct': 10, 'অক্টোবর': 10,
    'november': 11, 'nov': 11, 'নভেম্বর': 11,
    'december': 12, 'dec': 12, 'ডিসেম্বর': 12
}

BENGALI_DIGITS = {
    '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4',
    '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9'
}

def to_english_digits(text: str) -> str:
    """Converts Bengali numerals to standard ASCII digits."""
    if not text:
        return ""
    res = []
    for ch in str(text):
        res.append(BENGALI_DIGITS.get(ch, ch))
    return "".join(res)


# Built-in Official Bangladesh Government Holidays for 2026
OFFICIAL_BD_GOVT_HOLIDAYS_2026 = [
    {
        'name': 'Shab-e-Barat (শবে বরাত)',
        'start_date': date(2026, 2, 4),
        'end_date': date(2026, 2, 4),
        'type': 'Executive Order',
        'days': 1
    },
    {
        'name': 'Shaheed Day & Int. Mother Language Day (শহীদ দিবস ও আন্তর্জাতিক মাতৃভাষা দিবস)',
        'start_date': date(2026, 2, 21),
        'end_date': date(2026, 2, 21),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': "National Children's Day / Bangabandhu Birthday (জাতীয় শিশু দিবস)",
        'start_date': date(2026, 3, 17),
        'end_date': date(2026, 3, 17),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Shab-e-Qadr (পবিত্র শবে কদর)',
        'start_date': date(2026, 3, 17),
        'end_date': date(2026, 3, 17),
        'type': 'Executive Order',
        'days': 1
    },
    {
        'name': 'Jumatul Wida (জুমাতুল বিদা)',
        'start_date': date(2026, 3, 20),
        'end_date': date(2026, 3, 20),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Eid-ul-Fitr (পবিত্র ঈদুল ফিতর)',
        'start_date': date(2026, 3, 20),
        'end_date': date(2026, 3, 23),
        'type': 'General Holiday',
        'days': 4
    },
    {
        'name': 'Independence & National Day (স্বাধীনতা ও জাতীয় দিবস)',
        'start_date': date(2026, 3, 26),
        'end_date': date(2026, 3, 26),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Pahela Baishakh / Bangla New Year (পহেলা বৈশাখ / বাংলা নববর্ষ)',
        'start_date': date(2026, 4, 14),
        'end_date': date(2026, 4, 14),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'May Day (মে দিবস / আন্তর্জাতিক শ্রমিক দিবস)',
        'start_date': date(2026, 5, 1),
        'end_date': date(2026, 5, 1),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Buddha Purnima (বুদ্ধ পূর্ণিমা)',
        'start_date': date(2026, 5, 12),
        'end_date': date(2026, 5, 12),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Eid-ul-Azha (পবিত্র ঈদুল আজহা)',
        'start_date': date(2026, 5, 26),
        'end_date': date(2026, 5, 29),
        'type': 'General Holiday',
        'days': 4
    },
    {
        'name': 'Holy Ashura / Muharram (পবিত্র আশুরা)',
        'start_date': date(2026, 6, 26),
        'end_date': date(2026, 6, 26),
        'type': 'Executive Order',
        'days': 1
    },
    {
        'name': 'Eid-e-Miladunnabi (ঈদ-ই-মিলাদুন্নবী)',
        'start_date': date(2026, 8, 25),
        'end_date': date(2026, 8, 25),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Janmashtami (শুভ জন্মাষ্টমী)',
        'start_date': date(2026, 9, 4),
        'end_date': date(2026, 9, 4),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Durga Puja / Bijoya Dashami (শারদীয় দুর্গাপূজা / বিজয়া দশমী)',
        'start_date': date(2026, 10, 20),
        'end_date': date(2026, 10, 21),
        'type': 'General Holiday',
        'days': 2
    },
    {
        'name': 'Victory Day (মহান বিজয় দিবস)',
        'start_date': date(2026, 12, 16),
        'end_date': date(2026, 12, 16),
        'type': 'General Holiday',
        'days': 1
    },
    {
        'name': 'Christmas Day (যিশু খ্রিস্টের জন্মদিন / বড়দিন)',
        'start_date': date(2026, 12, 25),
        'end_date': date(2026, 12, 25),
        'type': 'General Holiday',
        'days': 1
    }
]


class GovtHolidayManager:
    """
    Parser and Analyzer for Government Holidays from PDF and Excel files.
    """

    def parse_excel_holiday_file(self, file_content_or_path, default_year=2026):
        """
        Parses Excel (.xlsx, .xls) holiday schedule.
        Returns list of parsed holiday dictionaries:
        [{ 'name': '...', 'start_date': date, 'end_date': date, 'type': '...', 'days': 1 }]
        """
        if hasattr(file_content_or_path, 'read'):
            bio = io.BytesIO(file_content_or_path.read())
        elif isinstance(file_content_or_path, bytes):
            bio = io.BytesIO(file_content_or_path)
        else:
            bio = file_content_or_path

        wb = openpyxl.load_workbook(bio, data_only=True)
        ws = wb.active

        holidays = []
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return holidays

        # Detect header row
        header_row_idx = 0
        date_col_idx = None
        name_col_idx = None
        type_col_idx = None
        end_date_col_idx = None

        for idx, row in enumerate(rows[:10]):
            row_strs = [str(c or '').strip().lower() for c in row]
            for c_i, val in enumerate(row_strs):
                if any(w in val for w in ['date', 'তারিখ', 'start date', 'from date']) or val in ['day', 'দিন']:
                    if date_col_idx is None:
                        date_col_idx = c_i
                        header_row_idx = idx
                elif any(w in val for w in ['holiday', 'occasion', 'name', 'উৎসব', 'নাম', 'বিবরণ', 'description', 'title']):
                    name_col_idx = c_i
                elif any(w in val for w in ['type', 'category', 'ধরন', 'শ্রেণীর']):
                    type_col_idx = c_i
                elif any(w in val for w in ['end date', 'to date', 'পর্যন্ত']):
                    end_date_col_idx = c_i

            if date_col_idx is not None and name_col_idx is not None:
                break

        # Fallback if no explicit headers found
        if date_col_idx is None or name_col_idx is None:
            # Look at first row with date object or date string
            for idx, row in enumerate(rows):
                for c_i, cell in enumerate(row):
                    if isinstance(cell, (datetime, date)):
                        date_col_idx = c_i
                        name_col_idx = 1 if c_i == 0 else (c_i + 1 if c_i + 1 < len(row) else 0)
                        header_row_idx = max(0, idx - 1)
                        break
                if date_col_idx is not None:
                    break

        if date_col_idx is None:
            date_col_idx = 0
        if name_col_idx is None:
            name_col_idx = 1

        for row in rows[header_row_idx + 1:]:
            if not row or len(row) <= max(date_col_idx, name_col_idx):
                continue
            date_val = row[date_col_idx]
            name_val = row[name_col_idx]
            if not date_val or not name_val:
                continue

            parsed_start, parsed_end = self._parse_date_cell(date_val, default_year)
            if not parsed_start:
                continue

            if end_date_col_idx is not None and end_date_col_idx < len(row) and row[end_date_col_idx]:
                e_start, _ = self._parse_date_cell(row[end_date_col_idx], default_year)
                if e_start:
                    parsed_end = e_start

            h_type = "General Holiday"
            if type_col_idx is not None and type_col_idx < len(row) and row[type_col_idx]:
                h_type = str(row[type_col_idx]).strip()

            clean_name = str(name_val).strip()
            days_count = (parsed_end - parsed_start).days + 1 if parsed_end else 1

            holidays.append({
                'name': clean_name,
                'start_date': parsed_start,
                'end_date': parsed_end or parsed_start,
                'type': h_type,
                'days': max(1, days_count)
            })

        return holidays

    def parse_pdf_holiday_file(self, file_content_or_path, default_year=2026):
        """
        Parses PDF official gazette or holiday schedule using PyMuPDF (fitz) or pypdf.
        Extracts both Bengali and English holiday notices.
        """
        raw_text = ""

        # Option A: PyMuPDF (fitz)
        try:
            import fitz
            if hasattr(file_content_or_path, 'read'):
                bytes_data = file_content_or_path.read()
            elif isinstance(file_content_or_path, bytes):
                bytes_data = file_content_or_path
            else:
                with open(file_content_or_path, 'rb') as f:
                    bytes_data = f.read()

            doc = fitz.open(stream=bytes_data, filetype="pdf")
            for page in doc:
                raw_text += page.get_text() + "\n"
        except Exception as e:
            logger.warning(f"fitz PDF parse fallback: {e}")
            try:
                import pypdf
                if hasattr(file_content_or_path, 'seek'):
                    file_content_or_path.seek(0)
                    reader = pypdf.PdfReader(file_content_or_path)
                elif isinstance(file_content_or_path, bytes):
                    reader = pypdf.PdfReader(io.BytesIO(file_content_or_path))
                else:
                    reader = pypdf.PdfReader(file_content_or_path)
                for page in reader.pages:
                    raw_text += page.extract_text() or "" + "\n"
            except Exception as e2:
                logger.error(f"pypdf extraction failed: {e2}")

        return self.parse_holiday_text_content(raw_text, default_year=default_year)

    def parse_holiday_text_content(self, text, default_year=2026):
        """
        Regex-powered parser that handles dates, multi-day ranges, and holiday titles
        from plain text or PDF text streams in both English and Bengali.
        """
        if not text:
            return []

        holidays = []
        lines = [l.strip() for l in text.splitlines() if l.strip()]

        # Common date patterns:
        # 1. '21 February 2026' or '21 February' or '২১ ফেব্রুয়ারি ২০২৬'
        # 2. '20-23 March 2026' or '২০-২৩ মার্চ'
        # 3. '2026-03-26' or '26/03/2026' or '26-03-2026'
        
        for line in lines:
            # Skip non-holiday header lines
            if len(line) < 5 or any(w in line.lower() for w in ['গণপ্রজাতন্ত্রী', 'জনপ্রশাসন', 'স্মারক নম্বর', 'সরকারি ছুটি', 'গেজেট', 'ছুটির তালিকা', 'table of contents']):
                # Check if year is mentioned
                m_yr = re.search(r'\b(202[4-9]|203[0-5])\b', line)
                if m_yr:
                    try:
                        default_year = int(m_yr.group(1))
                    except Exception:
                        pass
                continue

            parsed = self._extract_holiday_from_line(line, default_year)
            if parsed:
                holidays.extend(parsed)

        # De-duplicate holidays with same start_date and name
        unique_holidays = []
        seen = set()
        for h in holidays:
            key = (h['name'].lower(), h['start_date'])
            if key not in seen:
                seen.add(key)
                unique_holidays.append(h)

        return unique_holidays

    def _extract_holiday_from_line(self, line, default_year):
        """Extracts candidate holiday dict from a single text line."""
        norm_line = to_english_digits(line)
        
        # 1. Pattern: "20-23 March" or "20 to 23 March"
        m_range = re.search(r'(\d{1,2})\s*(?:-|to|থেকে|হতে)\s*(\d{1,2})\s+([a-zA-Z\u0980-\u09FF]+)(?:\s+(\d{4}))?', norm_line, re.IGNORECASE)
        if m_range:
            d_start = int(m_range.group(1))
            d_end = int(m_range.group(2))
            m_name = m_range.group(3).lower()
            yr = int(m_range.group(4)) if m_range.group(4) else default_year
            month_num = MONTH_MAP.get(m_name)
            if month_num and 1 <= d_start <= 31 and 1 <= d_end <= 31:
                try:
                    s_date = date(yr, month_num, d_start)
                    e_date = date(yr, month_num, d_end)
                    name_candidate = norm_line.replace(m_range.group(0), '').strip(' ,.-–—:|\t')
                    # Clean title
                    name_candidate = re.sub(r'^\d+[\.\)]\s*', '', name_candidate).strip()
                    if len(name_candidate) >= 3:
                        return [{
                            'name': name_candidate,
                            'start_date': s_date,
                            'end_date': e_date,
                            'type': 'General Holiday',
                            'days': (e_date - s_date).days + 1
                        }]
                except Exception:
                    pass

        # 2. Pattern: "21 February 2026" or "21 February"
        m_single = re.search(r'(\d{1,2})\s+([a-zA-Z\u0980-\u09FF]+)(?:\s+(\d{4}))?', norm_line, re.IGNORECASE)
        if m_single:
            d_num = int(m_single.group(1))
            m_name = m_single.group(2).lower()
            yr = int(m_single.group(3)) if m_single.group(3) else default_year
            month_num = MONTH_MAP.get(m_name)
            if month_num and 1 <= d_num <= 31:
                try:
                    s_date = date(yr, month_num, d_num)
                    name_candidate = norm_line.replace(m_single.group(0), '').strip(' ,.-–—:|\t')
                    name_candidate = re.sub(r'^\d+[\.\)]\s*', '', name_candidate).strip()
                    if len(name_candidate) >= 3:
                        return [{
                            'name': name_candidate,
                            'start_date': s_date,
                            'end_date': s_date,
                            'type': 'General Holiday',
                            'days': 1
                        }]
                except Exception:
                    pass

        # 3. ISO / Slash pattern: "2026-03-26" or "26/03/2026" or "26-03-2026"
        m_iso = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', norm_line)
        if m_iso:
            try:
                s_date = date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
                name_candidate = norm_line.replace(m_iso.group(0), '').strip(' ,.-–—:|\t')
                name_candidate = re.sub(r'^\d+[\.\)]\s*', '', name_candidate).strip()
                if len(name_candidate) >= 3:
                    return [{
                        'name': name_candidate,
                        'start_date': s_date,
                        'end_date': s_date,
                        'type': 'General Holiday',
                        'days': 1
                    }]
            except Exception:
                pass

        m_dmy = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', norm_line)
        if m_dmy:
            try:
                s_date = date(int(m_dmy.group(3)), int(m_dmy.group(2)), int(m_dmy.group(1)))
                name_candidate = norm_line.replace(m_dmy.group(0), '').strip(' ,.-–—:|\t')
                name_candidate = re.sub(r'^\d+[\.\)]\s*', '', name_candidate).strip()
                if len(name_candidate) >= 3:
                    return [{
                        'name': name_candidate,
                        'start_date': s_date,
                        'end_date': s_date,
                        'type': 'General Holiday',
                        'days': 1
                    }]
            except Exception:
                pass

        return None

    def _parse_date_cell(self, cell_value, default_year):
        """Converts an Excel cell value into start_date and end_date."""
        if isinstance(cell_value, datetime):
            d = cell_value.date()
            return d, d
        if isinstance(cell_value, date):
            return cell_value, cell_value

        val_str = to_english_digits(str(cell_value).strip())
        
        # Try range '20-23 March' or '2026-03-20 to 2026-03-23'
        if any(sep in val_str for sep in [' to ', ' – ', ' — ', 'থেকে', 'হতে']):
            parts = re.split(r'\s*(?:to|–|—|থেকে|হতে)\s*', val_str, flags=re.IGNORECASE)
            if len(parts) >= 2:
                s1, _ = self._parse_date_cell(parts[0], default_year)
                s2, _ = self._parse_date_cell(parts[1], default_year)
                if s1 and s2:
                    return s1, s2
                elif s1:
                    return s1, s1

        # Check '20-23 March' or '20 - 23 March'
        m_drange = re.search(r'^(\d{1,2})\s*(?:-|–|—)\s*(\d{1,2})\s+([a-zA-Z\u0980-\u09FF]+)', val_str)
        if m_drange:
            d1 = int(m_drange.group(1))
            d2 = int(m_drange.group(2))
            m_num = MONTH_MAP.get(m_drange.group(3).lower())
            if m_num:
                try:
                    return date(default_year, m_num, d1), date(default_year, m_num, d2)
                except Exception:
                    pass

        # Standard date formats
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d', '%d-%b-%Y', '%d %B %Y', '%d %b %Y'):
            try:
                d = datetime.strptime(val_str, fmt).date()
                return d, d
            except Exception:
                pass

        # Try day + month string: '21 February'
        m = re.search(r'(\d{1,2})\s+([a-zA-Z\u0980-\u09FF]+)', val_str)
        if m:
            d_num = int(m.group(1))
            m_name = m.group(2).lower()
            m_num = MONTH_MAP.get(m_name)
            if m_num:
                try:
                    d = date(default_year, m_num, d_num)
                    return d, d
                except Exception:
                    pass

        return None, None

    def load_or_sync_holidays_to_db(self, db_session, holidays_list, year=2026):
        """
        Saves parsed holidays into the GovtHoliday database table.
        Prevents duplicates by (name, holiday_date).
        Returns: (inserted_count, updated_count)
        """
        from app import GovtHoliday
        inserted = 0
        updated = 0

        for h in holidays_list:
            s_date = h['start_date']
            e_date = h.get('end_date') or s_date
            h_name = h['name'].strip()
            h_type = h.get('type', 'General Holiday')
            h_year = s_date.year if s_date else year
            days = (e_date - s_date).days + 1 if e_date >= s_date else 1

            existing = GovtHoliday.query.filter(
                GovtHoliday.name == h_name,
                GovtHoliday.holiday_date == s_date
            ).first()

            if existing:
                existing.end_date = e_date
                existing.holiday_type = h_type
                existing.year = h_year
                existing.days_count = days
                updated += 1
            else:
                new_h = GovtHoliday(
                    name=h_name,
                    holiday_date=s_date,
                    end_date=e_date,
                    holiday_type=h_type,
                    year=h_year,
                    day_name=s_date.strftime('%A'),
                    days_count=days
                )
                db_session.add(new_h)
                inserted += 1

        db_session.commit()
        return inserted, updated

    def get_holiday_map_for_month(self, year, month):
        """
        Retrieves a date -> GovtHoliday mapping for quick O(1) lookup during matrix building.
        """
        from app import GovtHoliday
        import calendar
        days_in_month = calendar.monthrange(year, month)[1]
        m_start = date(year, month, 1)
        m_end = date(year, month, days_in_month)

        holidays = GovtHoliday.query.filter(
            GovtHoliday.year == year,
            GovtHoliday.holiday_date <= m_end,
            GovtHoliday.end_date >= m_start
        ).all()

        h_map = {}
        for h in holidays:
            cur = max(h.holiday_date, m_start)
            end = min(h.end_date or h.holiday_date, m_end)
            while cur <= end:
                h_map[cur] = h
                cur += timedelta(days=1)
        return h_map


govt_holiday_manager = GovtHolidayManager()
