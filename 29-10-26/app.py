"""LinkedIn Experience Copier - local service (v2.0, fast write-behind Excel mode)

How it works
------------
1. The extension POSTs a capture to /capture.
2. The service stores it in memory and appends one line to profiles.jsonl
   (a small crash-safe journal). This takes about a millisecond, so the
   extension gets its answer immediately.
3. A background thread writes all pending captures into profiles.xlsx in ONE
   batch (open once, save once). If Excel has the file open, it keeps retrying
   quietly; nothing is lost because the journal still holds the data.
4. On startup the journal is replayed, so a crash or power cut never loses
   captures.
"""
from functools import lru_cache
from pathlib import Path
from threading import Event, Lock, Thread
import atexit
from copy import copy
from datetime import datetime
import hashlib
import hmac
import json
import os
import re
import sqlite3
import time
import uuid

from bs4 import BeautifulSoup
from flask import Flask, jsonify, request
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

try:  # lxml is much faster than html.parser; optional.
    import lxml  # noqa: F401
    HTML_PARSER = "lxml"
except ImportError:
    HTML_PARSER = "html.parser"

APP_DIR = Path(__file__).resolve().parent
WORKBOOK = APP_DIR / "profiles.xlsx"
JOURNAL = APP_DIR / "profiles.jsonl"
LEGACY_DATABASE = APP_DIR / "profiles.sqlite3"  # Optional one-time import from v1.3-v1.5

HOST = "127.0.0.1"
PORT = 8765
# Prefer an environment variable. The fallback is your current token so the
# extension keeps working; if you change it, change it in the extension too.
TOKEN = os.environ.get("LI_COPIER_TOKEN") or "6b749a1d0d2b7449bc8c365b61970bacee7bc422df060f6c"

HEADERS = [
    "Name", "Email", "Link", "QA Status", "Primary Reason", "Secondary Reason", "QA Comments",
    "First Experience", "Job Title (As Per Experience)", "Company (As Per Experience)",
    "Tenure (Overall)", "Tenure (Current)", "Experience's Geo", "Geo (Header)",
    "Connections", "Followers", "Activity", "ID",
    "Report's Job Title", "JT Match", "Open to Work", "Keyword Found", "Dual Employment",
    "Remark", "Experience Section", "Capture Status", "Target Company", "Company Check",
    "Matching Experience Entry",
]
REPORT_HEADERS = [
    "Email", "QA Status", "Primary Reason", "Secondary Reason", "QA Comments", "Geo (Header)",
    "Connections", "Followers", "Activity", "ID", "Report's Job Title", "JT Match", "Open to Work",
    "Keyword Found", "Dual Employment", "Capture Status", "Target Company", "Company Check",
    "Matching Experience Entry",
]
RPF_SHEET_NAME = "CW RPF"
RPF_HEADERS = [
    "Email", "S.No", "Lead Link (as in file)", "First Name", "Last Name", "Job Title",
    "Company Name", "Overall Tenure in Current Company (all positions)",
    "Tenure in Current Profile (job title)", "Dual Employment – Company",
    "Dual Employment – Tenure", "Geo (with company)", "Geo (profile)", "Connections",
    "PDC", "Exclusion Keywords", "Recent Activity", "Profile URL", "Visited At",
]
HEADER_TITLES = {"name": "Name", "linkedin url": "Link", "experience section": "Experience Section",
                 "first experience": "First Experience", "company": "Company (As Per Experience)",
                 "job title": "Job Title (As Per Experience)", "experience's geo": "Experience's Geo",
                 "geo": "Geo (Header)", "secondary reason": "Secondary Reason", "id": "ID",
                 "tenure overall": "Tenure (Overall)", "tenure current": "Tenure (Current)",
                 "remark": "Remark", "report job title": "Report's Job Title", "jt match": "JT Match",
                 "open order": "Open Order"}
# Columns derived from the FIRST experience, ordered after First Experience in the report.
DERIVED_KEYS = ("first experience", "company", "job title", "experience's geo", "tenure overall", "tenure current", "remark")
DERIVED_ALIASES = {
    "first experience": ("first experience",),
    "company": ("company (as per experience)", "company", "first experience company"),
    "job title": ("job title (as per experience)", "job title", "title", "first experience job title"),
    "experience's geo": ("experience's geo", "experience geo", "first experience geo"),
    "tenure overall": ("tenure (overall)", "tenure [overall]", "tenure overall", "overall tenure"),
    "tenure current": ("tenure (current)", "tenure [current]", "tenure current", "current tenure"),
    "remark": ("remark", "remarks"),
}

MAX_EXPERIENCE_CHARS = 32000  # Excel cell limit is 32767

# Profiles whose Experience has no dates/duration at all get this label.
MISSING_TENURE_LABEL = "Tenurity Missing"
# False: label goes on the first line and the captured text stays below it.
# True : the cell contains ONLY the label.
MISSING_TENURE_ONLY = False

# Remark labels and thresholds (months). Change them here if the rules change.
REMARK_DEAD = "Dead Contact"            # first experience has no "Present": the person left that job
REMARK_MISSING = MISSING_TENURE_LABEL   # first experience has no tenure information
REMARK_PDC1 = "PDC 1"                   # overall tenure at the company is over 10 years
REMARK_PDC2 = "PDC 2"      # current role tenure is over 10 years (up to 20)
REMARK_PDC3 = "PDC 3"      # current role tenure is over 20 years
PDC1_OVERALL_MONTHS = 120
PDC2_CURRENT_MONTHS = 120
PDC3_CURRENT_MONTHS = 240

# Always-on screening terms. Extension Keyword Settings remain editable and are merged below.
FIXED_SCREENING_KEYWORDS = (
    "Contract", "Apprenticeship", "Freelance", "Freelancer",
    "Intern", "Internship", "Seasonal", "Retired", "Retire", "Ret.", "Retiree",
    "Former", "In remembrance", "In memory of", "Condolence", "Part-time",
    "Self-employed", "Open to Work", "#OpenToWork", "Seeking new opportunities",
    "Looking for new opportunities", "Career break", "End of an era",
    "End of the journey", "End of an amazing journey", "Temporary",
    "Unemployed", "Currently unemployed", "Between jobs", "Between roles",
    "In transition", "Career transition", "Job seeker", "Actively seeking",
    "Actively looking", "Available immediately", "Available for work",
    "Seeking employment", "Looking for a job", "Exploring opportunities",
    "Pursuing new opportunities", "New opportunities welcome", "Laid off",
    "Made redundant", "Recently laid off", "Recently made redundant",
    "Notice period", "Immediate joiner",
)

# Keep rows in the order the profile links were OPENED (not the order captures finish).
# Uses an extra "Open Order" column. Set KEEP_OPEN_ORDER = False to simply append rows.
KEEP_OPEN_ORDER = True
HIDE_ORDER_COLUMN = True   # hide that column so the sheet looks the same as before

# Headings of other LinkedIn profile sections. Used to cut off text that belongs to them.
SECTION_HEADINGS = {
    "about", "activity", "highlights", "analytics", "featured", "education", "skills", "languages",
    "interests", "projects", "courses", "recommendations", "honors & awards", "licenses & certifications",
    "volunteering", "publications", "causes", "patents", "test scores", "organizations",
    "people also viewed", "people you may know", "explore premium profiles", "more profiles for you",
}

BATCH_SECONDS = 0.35    # minimum coalescing window before writing staged captures
ADAPTIVE_FACTOR = 2.0   # wait this many times the LAST save duration before the next save, so a big
MAX_GAP_SECONDS = 8.0   # workbook never keeps the writer (and the GIL) busy 100% of the time
RETRY_SECONDS = 3       # retry delay when Excel is open/locked
POLL_SECONDS = 2.0      # the writer thread also checks (cheaply) whether the file changed on disk

WRAP = Alignment(wrap_text=False, vertical="top")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@app.before_request
def _start_request_timer():
    request.environ["li_copier_started"] = time.perf_counter()


@app.after_request
def _log_capture_duration(response):
    if request.path == "/capture":
        started = request.environ.get("li_copier_started", time.perf_counter())
        elapsed_ms = (time.perf_counter() - started) * 1000
        print(f"[CaptureTiming] /capture {elapsed_ms:.1f} ms; HTTP {response.status_code}; "
              f"request {request.content_length or 0} bytes", flush=True)
    return response

EXCEL_LOCK = Lock()   # protects every open/save of the workbook file
FLUSH_LOCK = Lock()   # only one flush at a time
STATE_LOCK = Lock()   # protects PENDING / KNOWN / journal
WAKE = Event()
LAST_EXCEL_ERROR = ""
PROFILE_ID_LOCK = Lock()
PROFILE_IDS = set()

WB_STAMP = None  # (mtime, size) of profiles.xlsx the last time KNOWN was rebuilt from it
CACHE_WB = None      # workbook kept in memory between flushes (read ONCE, saved many times)
CACHE_STAMP = None   # file stamp the cached workbook corresponds to
CACHE_SHEETS = {}    # sheet casefold -> {"ws", "cols", "rows": {url_key: row}, "next": next free row}
LAST_FLUSH_SECONDS = 0.0
REPORT_MIGRATION_PENDING = False
PENDING = {}  # (sheet_key, url_key) -> capture waiting to be written
KNOWN = {}    # (sheet_key, url_key) -> (name, experience_hash)


# --------------------------------------------------------------------------- helpers
def authorized():
    supplied = request.headers.get("X-LI-Copier-Token", "").encode("utf-8")
    return hmac.compare_digest(supplied, TOKEN.encode("utf-8"))


def sanitize(text):
    """Remove control characters that openpyxl refuses to write."""
    return ILLEGAL_CHARACTERS_RE.sub("", text or "")


def normalize_url(url):
    url = (url or "").strip().split("?")[0].split("#")[0].rstrip("/")
    # Canonicalize host (www., in., m., ...) so duplicates are detected reliably.
    url = re.sub(r"^https?://(?:[a-z]{1,3}\.)?linkedin\.com/in/", "https://www.linkedin.com/in/", url, flags=re.I)
    return url


def url_key(url):
    return normalize_url(str(url or "")).lower()


def digest(text):
    text = (text or "").strip()
    return hashlib.md5(text.encode("utf-8")).hexdigest() if text else ""


def internal_duplicate_identity(url, email):
    return f"{url_key(url)}|internal-duplicate|{digest(str(email or '').strip().casefold())}"


def clean_lines(text):
    text = sanitize(text)
    return [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]


def make_soup(markup):
    return BeautifulSoup(markup, HTML_PARSER)


def get_name(soup, page_title):
    title_name = re.sub(r"\s*\|\s*LinkedIn.*$", "", page_title or "", flags=re.I).strip()
    title_name = re.sub(r"^\(\d+\+?\)\s*", "", title_name)  # "(3) Jane Doe" notification prefix
    if title_name and title_name.lower() not in ("linkedin", "feed"):
        return sanitize(title_name)
    if soup is None:
        return ""
    h1 = soup.find("h1")
    if h1 and h1.get_text(" ", strip=True):
        return sanitize(h1.get_text(" ", strip=True))
    return ""


def find_experience_section(soup):
    # Look only at text nodes that are exactly "Experience" instead of calling get_text()
    # on every div/span of a huge page (that was O(elements x text)).
    for text_node in soup.find_all(string=lambda s: s is not None and s.strip().lower() == "experience"):
        node = text_node.parent
        fallback = None
        for _ in range(12):
            if node is None:
                break
            text = node.get_text(" ", strip=True)
            if len(text) >= 30 and node.find(["li", "article"]):
                if getattr(node, "name", None) == "section":
                    return node
                if fallback is None and len(text) < 25000:
                    fallback = node
            node = node.parent
        if fallback is not None:
            return fallback
    return None


# --------------------------------------------------------------------------- first experience
MARKDOWN_LINK_RE = re.compile(r"\[([^\]]*)\]\([^)]*\)")
_MON = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
_DATE = rf"(?:{_MON}\s+)?(?:19|20)\d{{2}}"
DATE_LINE_RE = re.compile(
    rf"^(?P<start>{_DATE})\s*[-\u2013\u2014]\s*(?P<end>present|{_DATE})(?:\s*[\u00b7\u2022|]\s*(?P<dur>.+))?\s*$", re.I)
_YRS_RE = re.compile(r"(\d+)\s*(?:yrs?|years?)\b", re.I)
_MOS_RE = re.compile(r"(\d+)\s*(?:mos?|months?)\b", re.I)
_DUR_ONLY_RE = re.compile(r"^(?:\d+\s*(?:yrs?|years?)\s*)?(?:\d+\s*(?:mos?|months?))?$|^less than a year$", re.I)
_NOISE_RE = re.compile(r"^(present|show (all|more)\b.*|see more|[.\u2026]+\s*see more)$", re.I)
_MODE_RE = re.compile(r"^(on-site|onsite|hybrid|remote)$", re.I)


def parse_duration(text):
    """'2 yrs 7 mos' -> (31, '2 yrs 7 mos').  Returns (None, '') when there is no duration."""
    text = text or ""
    y, m = _YRS_RE.search(text), _MOS_RE.search(text)
    if y or m:
        months = (int(y.group(1)) * 12 if y else 0) + (int(m.group(1)) if m else 0)
        return months, " ".join(x.group(0) for x in (y, m) if x)
    if re.search(r"less than a year", text, re.I):
        return 0, "Less than a year"
    return None, ""


def _split_parts(line):
    return [p.strip() for p in re.split(r"\s*[\u00b7\u2022]\s*", line) if p.strip()]


def _is_total_line(line):
    """A company's overall-tenure line, e.g. 'Full-time · 3 yrs 2 mos' or '3 yrs 2 mos'."""
    if DATE_LINE_RE.match(line):
        return False
    parts = _split_parts(line)
    return bool(parts) and bool(parts[-1].strip()) and bool(_DUR_ONLY_RE.match(parts[-1].strip())) \
        and parse_duration(parts[-1])[1] != ""


def _is_location_or_mode(line):
    if len(line) > 60 or line.endswith("."):
        return False
    return bool(_MODE_RE.match(line)) or line.count(",") >= 2 \
        or bool(re.search(r"\b(area|region|metropolitan|remote)\b", line, re.I))


EMPLOYMENT_TYPE_RE = re.compile(
    r"\b(full[- ]time|part[- ]time|self[- ]employed|freelance|contract|internship|apprenticeship|temporary|seasonal|volunteer)\b",
    re.I)


def _company_from_employment_line(line):
    if "·" not in (line or ""):
        return ""
    company, detail = line.split("·", 1)
    if not company.strip() or not EMPLOYMENT_TYPE_RE.search(detail):
        return ""
    return re.sub(r"\s+", " ", company).strip(" ·|")


def _company_for_date(lines, date_index):
    # Standalone role: Title / Company · Full-time / dates, or Company / Full-time / dates.
    for idx in range(date_index - 1, max(-1, date_index - 8), -1):
        company = _company_from_employment_line(lines[idx])
        if company:
            return company
        if EMPLOYMENT_TYPE_RE.fullmatch(lines[idx].strip()) and idx > 0:
            candidate = re.sub(r"\s+", " ", lines[idx - 1]).strip(" ·|")
            if candidate and not DATE_LINE_RE.match(candidate) and not _is_location_or_mode(candidate) and not _is_experience_geo_line(candidate):
                return candidate
    # Grouped company roles: Company / Full-time · total tenure / location / title / dates.
    for idx in range(date_index - 1, max(-1, date_index - 8), -1):
        if _is_total_line(lines[idx]) and idx > 0:
            company = re.sub(r"\s+", " ", lines[idx - 1]).strip(" ·|")
            if company and not _is_location_or_mode(company) and not _is_experience_geo_line(company) and not DATE_LINE_RE.match(company):
                return company
    return ""


def _normalized_company(value):
    value = str(value or "").casefold()
    value = re.sub(r"\b(incorporated|inc|limited|ltd|llc|plc|corp|corporation|company|co)\b", " ", value)
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def _company_match_kind(target_key, company_key):
    """Return exact or conservative whole-word prefix match; never use broad fuzzy matching."""
    if not target_key or not company_key:
        return ""
    if company_key == target_key:
        return "exact"
    if company_key.startswith(target_key + " "):
        return "prefix"
    return ""


def _target_company_validation(experience, target_company, first_company=""):
    """Match company names first, then classify employment from each matched role's dates."""
    target_company = sanitize(str(target_company or "")).strip()[:320]
    if not target_company:
        return {"target_company": "", "check": "", "entries": ""}

    target_key = _normalized_company(target_company)
    lines = _entry_lines(experience or "")
    date_rows = [(i, DATE_LINE_RE.match(line)) for i, line in enumerate(lines) if DATE_LINE_RE.match(line)]
    if not date_rows:
        first_key = _normalized_company(first_company)
        entries = "Experience 1" if _company_match_kind(target_key, first_key) else ""
        return {"target_company": target_company, "check": "Needs Review", "entries": entries}

    roles = []
    unknown_company = False
    for entry_number, (line_index, date_match) in enumerate(date_rows, start=1):
        # The report's Company column is derived by analyze_first_experience; reuse it
        # for the first entry so validation agrees with the value already written to Excel.
        company = first_company if entry_number == 1 else _company_for_date(lines, line_index)
        company_key = _normalized_company(company)
        if not company_key:
            unknown_company = True
            continue
        roles.append({"number": entry_number, "company": company, "key": company_key,
                      "current": date_match.group("end").casefold() == "present"})

    exact_keys = {role["key"] for role in roles if _company_match_kind(target_key, role["key"]) == "exact"}
    if exact_keys:
        matched_keys = exact_keys
    else:
        prefix_keys = {role["key"] for role in roles if _company_match_kind(target_key, role["key"]) == "prefix"}
        if len(prefix_keys) > 1:
            return {"target_company": target_company, "check": "Needs Review", "entries": ""}
        matched_keys = prefix_keys

    matches = [role for role in roles if role["key"] in matched_keys]
    entry_label = ", ".join(f"Experience {match['number']}" for match in matches)
    if not matches:
        return {"target_company": target_company, "check": "Needs Review", "entries": ""}

    # Employment state is assessed only after company matching and uses dates
    # from the specific matched Experience entries.
    current_match = any(match["current"] for match in matches)
    past_match = any(not match["current"] for match in matches)
    if current_match and past_match:
        check = "Dual Employment"
    elif unknown_company:
        check = "Needs Review"
    elif current_match:
        check = "Pass"
    else:
        check = "Dead Contact"
    return {"target_company": target_company, "check": check, "entries": entry_label}


def _title_before(lines, boundary):
    for idx in range(boundary - 1, max(-1, boundary - 5), -1):
        line = lines[idx]
        if DATE_LINE_RE.match(line) or _is_total_line(line) or _is_location_or_mode(line) or _is_experience_geo_line(line):
            continue
        if _company_from_employment_line(line):
            continue
        if line.casefold() not in {"experience", "full-time", "part-time"}:
            return line
    return ""


def _entry_lines(text):
    out = []
    for ln in (text or "").splitlines():
        ln = MARKDOWN_LINK_RE.sub(r"\1", ln)
        ln = re.sub(r"^\s*[*\u2022-]\s+", "", ln)            # leading bullet
        ln = re.sub(r"\s+", " ", ln).strip()
        if ln and not _NOISE_RE.match(ln):
            out.append(ln)
    while out and out[0].lower() in ("experience", REMARK_MISSING.lower()):
        out.pop(0)
    return out


def _is_experience_geo_line(line):
    """Recognize a structured location line without mistaking work mode for a location."""
    value = re.sub(r"\s+", " ", str(line or "")).strip()
    if not value or len(value) > 100 or value.casefold() in {"experience", "location"}:
        return False
    if DATE_LINE_RE.match(value) or _is_total_line(value) or _company_from_employment_line(value):
        return False
    if _MODE_RE.fullmatch(value) or (EMPLOYMENT_TYPE_RE.fullmatch(value) and "·" not in value):
        return False
    if re.search(r"\b(?:on[- ]site|hybrid|remote)\b", value, re.I) and "," not in value:
        return False
    return "," in value or bool(re.search(r"\b(?:area|region|metropolitan)\b", value, re.I))


def _first_experience_geo(lines, date_index=None, company="", job_title=""):
    """Find the location attached to the first/current role using the already captured text."""
    if date_index is None:
        candidates = lines[:8]
    else:
        # Grouped roles put location between company and title; standalone roles often
        # put it immediately after the date line. Prefer the closest pre-date location.
        candidates = list(reversed(lines[max(0, date_index - 6):date_index])) + lines[date_index + 1:date_index + 4]
    excluded = {re.sub(r"\s+", " ", str(value or "")).strip().casefold()
                for value in (company, job_title) if value}
    for line in candidates:
        value = re.sub(r"\s+", " ", line).strip()
        if value.casefold() in excluded or DATE_LINE_RE.match(value) or _is_total_line(value):
            continue
        if EMPLOYMENT_TYPE_RE.search(value) or _MODE_RE.fullmatch(value):
            continue
        if _is_experience_geo_line(value):
            return value
    return ""


def analyze_first_experience(text):
    """Parse the top Experience entry and its company, title, location, tenure, and QA remark."""
    res = {"first_exp": "", "company": "", "job_title": "", "experience_geo": "", "overall": "", "current": "", "remark": "", "parsed": False}
    lines = _entry_lines(text)
    if not lines:
        return res
    d = next((i for i, ln in enumerate(lines) if DATE_LINE_RE.match(ln)), None)
    if d is None:
        res.update(first_exp="\n".join(lines[:2]), remark=REMARK_MISSING, parsed=True)
        if len(lines) >= 2 and _is_total_line(lines[1]):
            res["company"] = lines[0]
            res["job_title"] = next((line for line in lines[2:] if not _is_location_or_mode(line) and not _is_experience_geo_line(line)), "")
        else:
            company_idx = next((i for i, line in enumerate(lines[:5]) if _company_from_employment_line(line)), None)
            if company_idx is not None:
                res["company"] = _company_from_employment_line(lines[company_idx])
                res["job_title"] = _title_before(lines, company_idx)
        res["experience_geo"] = _first_experience_geo(lines, None, res["company"], res["job_title"])
        return res

    grouped = d >= 3 and _is_total_line(lines[1])
    m = DATE_LINE_RE.match(lines[d])
    is_current = m.group("end").lower() == "present"
    cur_months, cur_disp = parse_duration(m.group("dur") or "")
    ov_months, ov_disp = (None, "")
    if grouped:
        ov_months, ov_disp = parse_duration(_split_parts(lines[1])[-1])
        start = 0
        company = lines[0]
        job_title = _title_before(lines, d)
    else:
        company_idx = next((i for i in range(d - 1, max(-1, d - 5), -1)
                            if _company_from_employment_line(lines[i])), None)
        if company_idx is not None:
            company = _company_from_employment_line(lines[company_idx])
            job_title = _title_before(lines, company_idx)
            start = max(0, d - 2)
        else:
            start = max(0, d - 2)
            job_title = lines[start] if start < d else ""
            company_line = lines[start + 1] if start + 1 < d else ""
            company = company_line.split("·", 1)[0].strip()
    end = d + 1
    while end < len(lines) and end - (d + 1) < 3 and _is_location_or_mode(lines[end]):
        end += 1

    res.update(first_exp="\n".join(lines[start:end]), company=company, job_title=job_title,
               experience_geo=_first_experience_geo(lines, d, company, job_title),
               overall=ov_disp, current=cur_disp, parsed=True)
    if not is_current:
        res["remark"] = REMARK_DEAD
    elif cur_months is None and ov_months is None:
        res["remark"] = REMARK_MISSING
    elif cur_months is not None and cur_months > PDC3_CURRENT_MONTHS:
        res["remark"] = REMARK_PDC3
    elif cur_months is not None and cur_months > PDC2_CURRENT_MONTHS:
        res["remark"] = REMARK_PDC2
    elif ov_months is not None and ov_months > PDC1_OVERALL_MONTHS:
        res["remark"] = REMARK_PDC1
    return res


def apply_tenure_flag(experience, analysis):
    """Put 'Tenurity Missing' on the first line when the FIRST experience has no tenure."""
    if analysis.get("parsed"):
        flagged = analysis["remark"] == REMARK_MISSING
    else:                                           # could not read the structure: old whole-text rule
        return flag_missing_tenure(experience)
    if not flagged or experience.startswith(MISSING_TENURE_LABEL):
        return experience, flagged
    if MISSING_TENURE_ONLY:
        return MISSING_TENURE_LABEL, True
    return f"{MISSING_TENURE_LABEL}\n{experience}", True


def state_hash(experience, derived):
    """Fingerprint the Experience text and all derived first-experience fields."""
    parts = [str(value or "").strip() for value in derived]
    return digest((experience or "") + "\x1f" + "\x1f".join(parts))


def trim_to_experience(lines):
    """Start at the 'Experience' heading and stop at the next profile section (Education, Skills...).
    Safety net: even if the browser sends too much of the page, only Experience is stored."""
    start = next((i for i, ln in enumerate(lines) if ln.strip().lower() == "experience"), None)
    if start is None:
        return lines
    out = [lines[start]]
    for ln in lines[start + 1:]:
        if ln.strip().lower() in SECTION_HEADINGS:
            break
        out.append(ln)
    return out


def extract_experience(experience_text, experience_html, page_html):
    """Returns (text, too_broad). too_broad = the browser sent more than the Experience section."""
    text = (experience_text or "").strip()
    if not text and experience_html:
        text = make_soup(experience_html).get_text("\n", strip=True)
    if not text and page_html:
        section = find_experience_section(make_soup(page_html))
        if section is not None:
            text = section.get_text("\n", strip=True)
    all_lines = clean_lines(text)
    kept = trim_to_experience(all_lines)
    deduped = []
    for line in kept:
        if not deduped or line != deduped[-1]:
            deduped.append(line)
    # An empty Experience is a valid outcome: the profile URL/name are still stored.
    return "\n".join(deduped).strip()[:MAX_EXPERIENCE_CHARS], len(kept) < len(all_lines)


def extract_experience_section(experience_text, experience_html, page_html):
    return extract_experience(experience_text, experience_html, page_html)[0]


_MONTH = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
TENURE_RE = re.compile(
    rf"\b{_MONTH}\s+(?:19|20)\d{{2}}\b"            # Oct 2023
    r"|\b(?:19|20)\d{2}\s*[-\u2013\u2014]\s*(?:(?:19|20)\d{2}|present)\b"  # 2019 - 2021 / 2019 - Present
    r"|\bpresent\b"                                # ... - Present
    r"|\b\d+\s*(?:yrs?|mos?|years?|months?)\b",     # 3 yrs 1 mo
    re.I)


def flag_missing_tenure(experience):
    """If the Experience has content but no tenure anywhere, mark it. Returns (text, flagged)."""
    body = [ln for ln in experience.splitlines() if ln.strip() and ln.strip().lower() != "experience"]
    if not body or TENURE_RE.search(experience) or experience.startswith(MISSING_TENURE_LABEL):
        return experience, False
    if MISSING_TENURE_ONLY:
        return MISSING_TENURE_LABEL, True
    return f"{MISSING_TENURE_LABEL}\n{experience}", True


# --------------------------------------------------------------------------- workbook
def normalized_header(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def new_profile_id():
    """Create a readable local timestamp ID, suffixing collisions within the same second."""
    prefix = datetime.now().strftime("%d-%m-%y-%H-%M-%S")
    with PROFILE_ID_LOCK:
        candidate = prefix
        suffix = 2
        while candidate in PROFILE_IDS:
            candidate = f"{prefix}-{suffix:02d}"
            suffix += 1
        PROFILE_IDS.add(candidate)
    return candidate


def _reorder_worksheet_columns(ws):
    """Place standard report columns in canonical order, preserving extra columns and cell formatting."""
    max_col, max_row = ws.max_column, max(ws.max_row, 1)
    wanted = [normalized_header(title) for title in HEADERS]
    by_header = {}
    for col in range(1, max_col + 1):
        key = normalized_header(ws.cell(1, col).value)
        if key:
            by_header.setdefault(key, col)
    ordered = [by_header[key] for key in wanted if key in by_header]
    ordered.extend(col for col in range(1, max_col + 1) if col not in ordered)
    if ordered == list(range(1, max_col + 1)):
        return

    rows = []
    for row in range(1, max_row + 1):
        row_data = []
        for col in ordered:
            cell = ws.cell(row, col)
            row_data.append((cell.value, copy(cell._style), copy(cell.hyperlink), copy(cell.comment)))
        rows.append(row_data)
    dimensions = {}
    for old_col in range(1, max_col + 1):
        dim = ws.column_dimensions[get_column_letter(old_col)]
        dimensions[old_col] = (dim.width, dim.hidden, dim.bestFit, dim.outlineLevel, dim.collapsed)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            cell.value = None
            cell._style = None
            cell.hyperlink = None
            cell.comment = None
    for key in list(ws.column_dimensions):
        del ws.column_dimensions[key]

    for new_col, old_col in enumerate(ordered, start=1):
        dim = ws.column_dimensions[get_column_letter(new_col)]
        dim.width, dim.hidden, dim.bestFit, dim.outlineLevel, dim.collapsed = dimensions[old_col]
        for row_idx, row_data in enumerate(rows, start=1):
            value, style, hyperlink, comment = row_data[new_col - 1]
            cell = ws.cell(row_idx, new_col)
            cell.value = value
            cell._style = style
            if hyperlink:
                hyperlink.ref = cell.coordinate
                cell._hyperlink = hyperlink
            cell.comment = comment


def set_text(cell, value):
    """Write text; never let a value starting with '=' become an Excel formula."""
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


def _min_width(ws, col, width):
    letter = get_column_letter(col)
    ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, width)


def normalize_sheet_name(value, allow_profiles=True):
    name = str(value or "").strip()
    if not name:
        raise ValueError("Sheet name is required.")
    if len(name) > 31:
        raise ValueError("Sheet names can be at most 31 characters.")
    if re.search(r'[\\/*?:\[\]]', name) or name.startswith("'") or name.endswith("'"):
        raise ValueError("Sheet names cannot contain \\ / * ? : [ ] or begin/end with an apostrophe.")
    if name.casefold() == "history":
        raise ValueError("The sheet name 'History' is reserved by Excel.")
    if name.casefold() == RPF_SHEET_NAME.casefold():
        raise ValueError(f"The sheet name '{RPF_SHEET_NAME}' is reserved for the RPF export.")
    if not allow_profiles and name.casefold() == "profiles":
        raise ValueError("Choose a batch name other than the main Profiles sheet.")
    return name


def prepare_worksheet(ws):
    ws.sheet_format.defaultRowHeight = 14.4
    for row_idx in range(1, ws.max_row + 1):
        if ws.row_dimensions[row_idx].height is not None:
            ws.row_dimensions[row_idx].height = None
    # A truly empty sheet gets headers in A:C (not B:D).
    if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
        for i, title in enumerate(HEADERS, start=1):
            ws.cell(1, i).value = title

    header_map = {}
    for col in range(1, ws.max_column + 1):
        key = normalized_header(ws.cell(1, col).value)
        if key:
            header_map.setdefault(key, col)

    aliases = {
        "name": ("name", "profile name"),
        "linkedin url": ("link", "linkedin url", "profile url", "linkedin profile url", "url"),
        "experience section": ("experience section", "current experience", "experience"),
    }
    columns = {}
    for required, names in aliases.items():
        found = next((header_map[n] for n in names if n in header_map), None)
        if found is None:
            found = ws.max_column + 1
            ws.cell(1, found).value = HEADER_TITLES[required]
            header_map[normalized_header(HEADER_TITLES[required])] = found
        columns[required] = found

    prev = columns["experience section"]
    for key in DERIVED_KEYS:
        found = next((header_map[n] for n in DERIVED_ALIASES[key] if n in header_map), None)
        if found is None:
            at = prev + 1
            ws.insert_cols(at)
            ws.cell(1, at).value = HEADER_TITLES[key]
            for k, c in list(columns.items()):
                if c >= at:
                    columns[k] = c + 1
            for k, c in list(header_map.items()):
                if c >= at:
                    header_map[k] = c + 1
            found = at
            header_map[normalized_header(HEADER_TITLES[key])] = found
        ws.cell(1, found).value = HEADER_TITLES[key]
        header_map[normalized_header(HEADER_TITLES[key])] = found
        columns[key] = found
        prev = found

    # Preserve old profile-header Geo values under the clarified display label.
    old_geo_col = header_map.get("geo")
    canonical_geo_col = header_map.get(normalized_header(HEADER_TITLES["geo"]))
    if canonical_geo_col is None and old_geo_col is not None:
        ws.cell(1, old_geo_col).value = HEADER_TITLES["geo"]
        header_map[normalized_header(HEADER_TITLES["geo"])] = old_geo_col

    for key, width in (("first experience", 60), ("company", 30), ("job title", 38), ("experience's geo", 30), ("tenure overall", 18), ("tenure current", 18), ("remark", 22)):
        dim = ws.column_dimensions[get_column_letter(columns[key])]
        dim.width = max(dim.width or 0, width)
        dim.hidden = False

    if KEEP_OPEN_ORDER:
        found = header_map.get("open order")
        if found is None:
            found = max(ws.max_column, *columns.values()) + 1
            ws.cell(1, found).value = HEADER_TITLES["open order"]
            header_map["open order"] = found
        columns["open order"] = found
        letter = get_column_letter(found)
        ws.column_dimensions[letter].width = 14
        ws.column_dimensions[letter].hidden = HIDE_ORDER_COLUMN

    for title in REPORT_HEADERS:
        key = normalized_header(title)
        if key not in header_map:
            col = ws.max_column + 1
            ws.cell(1, col).value = title
            header_map[key] = col
        columns[key] = header_map[key]
    if columns.get("linkedin url"):
        ws.cell(1, columns["linkedin url"]).value = "Link"

    _reorder_worksheet_columns(ws)
    header_map = {}
    for col in range(1, ws.max_column + 1):
        key = normalized_header(ws.cell(1, col).value)
        if key:
            header_map.setdefault(key, col)
    for required, names in aliases.items():
        found = next((header_map[name] for name in names if name in header_map), None)
        if found is not None:
            columns[required] = found
    for key in DERIVED_KEYS:
        columns[key] = header_map[normalized_header(HEADER_TITLES[key])]
    for title in REPORT_HEADERS:
        columns[normalized_header(title)] = header_map[normalized_header(title)]
    if KEEP_OPEN_ORDER and "open order" in header_map:
        columns["open order"] = header_map["open order"]

    _min_width(ws, columns["report's job title"], 38)
    _min_width(ws, columns["jt match"], 12)
    ws.freeze_panes = "A2"
    _min_width(ws, columns["name"], 28)
    _min_width(ws, columns["linkedin url"], 48)
    _min_width(ws, columns["experience section"], 80)
    _min_width(ws, columns["target company"], 28)
    _min_width(ws, columns["company check"], 22)
    _min_width(ws, columns["matching experience entry"], 32)

    return columns


def _is_rpf_sheet(name):
    return str(name or "").strip().casefold() == RPF_SHEET_NAME.casefold()


def _rpf_profile_key(email, profile_url):
    email_key = str(email or "").strip().casefold()
    url_key_value = url_key(profile_url)
    if not email_key and not url_key_value:
        return None
    return email_key, url_key_value


def _rpf_visited_at(profile_id, open_order):
    try:
        stamp = float(open_order)
        if stamp > 0:
            return datetime.fromtimestamp(stamp / 1000 if stamp > 100_000_000_000 else stamp)
    except (TypeError, ValueError, OSError, OverflowError):
        pass
    profile_id = str(profile_id or "").strip()
    match = re.match(r"^(\d{2}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})(?:-\d+)?$", profile_id)
    if match:
        try:
            return datetime.strptime(match.group(1), "%d-%m-%y-%H-%M-%S")
        except ValueError:
            pass
    return None


def _rpf_name_parts(name):
    parts = re.sub(r"\s+", " ", str(name or "").strip()).split(" ")
    return (parts[0], " ".join(parts[1:])) if parts else ("", "")


def _rpf_date_sort_key(value):
    text = str(value or "").strip().replace(".", "")
    month_year = re.match(r"(?i)^([a-z]{3,9})\s+((?:19|20)\d{2})$", text)
    if month_year:
        month_names = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        months = {name.casefold(): number for number, name in enumerate(month_names, start=1)}
        month = months.get(month_year.group(1)[:3].casefold(), 1)
        return int(month_year.group(2)), month
    year = re.search(r"(?:19|20)\d{2}", text)
    return (int(year.group(0)), 1) if year else (0, 0)


def _rpf_format_role_tenure(start, end, duration, suffix=""):
    start = re.sub(r"\s+", " ", str(start or "").strip())
    end = re.sub(r"\s+", " ", str(end or "").strip())
    duration = re.sub(r"\s+", " ", str(duration or "").strip())
    if end.casefold() == "present":
        end = "Present"
    if start and end and duration:
        return f"{start} – {end} = {duration}{suffix}"
    return duration


def _rpf_tenure_fields(experience, company, overall_tenure, current_tenure):
    lines = _entry_lines(str(experience or ""))
    dated_roles = []
    company_key = _normalized_company(company)
    for index, line in enumerate(lines):
        match = DATE_LINE_RE.match(line)
        if not match:
            continue
        dated_roles.append({
            "start": match.group("start"),
            "end": match.group("end"),
            "duration": (match.group("dur") or "").strip(),
            "company": _company_for_date(lines, index),
            "order": index,
        })

    current_profile = ""
    if dated_roles:
        first_role = dated_roles[0]
        current_profile = _rpf_format_role_tenure(
            first_role["start"], first_role["end"], first_role["duration"] or current_tenure
        )
    if not current_profile:
        current_profile = str(current_tenure or "")

    company_roles = [
        role for role in dated_roles
        if company_key and _normalized_company(role["company"]) == company_key
    ]
    company_tenure = str(overall_tenure or "")
    if company_roles:
        earliest = min(company_roles, key=lambda role: (_rpf_date_sort_key(role["start"]), role["order"]))
        present_roles = [role for role in company_roles if role["end"].casefold() == "present"]
        if present_roles:
            end = "Present"
        else:
            latest = max(company_roles, key=lambda role: (_rpf_date_sort_key(role["end"]), role["order"]))
            end = latest["end"]
        duration = company_tenure or (company_roles[0]["duration"] if len(company_roles) == 1 else "Not Available")
        suffix = f" ({len(company_roles)} positions)" if len(company_roles) > 1 else ""
        formatted = _rpf_format_role_tenure(earliest["start"], end, duration, suffix)
        if formatted:
            company_tenure = formatted
    return company_tenure, current_profile


def _rpf_dual_employment(experience, primary_company, dual_flag):
    lines = _entry_lines(str(experience or ""))
    primary_key = _normalized_company(primary_company)
    employers = {}
    for index, line in enumerate(lines):
        match = DATE_LINE_RE.match(line)
        if not match or match.group("end").casefold() != "present":
            continue
        company = _company_for_date(lines, index)
        company_key = _normalized_company(company)
        if not company or not company_key or company_key == primary_key:
            continue
        entry = employers.setdefault(company_key, {"company": company, "tenures": []})
        duration = (match.group("dur") or "").strip()
        if duration and duration not in entry["tenures"]:
            entry["tenures"].append(duration)
    if not employers:
        return ("Needs Review", "") if str(dual_flag or "").strip().casefold() == "yes" else ("No", "")
    companies = []
    tenures = []
    for item in employers.values():
        companies.append(item["company"])
        tenures.append(", ".join(item["tenures"]) if item["tenures"] else "Not Available")
    return "; ".join(companies), "; ".join(tenures)


def _rpf_pdc_value(remark, qa_status, primary_reason, capture_status):
    remark = str(remark or "").strip()
    status = str(qa_status or "").strip().casefold()
    primary = str(primary_reason or "").strip()
    if remark.casefold() == "pdc 1" or primary.casefold() == "pdc 1" or status == "qualified pdc 1":
        return "PDC 1 (No dead info – verified by QA)"
    if remark.casefold() in {"pdc 2", "pdc 3"}:
        return remark
    if remark.casefold() == "dead contact":
        return "Dead Contact"
    if remark.casefold() == MISSING_TENURE_LABEL.casefold():
        return f"Review – {MISSING_TENURE_LABEL}"
    if status == "qualified":
        return "Not PDC"
    if status in {"review required", "unable to capture"}:
        reason = primary or str(capture_status or "").strip() or "reason unavailable"
        return f"Review – {reason}"
    if status == "disqualified":
        return primary or remark
    if remark:
        return remark
    return "Not PDC"


def _rpf_values_from_profile_row(ws, columns, row):
    def value(key):
        col = columns.get(key)
        return ws.cell(row, col).value if col else None

    name = str(value("name") or "")
    first_name, last_name = _rpf_name_parts(name)
    profile_url = value("linkedin url") or ""
    company = value("company") or ""
    overall_tenure, current_tenure = _rpf_tenure_fields(
        value("experience section"), company, value("tenure overall"), value("tenure current")
    )
    dual_company, dual_tenure = _rpf_dual_employment(
        value("experience section"), company, value("dual employment")
    )
    return {
        "Email": value("email") or "",
        "Lead Link (as in file)": profile_url,
        "First Name": first_name,
        "Last Name": last_name,
        "Job Title": value("job title") or "",
        "Company Name": company,
        "Overall Tenure in Current Company (all positions)": overall_tenure,
        "Tenure in Current Profile (job title)": current_tenure,
        "Dual Employment – Company": dual_company,
        "Dual Employment – Tenure": dual_tenure,
        "Geo (with company)": value("experience's geo") or "",
        "Geo (profile)": value("geo (header)") or "",
        "Connections": value("connections") or "",
        "PDC": _rpf_pdc_value(value("remark"), value("qa status"), value("primary reason"), value("capture status")),
        "Exclusion Keywords": value("keyword found") or "",
        "Recent Activity": value("activity") or "",
        "Profile URL": profile_url,
        "Visited At": _rpf_visited_at(value("id"), value("open order")),
    }


def _ensure_rpf_worksheet(wb):
    ws = next((sheet for sheet in wb.worksheets if _is_rpf_sheet(sheet.title)), None)
    created = False
    if ws is None:
        ws = wb.create_sheet(RPF_SHEET_NAME)
        created = True
    current = [normalized_header(ws.cell(1, col).value) for col in range(1, len(RPF_HEADERS) + 1)]
    expected = [normalized_header(header) for header in RPF_HEADERS]
    has_content = any(cell.value not in (None, "") for row in ws.iter_rows() for cell in row)
    if not any(current) and not has_content:
        for col, header in enumerate(RPF_HEADERS, start=1):
            cell = ws.cell(1, col)
            cell.value = header
            font = copy(cell.font)
            font.bold = True
            cell.font = font
        created = True
    elif current != expected:
        raise ValueError(
            f"Worksheet '{RPF_SHEET_NAME}' already exists with a different layout. "
            "It was left unchanged; rename or clear that sheet before retrying."
        )
    if any(ws.cell(1, col).value not in (None, "") for col in range(len(RPF_HEADERS) + 1, ws.max_column + 1)):
        raise ValueError(
            f"Worksheet '{RPF_SHEET_NAME}' has extra columns after the expected RPF schema. "
            "It was left unchanged."
        )
    for col, width in enumerate((30, 9, 48, 22, 24, 38, 32, 38, 34, 32, 26, 30, 30, 14, 38, 42, 20, 48, 22), start=1):
        letter = get_column_letter(col)
        if not ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.sheet_format.defaultRowHeight = 14.4
    ws.auto_filter.ref = f"A1:S{max(ws.max_row, 1)}"
    return ws, created


def _rpf_row_index(ws):
    index = {}
    next_serial = 1
    for row in range(2, ws.max_row + 1):
        key = _rpf_profile_key(ws.cell(row, 1).value, ws.cell(row, 18).value)
        if key:
            index.setdefault(key, row)
        serial = ws.cell(row, 2).value
        if isinstance(serial, (int, float)):
            next_serial = max(next_serial, int(serial) + 1)
        elif str(serial or "").strip().isdigit():
            next_serial = max(next_serial, int(str(serial).strip()) + 1)
    return index, next_serial


def _upsert_rpf_row(source_ws, columns, source_row, context):
    values = _rpf_values_from_profile_row(source_ws, columns, source_row)
    key = _rpf_profile_key(values["Email"], values["Profile URL"])
    if key is None:
        return
    rpf_ws = context["ws"]
    row = context["rows"].get(key)
    if row is None:
        row = max(rpf_ws.max_row + 1, 2)
        context["rows"][key] = row
        rpf_ws.cell(row, 2).value = context["next_serial"]
        context["next_serial"] += 1
        context["changed"] = True
    elif rpf_ws.cell(row, 2).value in (None, ""):
        rpf_ws.cell(row, 2).value = context["next_serial"]
        context["next_serial"] += 1
        context["changed"] = True
    for col, header in enumerate(RPF_HEADERS, start=1):
        if header == "S.No":
            continue
        cell = rpf_ws.cell(row, col)
        new_value = values[header]
        if cell.value != new_value:
            if header == "Visited At":
                cell.value = new_value
                cell.number_format = "dd-mm-yy hh:mm:ss"
            else:
                set_text(cell, "" if new_value is None else str(new_value))
            context["changed"] = True
        alignment = copy(cell.alignment)
        alignment.wrap_text = False
        cell.alignment = alignment
        rpf_ws.row_dimensions[row].height = 14.4
    rpf_ws.auto_filter.ref = f"A1:S{max(rpf_ws.max_row, 1)}"


def _rpf_context(wb):
    ws, created = _ensure_rpf_worksheet(wb)
    rows, next_serial = _rpf_row_index(ws)
    return {"ws": ws, "rows": rows, "next_serial": next_serial, "changed": created}


def _sync_rpf_sheet(wb, prepared_columns=None):
    context = _rpf_context(wb)
    prepared_columns = prepared_columns or {}
    for source_ws in wb.worksheets:
        if _is_rpf_sheet(source_ws.title):
            continue
        columns = prepared_columns.get(source_ws.title) or prepare_worksheet(source_ws)
        for row in range(2, source_ws.max_row + 1):
            if not source_ws.cell(row, columns["linkedin url"]).value and not source_ws.cell(row, columns["email"]).value:
                continue
            _upsert_rpf_row(source_ws, columns, row, context)
    return context["changed"]


def open_or_create_workbook(sheet_name=None):
    if WORKBOOK.exists():
        wb = load_workbook(WORKBOOK)
    else:
        wb = Workbook()
        base = wb.active
        base.title = "Profiles"
        base.append(HEADERS)

    if sheet_name:
        requested = normalize_sheet_name(sheet_name)
        actual = next((name for name in wb.sheetnames if name.casefold() == requested.casefold()), None)
        if actual:
            ws = wb[actual]
        else:
            ws = wb.create_sheet(requested)
            ws.append(HEADERS)
    else:
        ws = wb["Profiles"] if "Profiles" in wb.sheetnames else wb.active
    return wb, ws, prepare_worksheet(ws)

def atomic_save_workbook(wb):
    temp_path = APP_DIR / f"profiles_tmp_{uuid.uuid4().hex}.xlsx"
    try:
        wb.save(temp_path)
        # Excel locks the destination while the workbook is open. Retry briefly.
        last_exc = None
        for attempt in range(3):
            try:
                os.replace(temp_path, WORKBOOK)
                return
            except PermissionError as exc:
                last_exc = exc
                time.sleep(0.15 * (attempt + 1))
        if last_exc:
            raise last_exc
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass


def scan_workbook(ws, columns):
    """One fast pass over the sheet. Returns ({key: row}, {key: (name, exp_hash)}, {row: open_order})."""
    name_i = columns["name"] - 1
    url_i = columns["linkedin url"] - 1
    exp_i = columns["experience section"] - 1
    id_i = columns["id"] - 1
    ord_i = columns["open order"] - 1 if columns.get("open order") else None
    der_i = [columns[k] - 1 for k in DERIVED_KEYS if k in columns]
    rows, known, orders = {}, {}, {}
    profile_ids = set()
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if id_i < len(row) and row[id_i]:
            profile_ids.add(str(row[id_i]).strip())
        if ord_i is not None and ord_i < len(row) and isinstance(row[ord_i], (int, float)):
            orders[r] = float(row[ord_i])
        if url_i >= len(row):
            continue
        key = url_key(row[url_i])
        if not key:
            continue
        reason_i = columns.get("primary reason", 0) - 1
        email_i = columns.get("email", 0) - 1
        reason = str(row[reason_i] if 0 <= reason_i < len(row) else "").strip().casefold()
        email = str(row[email_i] if 0 <= email_i < len(row) else "").strip()
        if reason == "internal duplicate":
            key = internal_duplicate_identity(key, email)
        if key in rows:
            continue
        rows[key] = r
        name = row[name_i] if name_i < len(row) else None
        exp = row[exp_i] if exp_i < len(row) else None
        known[key] = (str(name or ""),
                      state_hash(str(exp or ""), [row[i] if i < len(row) else None for i in der_i]))
    if profile_ids:
        with PROFILE_ID_LOCK:
            PROFILE_IDS.update(profile_ids)
    return rows, known, orders


def find_rows_by_url(ws, columns, target_keys):
    """Find only the requested links without loading every large Experience cell."""
    url_col = columns["linkedin url"]
    found = {}
    targets = set(target_keys)
    if not targets:
        return found
    for row_num, (value,) in enumerate(
            ws.iter_rows(min_row=2, min_col=url_col, max_col=url_col, values_only=True), start=2):
        key = url_key(value)
        if key in targets and key not in found:
            found[key] = row_num
            if len(found) == len(targets):
                break
    return found


def scan_all_sheets(wb):
    known_all = {}
    for ws in wb.worksheets:
        if _is_rpf_sheet(ws.title):
            continue
        columns = prepare_worksheet(ws)
        _rows, known, _orders = scan_workbook(ws, columns)
        for key, value in known.items():
            known_all[(ws.title.casefold(), key)] = value
    return known_all


def file_stamp():
    try:
        st = WORKBOOK.stat()
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def ensure_workbook_exists():
    """Create the workbook on startup so the output file is easy to find."""
    if WORKBOOK.exists():
        return
    with EXCEL_LOCK:
        if WORKBOOK.exists():
            return
        wb, _ws, _columns = open_or_create_workbook()
        try:
            atomic_save_workbook(wb)
        finally:
            wb.close()


def migrate_legacy_sqlite_if_present():
    """One-time best-effort import of URLs missing from Excel from the old v1.3-v1.5 store."""
    if not LEGACY_DATABASE.exists():
        return
    try:
        conn = sqlite3.connect(LEGACY_DATABASE, timeout=3)
        conn.row_factory = sqlite3.Row
        try:
            tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "profiles" not in tables:
                return
            legacy_rows = conn.execute("SELECT name, url, experience FROM profiles ORDER BY rowid").fetchall()
        finally:
            conn.close()
        if not legacy_rows:
            return
        imported = 0
        with EXCEL_LOCK:
            wb, ws, columns = open_or_create_workbook()
            try:
                rows, _known, _orders = scan_workbook(ws, columns)
                for record in legacy_rows:
                    url = normalize_url(record["url"] or "")
                    key = url.lower()
                    exp = sanitize(str(record["experience"] or "").strip())[:MAX_EXPERIENCE_CHARS]
                    if not url or not exp or key in rows:
                        continue
                    row = max(ws.max_row + 1, 2)
                    set_text(ws.cell(row, columns["name"]), sanitize(str(record["name"] or "")))
                    set_text(ws.cell(row, columns["linkedin url"]), url)
                    set_text(ws.cell(row, columns["experience section"]), exp)
                    ws.cell(row, columns["experience section"]).alignment = WRAP
                    rows[key] = row
                    imported += 1
                if imported:
                    atomic_save_workbook(wb)
            finally:
                wb.close()
        if imported:
            print(f"Imported {imported} profiles that were missing from profiles.xlsx from the previous SQLite store.")
    except Exception as exc:
        print(f"Legacy SQLite import skipped; existing Excel data was left intact. Details: {exc}")


# --------------------------------------------------------------------------- in-memory workbook cache
def drop_cache():
    """Forget the in-memory workbook. Call with EXCEL_LOCK held (or at startup)."""
    global CACHE_WB, CACHE_STAMP, CACHE_SHEETS
    if CACHE_WB is not None:
        try:
            CACHE_WB.close()
        except Exception:
            pass
    CACHE_WB, CACHE_STAMP, CACHE_SHEETS = None, None, {}


def _index_sheet(ws):
    columns = prepare_worksheet(ws)
    rows, known, _orders = scan_workbook(ws, columns)
    return {"ws": ws, "cols": columns, "rows": rows, "next": max(ws.max_row + 1, 2)}, known


def load_cache():
    """Read the workbook ONCE, index every sheet (url -> row, next free row) and return the KNOWN index.
    EXCEL_LOCK must be held."""
    global CACHE_WB, CACHE_STAMP, CACHE_SHEETS
    drop_cache()
    stamp = file_stamp()
    wb, _ws, _columns = open_or_create_workbook()
    sheets, known_all = {}, {}
    for ws in wb.worksheets:
        if _is_rpf_sheet(ws.title):
            continue
        info, known = _index_sheet(ws)
        sheets[ws.title.casefold()] = info
        for key, value in known.items():
            known_all[(ws.title.casefold(), key)] = value
    CACHE_WB, CACHE_SHEETS, CACHE_STAMP = wb, sheets, stamp
    return known_all


def cached_sheet(requested):
    """Find or create a worksheet in the cached workbook. EXCEL_LOCK must be held."""
    sk = requested.casefold()
    info = CACHE_SHEETS.get(sk)
    if info is None:
        ws = CACHE_WB.create_sheet(requested)
        ws.append(HEADERS)
        info, _known = _index_sheet(ws)
        CACHE_SHEETS[sk] = info
    return info


def workbook_writable():
    """Cheap lock test (on Windows Excel holds the file). Avoids a heavy save that is bound to fail."""
    if not WORKBOOK.exists():
        return True
    try:
        with open(WORKBOOK, "r+b"):
            return True
    except PermissionError:
        return False
    except OSError:
        return True


def _replace_known(known):
    with STATE_LOCK:
        fresh = dict(known)
        for identity in PENDING:
            if identity in KNOWN:
                fresh[identity] = KNOWN[identity]
        KNOWN.clear()
        KNOWN.update(fresh)


# --------------------------------------------------------------------------- journal + staging
def _journal_append(entry):
    with open(JOURNAL, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        f.flush()


def _journal_rewrite():
    """Compact the journal down to whatever is still pending. Call with STATE_LOCK held."""
    if not PENDING:
        JOURNAL.unlink(missing_ok=True)
        return
    temp = APP_DIR / f"profiles_journal_{uuid.uuid4().hex}.tmp"
    with open(temp, "w", encoding="utf-8") as f:
        for entry in PENDING.values():
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    os.replace(temp, JOURNAL)


def stage_capture(name, url, experience, journal=True, opened_at=None, details=None, metadata=None, sheet_name="Profiles", clear_empty_metadata=False, identity_key=None):
    """Stage a record under its worksheet + LinkedIn URL identity."""
    sheet_name = normalize_sheet_name(sheet_name)
    key = url_key(url)
    row_key = str(identity_key or key)
    identity = (sheet_name.casefold(), row_key)
    experience = experience or ""
    try:
        opened_at = int(float(opened_at)) if opened_at else None
    except (TypeError, ValueError):
        opened_at = None
    if experience and details is None:
        details = analyze_first_experience(experience)
    derived = ([details.get("first_exp", ""), details.get("company", ""), details.get("job_title", ""),
                details.get("experience_geo", ""), details.get("overall", ""), details.get("current", ""),
                details.get("remark", "")] if details else [])
    # Successful captures must be able to clear stale QA reasons/comments from an earlier failed attempt.
    # Status-only updates omit empty values so they do not erase previously captured profile data.
    metadata = {k: v for k, v in (metadata or {}).items() if v is not None and (clear_empty_metadata or str(v).strip())}
    with STATE_LOCK:
        old = KNOWN.get(identity)
        is_new = old is None
        eff_name = name or (old[0] if old else "")
        base_hash = state_hash(experience, derived) if experience else ((old[1].split("|meta:", 1)[0]) if old else "")
        metadata_hash = digest(json.dumps(metadata, ensure_ascii=False, sort_keys=True))
        exp_hash = base_hash + "|meta:" + metadata_hash
        if old is not None and old[0] == eff_name and old[1] == exp_hash:
            return False, False, len(KNOWN)
        entry = dict(PENDING.get(identity) or {"name": "", "url": url, "exp": None, "opened_at": None})
        entry["url"] = url
        entry["sheet_name"] = sheet_name
        if identity_key:
            entry["identity_key"] = row_key
        entry["clear_empty_metadata"] = bool(entry.get("clear_empty_metadata") or clear_empty_metadata)
        if name:
            entry["name"] = name
        if experience:
            entry["exp"] = experience
            entry["first_exp"], entry["company"], entry["job_title"], entry["experience_geo"], entry["overall"], entry["current"], entry["remark"] = derived
        entry.setdefault("metadata", {}).update(metadata)
        entry["completed_at"] = int(time.time() * 1000)
        prev = entry.get("opened_at")
        if opened_at and (not prev or opened_at < prev):
            entry["opened_at"] = opened_at
        if not entry.get("opened_at"):
            entry["opened_at"] = int(time.time() * 1000)
        PENDING[identity] = entry
        KNOWN[identity] = (eff_name, exp_hash)
        if journal:
            _journal_append(entry)
        count = len(KNOWN)
    WAKE.set()
    return is_new, True, count

def replay_journal():
    if not JOURNAL.exists():
        return
    replayed = 0
    with open(JOURNAL, "r", encoding="utf-8") as f:
        for line in f:
            try:
                rec = json.loads(line)
                stage_capture(rec.get("name") or "", rec["url"], rec.get("exp") or "", journal=False,
                              opened_at=rec.get("opened_at"),
                              details=({"first_exp": rec.get("first_exp", ""), "company": rec.get("company", ""),
                                        "job_title": rec.get("job_title", ""), "experience_geo": rec.get("experience_geo") or analyze_first_experience(rec.get("exp") or "").get("experience_geo", ""),
                                        "overall": rec.get("overall", ""),
                                        "current": rec.get("current", ""), "remark": rec.get("remark", "")}
                                       if "first_exp" in rec else None),
                              metadata=rec.get("metadata", {}),
                              sheet_name=rec.get("sheet_name") or "Profiles",
                              clear_empty_metadata=bool(rec.get("clear_empty_metadata")),
                              identity_key=rec.get("identity_key"))
                replayed += 1
            except Exception:
                continue  # a half-written last line after a crash is ignored
    with STATE_LOCK:
        _journal_rewrite()
    if replayed:
        print(f"Recovered {len(PENDING)} unsaved capture(s) from the journal.")


def load_known_from_workbook():
    global LAST_EXCEL_ERROR, WB_STAMP
    try:
        with EXCEL_LOCK:
            known = load_cache()          # also primes the write cache: the first flush is then fast
            WB_STAMP = CACHE_STAMP
        _replace_known(known)
    except Exception as exc:
        LAST_EXCEL_ERROR = str(exc)
        print(f"Warning: could not read {WORKBOOK.name} at startup: {exc}")


def refresh_known_if_changed():
    """Re-index every sheet only when the workbook file changed on disk (one read, shared with the writer).
    Called from the writer thread, NEVER from a capture request."""
    global WB_STAMP
    stamp = file_stamp()
    if stamp == WB_STAMP:
        return
    if not EXCEL_LOCK.acquire(timeout=0.05):
        return
    try:
        stamp = file_stamp()
        if stamp == WB_STAMP:
            return
        try:
            known = load_cache()
        except Exception:
            return
        _replace_known(known)
        WB_STAMP = CACHE_STAMP
    finally:
        EXCEL_LOCK.release()


def write_record(ws, columns, row, rec, new):
    id_cell = ws.cell(row, columns["id"])
    if not str(id_cell.value or "").strip():
        set_text(id_cell, new_profile_id())
    if new:
        set_text(ws.cell(row, columns["linkedin url"]), rec["url"])
        set_text(ws.cell(row, columns["name"]), rec.get("name") or "")
    elif rec.get("name"):
        set_text(ws.cell(row, columns["name"]), rec["name"])
    if rec.get("exp") is not None:
        cell = ws.cell(row, columns["experience section"])
        set_text(cell, rec["exp"])
        cell.alignment = WRAP
        if "first_exp" in rec:                      # derived from the FIRST experience
            for key, field in (("first experience", "first_exp"), ("company", "company"),
                               ("job title", "job_title"), ("experience's geo", "experience_geo"),
                               ("tenure overall", "overall"),
                               ("tenure current", "current"), ("remark", "remark")):
                c = ws.cell(row, columns[key])
                set_text(c, rec.get(field) or "")
                c.alignment = WRAP
    for key, value in (rec.get("metadata") or {}).items():
        normalized_key = normalized_header(key)
        if normalized_key == "geo":
            normalized_key = normalized_header(HEADER_TITLES["geo"])
        col = columns.get(normalized_key)
        if col:
            cell = ws.cell(row, col)
            set_text(cell, "" if value is None else str(value))
            if normalized_header(key) in {"qa comments", "keyword found"}:
                cell.alignment = WRAP


def allocate_row(ws, rows, opened_at):
    """Append newly completed profiles; no visible serial number is used."""
    return max(ws.max_row + 1, 2)


# --------------------------------------------------------------------------- flushing to Excel
def flush_once():
    """Write pending captures to their requested sheets in one workbook save."""
    global LAST_EXCEL_ERROR, WB_STAMP, CACHE_STAMP, REPORT_MIGRATION_PENDING, LAST_FLUSH_SECONDS
    with FLUSH_LOCK:
        if REPORT_MIGRATION_PENDING:
            try:
                migrate_report_workbook()
                REPORT_MIGRATION_PENDING = False
            except Exception as exc:
                LAST_EXCEL_ERROR = str(exc)
        with STATE_LOCK:
            batch = {identity: dict(value) for identity, value in PENDING.items()}
        if not batch:
            return True
        started = time.monotonic()
        try:
            with EXCEL_LOCK:
                if not workbook_writable():
                    raise PermissionError("profiles.xlsx is open in another program (Excel?). "
                                          "Close it so pending captures can be saved.")
                if CACHE_WB is None or file_stamp() != CACHE_STAMP:
                    _replace_known(load_cache())      # file changed outside: re-read once
                rpf_context = _rpf_context(CACHE_WB)
                grouped = {}
                for identity, rec in batch.items():
                    sheet_name = normalize_sheet_name(rec.get("sheet_name") or "Profiles")
                    grouped.setdefault(sheet_name, {})[identity[1]] = rec
                for requested_sheet, sheet_batch in grouped.items():
                    info = cached_sheet(requested_sheet)
                    ws, columns, rows = info["ws"], info["cols"], info["rows"]
                    order_col = columns.get("open order")
                    for key, rec in sheet_batch.items():
                        row = rows.get(key)                       # O(1): no scan of the url column
                        if row is not None:
                            write_record(ws, columns, row, rec, new=False)
                            _upsert_rpf_row(ws, columns, row, rpf_context)
                            continue
                        row = info["next"]                        # O(1): no ws.max_row scan
                        info["next"] += 1
                        rows[key] = row
                        write_record(ws, columns, row, rec, new=True)
                        if order_col:
                            ws.cell(row, order_col).value = rec.get("completed_at") or int(time.time() * 1000)
                        _upsert_rpf_row(ws, columns, row, rpf_context)
                atomic_save_workbook(CACHE_WB)
                CACHE_STAMP = WB_STAMP = file_stamp()
        except Exception as exc:
            LAST_EXCEL_ERROR = str(exc)
            print(f"[ExcelWriter] save failed; {len(batch)} capture(s) remain queued: {exc}", flush=True)
            return False

        LAST_FLUSH_SECONDS = time.monotonic() - started
        print(f"[ExcelWriter] saved {len(batch)} capture(s) in {LAST_FLUSH_SECONDS:.2f}s", flush=True)
        with STATE_LOCK:
            for identity, rec in batch.items():
                if PENDING.get(identity) == rec:
                    del PENDING[identity]
            _journal_rewrite()
        LAST_EXCEL_ERROR = ""
        return True


def writer_loop():
    next_ok = 0.0
    while True:
        WAKE.wait(timeout=POLL_SECONDS)
        try:
            refresh_known_if_changed()        # cheap os.stat; re-reads only if the file really changed
        except Exception:
            pass
        if not PENDING:
            WAKE.clear()
            continue
        # Coalesce captures, and space saves out in proportion to how long the last save took.
        time.sleep(max(BATCH_SECONDS, next_ok - time.monotonic()))
        WAKE.clear()  # captures arriving during the flush will set it again
        if not flush_once():
            time.sleep(RETRY_SECONDS)
            WAKE.set()
            continue
        next_ok = time.monotonic() + max(BATCH_SECONDS, min(MAX_GAP_SECONDS, ADAPTIVE_FACTOR * LAST_FLUSH_SECONDS))


def _flush_at_exit():
    try:
        flush_once()
    except Exception:
        pass


# --------------------------------------------------------------------------- routes
@app.get("/health")
def health():
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    return jsonify(ok=True, service="LinkedIn Experience Copier", storage="write-behind Excel")


@app.post("/configure-workbook")
def configure_workbook():
    """Select a safe output workbook filename before a capture queue starts."""
    global WORKBOOK, WB_STAMP, REPORT_MIGRATION_PENDING, LAST_EXCEL_ERROR
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    data = request.get_json(silent=True) or {}
    filename = Path(str(data.get("filename") or "profiles.xlsx")).name.strip()
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    if not filename or filename in {".xlsx", "..xlsx"} or re.search(r'[<>:"/\\|?*\x00-\x1f]', filename):
        return jsonify(ok=False, error="Invalid Excel filename."), 400
    with EXCEL_LOCK:
        WORKBOOK = APP_DIR / filename
        WB_STAMP = None
        drop_cache()
    ensure_workbook_exists()
    try:
        migrate_report_workbook()
        REPORT_MIGRATION_PENDING = False
    except Exception as exc:
        REPORT_MIGRATION_PENDING = True
        LAST_EXCEL_ERROR = str(exc)
    refresh_known_if_changed()
    return jsonify(ok=True, file=WORKBOOK.name)


@app.get("/sheets")
def list_workbook_sheets():
    """List worksheets available in a safe workbook filename without changing the active workbook."""
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    filename = Path(str(request.args.get("filename") or "profiles.xlsx")).name.strip()
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    if not filename or filename in {".xlsx", "..xlsx"} or re.search(r'[<>:"/\\|?*\x00-\x1f]', filename):
        return jsonify(ok=False, error="Invalid Excel filename."), 400
    path = APP_DIR / filename
    if not path.exists():
        return jsonify(ok=True, file=filename, sheets=["Profiles"])
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            sheets = [name for name in wb.sheetnames if not _is_rpf_sheet(name)]
        finally:
            wb.close()
        if not any(name.casefold() == "profiles" for name in sheets):
            sheets.insert(0, "Profiles")
        return jsonify(ok=True, file=filename, sheets=sheets)
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not read workbook worksheets: {exc}"), 500


@app.post("/create-sheet")
def create_batch_sheet():
    """Create a fresh named worksheet for one bulk batch, without overwriting existing rows."""
    global WB_STAMP
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    data = request.get_json(silent=True) or {}
    try:
        sheet_name = normalize_sheet_name(data.get("sheet_name") or data.get("name"), allow_profiles=False)
        ensure_workbook_exists()
        with EXCEL_LOCK:
            wb = load_workbook(WORKBOOK)
            try:
                actual = next((name for name in wb.sheetnames if name.casefold() == sheet_name.casefold()), None)
                if actual:
                    ws = wb[actual]
                    if ws.max_row > 1:
                        return jsonify(ok=False, error=f"Sheet '{actual}' already contains data. Choose a new batch sheet name."), 409
                    columns = prepare_worksheet(ws)
                    sheet_name = actual
                    created = False
                else:
                    ws = wb.create_sheet(sheet_name)
                    ws.append(HEADERS)
                    columns = prepare_worksheet(ws)
                    created = True
                atomic_save_workbook(wb)
                WB_STAMP = file_stamp()
            finally:
                wb.close()
        return jsonify(ok=True, sheet=sheet_name, created=created, file=WORKBOOK.name)
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not create batch sheet: {exc}"), 500


@app.get("/status")
def status():
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    identity = None
    raw_url = str(request.args.get("url") or "").strip()
    if raw_url:
        url = normalize_url(raw_url)
        if not url:
            return jsonify(ok=False, error="Invalid profile URL."), 400
        try:
            sheet_name = normalize_sheet_name(request.args.get("sheet_name") or "Profiles")
        except ValueError as exc:
            return jsonify(ok=False, error=str(exc)), 400
        identity = (sheet_name.casefold(), url_key(url))
    with STATE_LOCK:
        count, pending = len(KNOWN), len(PENDING)
        profile_pending = identity in PENDING if identity is not None else None
    return jsonify(ok=True, count=count, file=WORKBOOK.name, excel_pending=pending > 0,
                   profile_pending=profile_pending, pending=pending, excel_error=LAST_EXCEL_ERROR,
                   last_flush_seconds=LAST_FLUSH_SECONDS)



def _first_experience_keyword_text(experience, analysis):
    """Return only the first Experience item, including its description when present."""
    lines = _entry_lines(experience or "")
    date_rows = [i for i, line in enumerate(lines) if DATE_LINE_RE.match(line)]
    if len(date_rows) >= 2:
        # The next role's title/company appear before its date row. Exclude that
        # header too, so keywords from the next role cannot match the first role.
        next_date = date_rows[1]
        cutoff = max(date_rows[0] + 1, next_date - 2)
        if next_date >= 2 and _is_location_or_mode(lines[next_date - 1]):
            cutoff = max(date_rows[0] + 1, next_date - 3)
        return "\n".join(lines[:cutoff])
    if date_rows:
        return "\n".join(lines)
    # If dates are absent/unparseable, use bounded parser output, never all roles.
    return str((analysis or {}).get("first_exp") or "")


@lru_cache(maxsize=2048)
def _keyword_pattern(word):
    return re.compile(r"(?<!\w)" + re.escape(word) + r"(?!\w)", re.I)


def _keyword_matches(name, first_experience, metadata, keywords):
    # Keep whole-profile keyword screening for the existing fields; only Experience is scoped
    # to its first entry by _first_experience_keyword_text().
    fields = {
        "Name": name or "", "Geo": metadata.get("Geo", ""),
        "Headline": metadata.get("Headline", ""), "About": metadata.get("About", ""),
        "Experience": first_experience or "", "Activity": metadata.get("Activity", "")
    }
    found = []
    seen_keywords = set()
    for raw in (*FIXED_SCREENING_KEYWORDS, *(keywords or [])):
        word = str(raw or "").strip()
        key = word.casefold()
        if not word or key in seen_keywords:
            continue
        seen_keywords.add(key)
        pat = _keyword_pattern(word)
        for section, value in fields.items():
            if value and pat.search(str(value)):
                item = f"{word} ({section})"
                if item not in found:
                    found.append(item)
    return found


def _qa_rule_outcome(analysis, open_status):
    remark = str((analysis or {}).get("remark") or "").strip()
    open_status = str(open_status or "").strip().casefold()
    if remark == REMARK_DEAD:
        return "Disqualified", "Dead Contact", ""
    if open_status in {"yes", "true", "open to work", "open-to-work"}:
        return "Disqualified", "Open To Work", ""
    if remark == REMARK_MISSING:
        return "Disqualified", MISSING_TENURE_LABEL, ""
    if remark in {REMARK_PDC2, REMARK_PDC3}:
        current = str((analysis or {}).get("current") or "Not Available")
        return "Disqualified", remark, f"Current Tenure: {current}"
    if remark == REMARK_PDC1:
        return "Qualified PDC 1", "PDC 1", "No Fatal Info Found"
    return None


def _normalized_job_title(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()


def _qa_fields(name, experience, metadata, keywords, analysis, capture_status, company_validation=None,
               company_page_evidence=None):
    first_experience = _first_experience_keyword_text(experience, analysis)
    matches = _keyword_matches(name, first_experience, metadata, keywords)
    evidence = company_page_evidence if isinstance(company_page_evidence, dict) else {}
    evidence_company = _normalized_company(evidence.get("company", ""))
    first_company = _normalized_company((analysis or {}).get("company", ""))
    unclaimed_company_detail = ""
    if evidence.get("status") == "unclaimed" and evidence_company and evidence_company == first_company:
        company_label = str((analysis or {}).get("company", "First Experience company"))
        matches.append(f"Unclaimed company page ({company_label})")
        unclaimed_company_detail = f"Unclaimed company page detected for first Experience company: {company_label}."
    lines = _entry_lines(experience or "")
    current_company_candidates = set()
    for i, line in enumerate(lines):
        match = DATE_LINE_RE.match(line)
        if not match or match.group("end").lower() != "present":
            continue
        company = _company_for_date(lines, i)
        normalized = _normalized_company(company)
        if normalized:
            current_company_candidates.add(normalized)
    dual = len(current_company_candidates) >= 2
    reasons = []
    if matches:
        reasons.append(("Keyword Match", "; ".join(matches)))
    if dual:
        reasons.append(("Dual Employment", "Two or more distinct companies appear to have roles marked Present"))

    decision = _qa_rule_outcome(analysis, metadata.get("Open to Work"))
    status, primary, comments = "Qualified", "", ""
    if decision:
        status, primary, comments = decision
        if unclaimed_company_detail:
            comments = f"{comments}; {unclaimed_company_detail}" if comments else unclaimed_company_detail
    elif reasons:
        status = "Review Required"
        primary = reasons[0][0]
        comments = "; ".join(f"{reason}: {detail}" for reason, detail in reasons)
    validation = company_validation or {"target_company": "", "check": "", "entries": ""}
    target_company = validation.get("target_company", "")
    company_check = validation.get("check", "")
    matching_entries = validation.get("entries", "")
    if company_check in {"Dead Contact", "Company Not Found", "Needs Review"}:
        comments = "; ".join(part for part in comments.split("; ") if part and part != "No Fatal Info Found")
        if company_check == "Dead Contact":
            status, primary = "Disqualified", "Dead Contact"
            detail = (f"Target company {target_company} appears only in past {matching_entries}; "
                      "the prospect no longer works there.")
        elif company_check == "Company Not Found":
            status, primary = "Review Required", "Company Not Found"
            detail = f"Target company {target_company} was not found in the captured Experience history."
        else:
            status, primary = "Review Required", "Needs Review"
            detail = f"Could not reliably determine employment at target company {target_company}."
        comments = f"{comments}; {detail}" if comments else detail
    elif company_check == "Dual Employment":
        detail = f"Target company Dual Employment: {matching_entries}."
        comments = f"{comments}; {detail}" if comments else detail

    metadata.update({
        "QA Status": status,
        "Primary Reason": primary,
        "Secondary Reason": "",
        "QA Comments": comments,
        "Keyword Found": "; ".join(matches),
        "Dual Employment": "Yes" if dual else "No",
        "Capture Status": capture_status,
        "Target Company": target_company,
        "Company Check": company_check,
        "Matching Experience Entry": matching_entries,
    })
    return metadata


@app.post("/record-status")
def record_status():
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    data = request.get_json(silent=True) or {}
    url = normalize_url(str(data.get("url") or ""))
    status = str(data.get("capture_status") or "Capture Failed")
    email = sanitize(str(data.get("email") or ""))[:320]
    internal_duplicate = bool(data.get("internal_duplicate")) or status.casefold() == "internal duplicate"
    if internal_duplicate:
        status = "Internal Duplicate"
        qa_status, primary = "Disqualified", "Internal Duplicate"
        default_comment = "Duplicate LinkedIn URL with a different email."
        identity_key = internal_duplicate_identity(url, email)
    elif status == "Invalid Link":
        qa_status, primary = "Disqualified", "Invalid JT Link"
        default_comment = "LinkedIn redirected this URL to /404/."
        identity_key = None
    else:
        qa_status = status if status == "Unable to Capture" else "Review Required"
        primary = status
        default_comment = status
        identity_key = None
    try:
        sheet_name = normalize_sheet_name(data.get("sheet_name") or data.get("sheet") or "Profiles")
    except ValueError:
        sheet_name = "Profiles"
    report_job_title = sanitize(str(data.get("report_job_title") or ""))[:320]
    target_company = sanitize(str(data.get("target_company") or "")).strip()[:320]
    metadata = {
        "Email": email,
        "Target Company": target_company,
        "Company Check": "",
        "Matching Experience Entry": "",
        "QA Status": qa_status,
        "Primary Reason": primary,
        "QA Comments": str(data.get("comment") or default_comment), "Report's Job Title": report_job_title,
        "JT Match": "No" if report_job_title else "", "Capture Status": status,
        "Keyword Found": "", "Dual Employment": "Not Evaluated",
        "Geo": "", "Open to Work": "Not Available", "Connections": "", "Followers": "", "Activity": ""
    }
    try:
        is_new, changed, count = stage_capture("", url, "", metadata=metadata, sheet_name=sheet_name,
                                               identity_key=identity_key)
        return jsonify(ok=True, changed=changed, count=count, file=WORKBOOK.name, message=status)
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.post("/capture")
def capture():
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    data = request.get_json(silent=True) or {}
    url = normalize_url(data.get("url", ""))
    try:
        sheet_name = normalize_sheet_name(data.get("sheet_name") or data.get("sheet") or "Profiles")
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    if not re.match(r"^https://www\.linkedin\.com/in/[^/]+$", url, re.I):
        return jsonify(ok=False, error="Not a LinkedIn member profile URL."), 400
    html = data.get("page_html", "")
    # Name is helpful metadata; the URL is the identifier and a missing name never blocks saving.
    name = get_name(None, data.get("page_title", "")) or sanitize(str(data.get("page_name", "")).strip())[:200]
    if not name and html:
        name = get_name(make_soup(html), "")
    experience, too_broad = extract_experience(
        data.get("experience_text", ""), data.get("experience_html", ""), html)
    analysis = analyze_first_experience(experience)            # first experience only, from the clean text
    experience, tenure_missing = apply_tenure_flag(experience, analysis)
    remark = analysis["remark"] if experience else ""
    if not experience:
        quality = "empty"
    elif too_broad:
        quality = "too_broad"       # the browser sent more than Experience (we trimmed it)
    elif not analysis["parsed"] or not analysis["first_exp"]:
        quality = "unparsed"        # text present, but the first experience could not be read
    else:
        quality = "ok"

    metadata = dict(data.get("metadata") or {})
    report_job_title = sanitize(str(data.get("report_job_title") or ""))[:320]
    target_company = sanitize(str(data.get("target_company") or "")).strip()[:320]
    company_validation = _target_company_validation(experience, target_company, analysis.get("company", ""))
    derived_job_title = str(analysis.get("job_title") or "")
    metadata.update({
        "Email": sanitize(str(data.get("email") or ""))[:320],
        "Target Company": target_company,
        "Company Check": company_validation["check"],
        "Matching Experience Entry": company_validation["entries"],
        "Report's Job Title": report_job_title,
        "JT Match": ("Yes" if _normalized_job_title(report_job_title) and
                     _normalized_job_title(report_job_title) == _normalized_job_title(derived_job_title) else "No") if report_job_title else "",
        "Geo": metadata.get("Geo", ""), "Open to Work": metadata.get("Open to Work", "Not Available"),
        "Connections": metadata.get("Connections", ""), "Followers": metadata.get("Followers", ""),
        "Activity": metadata.get("Activity", ""),
    })
    metadata = _qa_fields(name, experience, metadata, data.get("keywords") or [], analysis,
                          "Captured" if experience else "Captured with Missing Data", company_validation,
                          data.get("company_page_evidence"))
    internal_duplicate = bool(data.get("internal_duplicate"))
    identity_key = None
    if internal_duplicate:
        email = metadata.get("Email", "")
        metadata.update({"QA Status": "Disqualified", "Primary Reason": "Duplicated Link",
                         "QA Comments": "Duplicate LinkedIn URL with a different email."})
        identity_key = internal_duplicate_identity(url, email)

    try:
        # (the writer thread notices rows you deleted from Excel; a request never reads the workbook)
        is_new, changed, count = stage_capture(name, url, experience, opened_at=data.get("opened_at"), details=analysis,
                                               metadata=metadata, sheet_name=sheet_name,
                                               clear_empty_metadata=bool(experience), identity_key=identity_key)
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not store capture. Details: {exc}", count=0, file=WORKBOOK.name), 500

    if not experience and not changed:
        message = "No Experience text found on this page (nothing new saved)."
    elif not changed:
        message = "Already up to date (no change)."
    elif is_new and not experience:
        message = "Saved profile, but its Experience was empty."
    elif is_new:
        message = "Saved new profile."
    else:
        message = "Updated existing profile."
    label = remark or (MISSING_TENURE_LABEL if tenure_missing else "")
    if label:
        message = message.rstrip(".") + f" - {label}."
    identity = (sheet_name.casefold(), url_key(url))
    with STATE_LOCK:
        pending = len(PENDING)
        profile_pending = identity in PENDING
        excel_error = LAST_EXCEL_ERROR if profile_pending else ""
    if profile_pending and excel_error:
        message += " Excel is busy; it will retry when the workbook is available."
    elif changed:
        # Match the earlier app: acknowledge the journaled capture as Saved immediately;
        # the background writer flushes the workbook without another request from the extension.
        message = "Saved new profile." if is_new else "Updated existing profile."
    return jsonify(ok=True, message=message, count=count, file=WORKBOOK.name, sheet_name=sheet_name,
                   excel_pending=profile_pending, pending=pending, excel_error=excel_error,
                   last_flush_seconds=LAST_FLUSH_SECONDS, quality=quality, remark=remark,
                   tenure_overall=analysis["overall"], tenure_current=analysis["current"],
                   first_experience=analysis["first_exp"], company=analysis["company"], job_title=analysis["job_title"],
                   experience_geo=analysis["experience_geo"],
                   record={"name": name, "url": url, "experience_section": experience,
                           "capture_method": data.get("capture_method", "browser_dom_section_no_scroll")})


@app.post("/export")
def export():
    """Force an immediate write to Excel (use before opening profiles.xlsx)."""
    if not authorized():
        return jsonify(ok=False, error="Unauthorized"), 401
    if not flush_once():
        return jsonify(ok=False, error=f"Excel could not be updated. Close profiles.xlsx if it is open. "
                                       f"Details: {LAST_EXCEL_ERROR}"), 423
    with STATE_LOCK:
        count = len(KNOWN)
    return jsonify(ok=True, file=str(WORKBOOK), count=count, excel_pending=False,
                   message="Excel is up to date.")


def _disable_worksheet_wrapping(ws):
    for row in ws.iter_rows():
        for cell in row:
            align = cell.alignment
            if align.wrap_text:
                cell.alignment = Alignment(horizontal=align.horizontal, vertical=align.vertical,
                                           textRotation=align.textRotation, wrap_text=False,
                                           shrink_to_fit=align.shrink_to_fit, indent=align.indent,
                                           relativeIndent=align.relativeIndent,
                                           justifyLastLine=align.justifyLastLine,
                                           readingOrder=align.readingOrder)


def migrate_report_workbook():
    """Migrate report column order, assign stable IDs, and refresh existing QA fields."""
    global WB_STAMP
    if not WORKBOOK.exists():
        return
    changed = False
    with EXCEL_LOCK:
        wb = load_workbook(WORKBOOK)
        try:
            before_headers = {}
            prepared_columns = {}
            for ws in wb.worksheets:
                if _is_rpf_sheet(ws.title):
                    continue
                before_headers[ws.title] = tuple(ws.cell(1, col).value for col in range(1, ws.max_column + 1))
                prepared_columns[ws.title] = prepare_worksheet(ws)

            # Reserve all existing IDs first so generated migration IDs cannot collide with later rows.
            with PROFILE_ID_LOCK:
                for ws in wb.worksheets:
                    if _is_rpf_sheet(ws.title):
                        continue
                    id_col = prepared_columns[ws.title]["id"]
                    for row in range(2, ws.max_row + 1):
                        existing_id = str(ws.cell(row, id_col).value or "").strip()
                        if existing_id:
                            PROFILE_IDS.add(existing_id)

            seen_profile_ids = set()
            for ws in wb.worksheets:
                if _is_rpf_sheet(ws.title):
                    continue
                columns = prepared_columns[ws.title]
                _disable_worksheet_wrapping(ws)
                after = tuple(ws.cell(1, col).value for col in range(1, ws.max_column + 1))
                changed = changed or before_headers[ws.title] != after
                for row in range(2, ws.max_row + 1):
                    if not any(ws.cell(row, col).value not in (None, "") for col in range(1, ws.max_column + 1)):
                        continue
                    id_cell = ws.cell(row, columns["id"])
                    profile_id = str(id_cell.value or "").strip()
                    if not profile_id or profile_id in seen_profile_ids:
                        profile_id = new_profile_id()
                        set_text(id_cell, profile_id)
                        changed = True
                    seen_profile_ids.add(profile_id)
                    open_cell = ws.cell(row, columns["open to work"])
                    # The previous extractor marked generic intro-card labels as Yes.
                    # Treat legacy Yes values as unknown until a fresh capture verifies the badge.
                    if str(open_cell.value or "").strip().casefold() == "yes":
                        set_text(open_cell, "Not Available")
                        changed = True
                    comments_cell = ws.cell(row, columns["qa comments"])
                    if str(comments_cell.value or "").strip() == "No configured screening flags found.":
                        set_text(comments_cell, "")
                        changed = True
                    exp_value = ws.cell(row, columns["experience section"]).value
                    first_value = ws.cell(row, columns["first experience"]).value
                    if not (first_value or exp_value):
                        continue
                    analysis = analyze_first_experience(str(exp_value or first_value or ""))
                    company_cell = ws.cell(row, columns["company"])
                    title_cell = ws.cell(row, columns["job title"])
                    geo_cell = ws.cell(row, columns["experience's geo"])
                    if analysis.get("experience_geo") and str(geo_cell.value or "").strip() != analysis["experience_geo"]:
                        set_text(geo_cell, analysis["experience_geo"])
                        changed = True
                    if analysis.get("company") and str(company_cell.value or "").strip() != analysis["company"]:
                        set_text(company_cell, analysis["company"])
                        changed = True
                    if analysis.get("job_title") and str(title_cell.value or "").strip() != analysis["job_title"]:
                        set_text(title_cell, analysis["job_title"])
                        changed = True
                    current_companies = set()
                    exp_lines = _entry_lines(str(exp_value or ""))
                    for idx, line in enumerate(exp_lines):
                        match = DATE_LINE_RE.match(line)
                        if match and match.group("end").lower() == "present":
                            normalized = _normalized_company(_company_for_date(exp_lines, idx))
                            if normalized:
                                current_companies.add(normalized)
                    dual_now = len(current_companies) >= 2
                    report_title_cell = ws.cell(row, columns["report's job title"])
                    report_title = str(report_title_cell.value or "").strip()
                    jt_match_cell = ws.cell(row, columns["jt match"])
                    if report_title:
                        jt_match = "Yes" if (_normalized_job_title(report_title) and
                                               _normalized_job_title(report_title) == _normalized_job_title(analysis.get("job_title"))) else "No"
                    else:
                        jt_match = ""
                    if jt_match_cell.value != jt_match:
                        set_text(jt_match_cell, jt_match)
                        changed = True
                    dual_cell = ws.cell(row, columns["dual employment"])
                    dual_value = "Yes" if dual_now else "No"
                    if dual_cell.value != dual_value:
                        set_text(dual_cell, dual_value)
                        changed = True
                    decision = _qa_rule_outcome(analysis, open_cell.value)
                    if decision:
                        status, primary, comments = decision
                        values = (("qa status", status), ("primary reason", primary),
                                  ("qa comments", comments))
                    elif not dual_now and any("dual employment" in str(ws.cell(row, columns[key]).value or "").casefold()
                                                   for key in ("primary reason", "qa comments")):
                        keyword_value = str(ws.cell(row, columns["keyword found"]).value or "").strip()
                        if keyword_value:
                            values = (("qa status", "Review Required"), ("primary reason", "Keyword Match"),
                                      ("qa comments", f"Keyword Match: {keyword_value}"))
                        else:
                            values = (("qa status", "Qualified"), ("primary reason", ""),
                                      ("qa comments", ""))
                    else:
                        values = ()
                    for key, value in values:
                        cell = ws.cell(row, columns[key])
                        if cell.value != value:
                            set_text(cell, value)
                            changed = True
            changed = _sync_rpf_sheet(wb, prepared_columns) or changed
            if changed:
                atomic_save_workbook(wb)
                WB_STAMP = None
        finally:
            wb.close()
    if changed:
        print("Updated report schema, assigned profile IDs, and refreshed existing QA fields.")


# --------------------------------------------------------------------------- startup
def startup():
    global LAST_EXCEL_ERROR, REPORT_MIGRATION_PENDING
    try:
        ensure_workbook_exists()
    except Exception as exc:
        LAST_EXCEL_ERROR = str(exc)
        print(f"Warning: could not initialize profiles.xlsx: {exc}")
    migrate_legacy_sqlite_if_present()
    try:
        migrate_report_workbook()
        REPORT_MIGRATION_PENDING = False
    except Exception as exc:
        REPORT_MIGRATION_PENDING = True
        print(f"Report layout migration will retry on the next workbook write: {exc}")
    load_known_from_workbook()
    replay_journal()
    atexit.register(_flush_at_exit)
    Thread(target=writer_loop, daemon=True).start()
    if PENDING:
        WAKE.set()


if __name__ == "__main__":
    startup()
    print("LinkedIn Experience Copier service (fast write-behind mode)")
    print(f"Excel output: {WORKBOOK}")
    print("Captures are acknowledged instantly (journaled), then written to Excel in batches.")
    print(f"Listening only on http://{HOST}:{PORT}")
    try:
        from waitress import serve
    except ImportError:
        print("Tip: 'pip install waitress' for a faster server than Flask's built-in one.")
        app.run(host=HOST, port=PORT, debug=False, threaded=True)
    else:
        serve(app, host=HOST, port=PORT, threads=4)
