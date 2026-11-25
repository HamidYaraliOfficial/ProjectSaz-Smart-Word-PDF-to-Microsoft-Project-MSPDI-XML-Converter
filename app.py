# -*- coding: utf-8 -*-
"""
🌐 تبدیل‌گر وب Word/PDF فارسی به Microsoft Project (MSPDI XML)
نسخه ۲.۰ — با خلاصه‌سازی هوشمند، استخراج جدول/عدد/فاز، و ساختار WBS

اجرا:
    pip install Flask python-docx pdfplumber PyPDF2
    python app.py
"""

import os
import re
import html
import uuid
import threading
import traceback
import webbrowser
from datetime import datetime, timedelta
from collections import OrderedDict

from flask import (
    Flask, render_template_string, request, jsonify, send_file, abort
)

from docx import Document

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False


# ==============================================================
# تنظیمات
# ==============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

JOBS = {}
JOBS_LOCK = threading.Lock()


# ==============================================================
# خواندن فایل‌ها (با حفظ ساختار جدول)
# ==============================================================
def read_docx_structured(path):
    """برمی‌گرداند: لیستی از بلوک‌ها با نوع 'para' یا 'table'"""
    doc = Document(path)
    blocks = []

    # نگاشت پاراگراف‌ها به ترتیب سند
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            style = (p.style.name or "").lower() if p.style else ""
            blocks.append({"type": "para", "text": t, "style": style})

    # جدول‌ها را جداگانه اضافه می‌کنیم (ترتیب دقیق با python-docx پیچیده است)
    for ti, table in enumerate(doc.tables):
        rows = []
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                rows.append(cells)
        if rows:
            blocks.append({"type": "table", "rows": rows, "index": ti})

    return blocks


def read_pdf_structured(path):
    """PDF را به بلوک‌های متنی/جدولی تبدیل می‌کند (جدول‌ها با pdfplumber)."""
    blocks = []
    if HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    # اول جدول‌ها
                    try:
                        tables = page.extract_tables() or []
                    except Exception:
                        tables = []
                    for tbl in tables:
                        rows = [[(c or "").strip() for c in r] for r in tbl if r]
                        rows = [r for r in rows if any(r)]
                        if rows:
                            blocks.append({"type": "table", "rows": rows})
                    # بعد متن
                    txt = page.extract_text() or ""
                    for line in txt.split("\n"):
                        line = line.strip()
                        if line:
                            blocks.append({"type": "para", "text": line, "style": ""})
            if blocks:
                return blocks
        except Exception:
            pass

    if HAS_PYPDF2:
        try:
            with open(path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                text = "\n".join((p.extract_text() or "") for p in reader.pages)
            for line in text.split("\n"):
                line = line.strip()
                if line:
                    blocks.append({"type": "para", "text": line, "style": ""})
        except Exception:
            pass

    return blocks


def read_input_structured(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".docx", ".doc"):
        return read_docx_structured(path)
    if ext == ".pdf":
        return read_pdf_structured(path)
    raise ValueError(f"فرمت پشتیبانی نمی‌شود: {ext}")


# ==============================================================
# استخراج اطلاعات مهم از متن (اعداد، فازها، تاریخ‌ها، مبالغ ...)
# ==============================================================
NUM_FA = "۰۱۲۳۴۵۶۷۸۹"
NUM_AR = "٠١٢٣٤٥٦٧٨٩"

def fa_to_en_digits(s):
    for i, ch in enumerate(NUM_FA):
        s = s.replace(ch, str(i))
    for i, ch in enumerate(NUM_AR):
        s = s.replace(ch, str(i))
    return s


RE_AMOUNT = re.compile(
    r"(?:مبلغ|بودجه|هزینه|اعتبار|قیمت)[^\d۰-۹]{0,20}"
    r"([\d۰-۹٠-٩][\d۰-۹٠-٩,٫٬\.]*)"
    r"\s*(میلیون|میلیارد|هزار)?\s*(ریال|تومان|دلار|یورو)?",
    re.UNICODE,
)
RE_DURATION = re.compile(
    r"(?:مدت|طول|زمان)\s*(?:اجرا|پروژه|طرح)?[^\d۰-۹]{0,15}"
    r"([\d۰-۹٠-٩]+)\s*(ماه|هفته|روز|سال)",
    re.UNICODE,
)
RE_PHASE = re.compile(
    r"(?:فاز|مرحله|گام|فصل)\s*([\d۰-۹٠-٩]+|اول|دوم|سوم|چهارم|پنجم|ششم|هفتم|هشتم|نهم|دهم)"
    r"[\s:：\-–]*(.{0,80})",
    re.UNICODE,
)
RE_DATE = re.compile(
    r"([\d۰-۹٠-٩]{4}[/\-][\d۰-۹٠-٩]{1,2}[/\-][\d۰-۹٠-٩]{1,2})"
)
RE_VERSION = re.compile(r"(نسخه|ویرایش)\s*([\d۰-۹٠-٩\.]+)", re.UNICODE)
RE_PERCENT = re.compile(r"([\d۰-۹٠-٩]+)\s*٪|([\d۰-۹٠-٩]+)\s*درصد", re.UNICODE)
RE_DELIVERABLE = re.compile(
    r"(?:تحویل[‌ ]?دادنی|خروجی|محصول| deliverable|تحویل)[\s:：\-–]*(.{3,120})",
    re.UNICODE | re.IGNORECASE,
)
RE_OBJECTIVE = re.compile(
    r"(?:هدف|اهداف)\s*(?:اصلی|اختصاصی|فرعی|کلی)?[\s:：\-–]*(.{3,200})",
    re.UNICODE,
)
RE_HEADING = re.compile(
    r"^(?:[•●▪◦\-\*]\s*)?"
    r"(?:([\d۰-۹٠-٩]+(?:[\.\-][\d۰-۹٠-٩]+)*)[\s\.\-–:]*)?"
    r"(.{3,})$"
)
RE_BULLET = re.compile(r"^[\s]*(?:[-*•●▪◦–—]|\u2022)\s+(.+)$")
RE_NUMBERED = re.compile(r"^[\s]*([\d۰-۹٠-٩]+(?:[\.\-][\d۰-۹٠-٩]+)*)[\.\)\-–]?\s+(.+)$")
RE_COLON = re.compile(r"^(.{3,60}?)[：:]\s*(.{2,})$")

# کلمات کلیدی که نشان می‌دهند یک خط عنوان/سرتیتر است
HEADING_KEYWORDS = [
    "هدف", "اهداف", "محدوده", "تعریف", "مقدمه", "چکیده", "نتیجه",
    "روش", "روش‌ها", "فاز", "مرحله", "گام", "خروجی", "تحویل",
    "جدول", "پیوست", "منابع", "مراجع", "زمان‌بندی", "بودجه",
    "ریسک", "فرض", "پیش‌فرض", "معیار", "شاخص", "نتایج",
    "مسئولیت", "نقش", "ساختار", "معماری", "الزامات", "پیشنهاد",
    "مشخصات", "خلاصه", "معرفی", "ضرورت", "بیان مسئله",
]


def classify_line(line):
    """نوع خط را تشخیص می‌دهد: heading / numbered / bullet / colon / body"""
    line = line.strip()
    if not line or len(line) < 2:
        return "skip"

    # عنوان‌های شماره‌دار مثل ۱-۲-۱ یا 3.1.4
    m = RE_NUMBERED.match(line)
    if m:
        return "numbered"

    if RE_BULLET.match(line):
        return "bullet"

    # عنوان‌های کوتاه با کلمات کلیدی
    short = line.rstrip(":：")
    if len(short) <= 80 and any(k in short for k in HEADING_KEYWORDS):
        # اگر خیلی طولانی نیست و شبیه جمله نیست
        if short.count(" ") <= 12:
            return "heading"

    m = RE_COLON.match(line)
    if m and len(m.group(1)) < 50:
        return "colon"

    return "body"


def extract_values(text):
    """مقادیر مهم را از متن استخراج می‌کند."""
    t = fa_to_en_digits(text)
    out = {}

    amounts = RE_AMOUNT.findall(t)
    if amounts:
        out["amount"] = []
        for num, unit, cur in amounts[:5]:
            out["amount"].append(f"{num} {unit} {cur}".strip())

    durs = RE_DURATION.findall(t)
    if durs:
        out["duration"] = [f"{n} {u}" for n, u in durs[:5]]

    phases = RE_PHASE.findall(t)
    if phases:
        out["phase"] = []
        for num, desc in phases[:10]:
            desc = desc.strip(" .:،؛")
            out["phase"].append((num, desc[:100]))

    dates = RE_DATE.findall(t)
    if dates:
        out["date"] = list(dict.fromkeys(dates))[:5]

    vers = RE_VERSION.findall(t)
    if vers:
        out["version"] = [f"{k} {v}" for k, v in vers[:3]]

    pcts = RE_PERCENT.findall(t)
    if pcts:
        out["percent"] = [f"{p[0] or p[1]}٪" for p in pcts[:5]]

    deliver = RE_DELIVERABLE.findall(t)
    if deliver:
        out["deliverable"] = [d.strip(" .:،؛")[:150] for d in deliver[:8]]

    objs = RE_OBJECTIVE.findall(t)
    if objs:
        out["objective"] = [o.strip(" .:،؛")[:200] for o in objs[:8]]

    return out


def summarize_table(rows, max_rows=12, max_cols=6):
    """جدول را به یک لیست از دیکشنری‌های ردیف تبدیل می‌کند (با سرستون)."""
    if not rows:
        return []
    header = rows[0]
    body = rows[1:] if len(rows) > 1 else []
    result = []
    for r in body[:max_rows]:
        item = OrderedDict()
        for i, cell in enumerate(r[:max_cols]):
            key = header[i] if i < len(header) and header[i] else f"ستون {i+1}"
            item[key] = cell
        if any(item.values()):
            result.append(item)
    return result


# ==============================================================
# حالت ۱: تجزیه کلاسیک (مثل نسخه قبل — فقط خطوط شماره‌دار/بولت)
# ==============================================================
_P_NUM = re.compile(r"^[\s]*[0-9۰-۹]+[\s]*[-.)]\s*(.+)$")
_P_BUL = re.compile(r"^[\s]*[-*•●▪◦]\s*(.+)$")
_P_COL = re.compile(r"^(.+?)[：:]\s*(.+)$")


def parse_tasks_classic(blocks, strict=True):
    tasks = []
    for blk in blocks:
        if blk["type"] == "table":
            continue
        for raw in blk["text"].split("\n"):
            line = raw.strip()
            if not line or len(line) < 3:
                continue
            m = _P_NUM.match(line)
            if m:
                tasks.append(m.group(1).strip())
                continue
            m = _P_BUL.match(line)
            if m:
                tasks.append(m.group(1).strip())
                continue
            m = _P_COL.match(line)
            if m:
                tasks.append(f"{m.group(1).strip()} - {m.group(2).strip()}")
                continue
            if not strict and len(line) < 120:
                tasks.append(line)
    return tasks


# ==============================================================
# حالت ۲: تجزیه ساختاریافته (WBS واقعی + استخراج مقادیر)
# ==============================================================
def parse_tasks_structured(blocks):
    """
    خروجی: لیستی از تسک‌ها با ساختار:
    {
      "name": str,
      "level": 0/1/2,
      "kind": "phase|objective|task|deliverable|table|value|note",
      "value": optional str,
      "source": "text|table"
    }
    """
    tasks = []

    # —— ۱) پردازش بلوک‌های متنی ——
    text_lines = []
    for blk in blocks:
        if blk["type"] == "para":
            text_lines.extend(
                [ln.strip() for ln in blk["text"].split("\n") if ln.strip()]
            )

    # استخراج مقادیر سراسری
    full_text = "\n".join(text_lines)
    values = extract_values(full_text)

    # ۱-الف) فازها/مراحل (اگر در متن آمده)
    if "phase" in values:
        tasks.append({
            "name": "فازها و مراحل پروژه",
            "level": 0,
            "kind": "phase_group",
        })
        for num, desc in values["phase"]:
            name = f"فاز {num}"
            if desc:
                name += f" — {desc}"
            tasks.append({
                "name": name[:150],
                "level": 1,
                "kind": "phase",
            })

    # ۱-ب) اهداف
    if "objective" in values:
        tasks.append({
            "name": "اهداف پروژه",
            "level": 0,
            "kind": "objective_group",
        })
        for o in values["objective"]:
            tasks.append({
                "name": o[:180],
                "level": 1,
                "kind": "objective",
            })

    # ۱-ج) تحویل‌دادنی‌ها
    if "deliverable" in values:
        tasks.append({
            "name": "تحویل‌دادنی‌ها و خروجی‌ها",
            "level": 0,
            "kind": "deliverable_group",
        })
        for d in values["deliverable"]:
            tasks.append({
                "name": d[:180],
                "level": 1,
                "kind": "deliverable",
            })

    # ۱-د) مقادیر کمی (بودجه، مدت، نسخه، درصد)
    quantitative = []
    if "amount" in values:
        quantitative.append(("بودجه / مبلغ", "، ".join(values["amount"])))
    if "duration" in values:
        quantitative.append(("مدت اجرا", "، ".join(values["duration"])))
    if "version" in values:
        quantitative.append(("نسخه سند", "، ".join(values["version"])))
    if "date" in values:
        quantitative.append(("تاریخ‌ها", "، ".join(values["date"])))
    if "percent" in values:
        quantitative.append(("درصدها", "، ".join(values["percent"])))

    if quantitative:
        tasks.append({
            "name": "مشخصات کمی و کلیدی",
            "level": 0,
            "kind": "quantitative_group",
        })
        for label, val in quantitative:
            tasks.append({
                "name": f"{label}: {val[:180]}",
                "level": 1,
                "kind": "value",
                "value": val,
            })

    # ۱-ه) ساختار سلسله‌مراتبی از متن (عنوان‌ها + بولت‌ها)
    current_l1 = None
    for line in text_lines:
        kind = classify_line(line)
        if kind == "skip":
            continue

        if kind == "heading":
            # عنوان اصلی
            name = line.rstrip(":：").strip()
            tasks.append({
                "name": name[:150],
                "level": 0,
                "kind": "heading",
            })
            current_l1 = name
            continue

        if kind in ("numbered", "bullet"):
            # زیرآیتم
            clean = re.sub(
                r"^[\s]*(?:[-*•●▪◦–—]|\u2022)?\s*"
                r"(?:[\d۰-۹٠-٩]+(?:[\.\-][\d۰-۹٠-٩]+)*[\.\)\-–]?)?\s*",
                "", line
            ).strip()
            if len(clean) < 3:
                continue
            # اگر خیلی طولانی است، به ۱۸۰ کاراکتر محدود کن
            clean = clean[:200]
            tasks.append({
                "name": clean,
                "level": 1 if current_l1 else 0,
                "kind": "task",
            })
            continue

        # جمله‌های بلند فقط در حالت strict حذف می‌شوند
        if kind == "colon":
            m = RE_COLON.match(line)
            if m:
                tasks.append({
                    "name": f"{m.group(1).strip()}: {m.group(2).strip()[:150]}",
                    "level": 1,
                    "kind": "task",
                })

    # —— ۲) پردازش جدول‌ها ——
    for blk in blocks:
        if blk["type"] != "table":
            continue
        rows = blk["rows"]
        if not rows:
            continue
        header = [c for c in rows[0] if c]
        title = " / ".join(header[:3]) if header else "جدول"
        tasks.append({
            "name": f"جدول: {title[:120]}",
            "level": 0,
            "kind": "table_group",
        })
        summarized = summarize_table(rows)
        for item in summarized:
            parts = []
            for k, v in item.items():
                if v:
                    parts.append(f"{k}: {v}")
            if parts:
                tasks.append({
                    "name": " | ".join(parts)[:220],
                    "level": 1,
                    "kind": "table_row",
                })

    return tasks


# ==============================================================
# تبدیل لیست تسک‌ها به MSPDI XML
# ==============================================================
def _iso(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def create_mspdi_xml(task_list, output_path, project_title, start_dt):
    """
    task_list: لیستی از دیکشنری‌ها {name, level, kind}
    اگر kind نداشت، به عنوان task ساده در نظر گرفته می‌شود.
    """
    if not task_list:
        raise ValueError("لیست تسک‌ها خالی است")

    n = len(task_list)
    # تخمین مدت: هر تسک سطح ۱ = ۱ روز، سطح ۰ = ۰.۵ روز
    estimated_days = 0
    for t in task_list:
        lvl = t.get("level", 0) if isinstance(t, dict) else 0
        estimated_days += 1 if lvl == 0 else 0.5
    estimated_days = max(1, estimated_days)

    finish_dt = start_dt + timedelta(days=estimated_days + 1)
    now_iso = _iso(datetime.now())

    L = []
    L.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    L.append('<Project xmlns="http://schemas.microsoft.com/project">')
    L.append('  <SaveVersion>14</SaveVersion>')
    L.append(f'  <Name>{html.escape(project_title)}</Name>')
    L.append(f'  <Title>{html.escape(project_title)}</Title>')
    L.append('  <ScheduleFromStart>1</ScheduleFromStart>')
    L.append(f'  <StartDate>{_iso(start_dt)}</StartDate>')
    L.append(f'  <FinishDate>{_iso(finish_dt)}</FinishDate>')
    L.append('  <FYStartDate>1</FYStartDate>')
    L.append('  <CurrencyDigits>0</CurrencyDigits>')
    L.append('  <CurrencySymbol>ریال</CurrencySymbol>')
    L.append('  <CurrencyCode>IRR</CurrencyCode>')
    L.append('  <CalendarUID>1</CalendarUID>')
    L.append('  <DefaultStartTime>08:00:00</DefaultStartTime>')
    L.append('  <DefaultFinishTime>17:00:00</DefaultFinishTime>')
    L.append('  <MinutesPerDay>480</MinutesPerDay>')
    L.append('  <MinutesPerWeek>2400</MinutesPerWeek>')
    L.append('  <DaysPerMonth>20</DaysPerMonth>')
    L.append('  <DefaultTaskType>0</DefaultTaskType>')
    L.append('  <DurationFormat>7</DurationFormat>')
    L.append('  <WorkFormat>2</WorkFormat>')
    L.append('  <WeekStartDay>6</WeekStartDay>')
    L.append(f'  <CurrentDate>{now_iso}</CurrentDate>')
    L.append('  <NewTasksEstimated>1</NewTasksEstimated>')
    L.append('  <Tasks>')

    # تسک ریشه
    L.append('    <Task>')
    L.append('      <UID>0</UID>')
    L.append('      <ID>0</ID>')
    L.append(f'      <Name>{html.escape(project_title)}</Name>')
    L.append('      <Type>1</Type>')
    L.append('      <IsNull>0</IsNull>')
    L.append(f'      <CreateDate>{now_iso}</CreateDate>')
    L.append('      <WBS>0</WBS>')
    L.append('      <OutlineNumber>0</OutlineNumber>')
    L.append('      <OutlineLevel>0</OutlineLevel>')
    L.append('      <Priority>500</Priority>')
    L.append(f'      <Start>{_iso(start_dt)}</Start>')
    L.append(f'      <Finish>{_iso(finish_dt)}</Finish>')
    L.append('      <Duration>PT8H0M0S</Duration>')
    L.append('      <DurationFormat>7</DurationFormat>')
    L.append('      <Summary>1</Summary>')
    L.append('      <Milestone>0</Milestone>')
    L.append('      <Critical>0</Critical>')
    L.append('      <PercentComplete>0</PercentComplete>')
    L.append('      <CalendarUID>-1</CalendarUID>')
    L.append('    </Task>')

    # شمارنده WBS بر اساس سطح
    wbs_counters = {0: 0, 1: 0, 2: 0}
    current_uid = 0
    cursor_day = 0.0

    for i, t in enumerate(task_list, 1):
        if isinstance(t, str):
            name, level, kind = t, 0, "task"
        else:
            name = t.get("name", "").strip()
            level = int(t.get("level", 0))
            kind = t.get("kind", "task")

        if not name:
            continue

        current_uid += 1
        wbs_counters[level] = wbs_counters.get(level, 0) + 1
        # ریست سطوح پایین‌تر
        for lv in list(wbs_counters.keys()):
            if lv > level:
                wbs_counters[lv] = 0

        wbs_parts = []
        for lv in range(level + 1):
            wbs_parts.append(str(wbs_counters.get(lv, 0) or 1))
        wbs = ".".join(wbs_parts)

        dur_days = 1 if level == 0 else 0.5
        s = start_dt + timedelta(days=cursor_day)
        f = s + timedelta(days=max(0.5, dur_days))
        cursor_day += dur_days

        safe = html.escape(name.replace("\x00", "").strip()[:400])

        # Summary برای گروه‌ها
        is_summary = kind.endswith("_group") or kind in ("heading",)
        summary_flag = "1" if is_summary else "0"
        dur = "PT8H0M0S" if level == 0 else "PT4H0M0S"

        L.append('    <Task>')
        L.append(f'      <UID>{current_uid}</UID>')
        L.append(f'      <ID>{current_uid}</ID>')
        L.append(f'      <Name>{safe}</Name>')
        L.append('      <Type>0</Type>')
        L.append('      <IsNull>0</IsNull>')
        L.append(f'      <CreateDate>{now_iso}</CreateDate>')
        L.append(f'      <WBS>{wbs}</WBS>')
        L.append(f'      <OutlineNumber>{wbs}</OutlineNumber>')
        L.append(f'      <OutlineLevel>{level + 1}</OutlineLevel>')
        L.append('      <Priority>500</Priority>')
        L.append(f'      <Start>{_iso(s)}</Start>')
        L.append(f'      <Finish>{_iso(f)}</Finish>')
        L.append(f'      <Duration>{dur}</Duration>')
        L.append('      <DurationFormat>7</DurationFormat>')
        L.append(f'      <Work>{dur}</Work>')
        L.append(f'      <Summary>{summary_flag}</Summary>')
        L.append('      <Milestone>0</Milestone>')
        L.append('      <Critical>0</Critical>')
        L.append('      <PercentComplete>0</PercentComplete>')
        L.append('      <CalendarUID>-1</CalendarUID>')
        L.append('    </Task>')

    L.append('  </Tasks>')
    L.append('</Project>')

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))


# ==============================================================
# Worker
# ==============================================================
def _update_job(job_id, **kw):
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(kw)


def _append_log(job_id, msg):
    with JOBS_LOCK:
        if job_id in JOBS:
            ts = datetime.now().strftime("%H:%M:%S")
            JOBS[job_id]["log"].append(f"[{ts}] {msg}")


def worker(job_id, input_path, output_path, project_title, start_dt, mode):
    try:
        _update_job(job_id, status="running", progress=5)
        _append_log(job_id, f"📖 خواندن فایل: {os.path.basename(input_path)}")

        blocks = read_input_structured(input_path)
        if not blocks:
            raise ValueError(
                "محتوایی استخراج نشد. اگر PDF اسکن‌شده است، ابتدا OCR لازم است."
            )

        para_count = sum(1 for b in blocks if b["type"] == "para")
        table_count = sum(1 for b in blocks if b["type"] == "table")
        _update_job(job_id, progress=30)
        _append_log(
            job_id,
            f"📝 محتوا استخراج شد: {para_count} پاراگراف، {table_count} جدول"
        )

        if mode == "classic":
            _append_log(job_id, "🎯 حالت: تجزیه کلاسیک (خطوط شماره‌دار/بولت)")
            tasks = parse_tasks_classic(blocks, strict=True)
            task_objs = [{"name": t, "level": 0, "kind": "task"} for t in tasks]
        elif mode == "classic_all":
            _append_log(job_id, "📋 حالت: همه خطوط متنی")
            tasks = parse_tasks_classic(blocks, strict=False)
            task_objs = [{"name": t, "level": 0, "kind": "task"} for t in tasks]
        else:  # smart
            _append_log(job_id, "🧠 حالت: خلاصه‌سازی هوشمند + استخراج ساختاریافته")
            task_objs = parse_tasks_structured(blocks)

        if not task_objs:
            raise ValueError("هیچ تسکی از متن شناسایی نشد.")

        _update_job(job_id, progress=60)
        _append_log(job_id, f"🔍 {len(task_objs)} تسک/آیتم شناسایی شد")
        for i, t in enumerate(task_objs[:12], 1):
            nm = t["name"] if isinstance(t, dict) else str(t)
            preview = nm if len(nm) <= 90 else nm[:90] + "…"
            lvl = t.get("level", 0) if isinstance(t, dict) else 0
            indent = "  " * lvl
            _append_log(job_id, f"    {i}. {indent}{preview}")
        if len(task_objs) > 12:
            _append_log(job_id, f"    … و {len(task_objs) - 12} آیتم دیگر")

        _update_job(job_id, progress=80)
        _append_log(job_id, "🛠  در حال ساخت فایل MSPDI XML...")

        create_mspdi_xml(task_objs, output_path, project_title, start_dt)

        _update_job(
            job_id, progress=100, status="done",
            output_path=output_path,
            output_name=os.path.basename(output_path),
        )
        _append_log(job_id, f"✅ فایل ساخته شد: {os.path.basename(output_path)}")

    except Exception as e:
        tb = traceback.format_exc()
        _append_log(job_id, f"❌ خطا: {e}")
        _update_job(job_id, status="error", error=str(e), trace=tb)


# ==============================================================
# Routes
# ==============================================================
@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/convert", methods=["POST"])
def api_convert():
    if "file" not in request.files:
        return jsonify({"error": "فایلی ارسال نشد"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "نام فایل خالی است"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".docx", ".doc", ".pdf"):
        return jsonify({"error": f"فرمت پشتیبانی نمی‌شود: {ext}"}), 400

    project_title = (request.form.get("project_title") or "پروژه تبدیل‌شده").strip()
    start_date_str = request.form.get("start_date") or ""
    mode = (request.form.get("mode") or "smart").lower()
    if mode not in ("smart", "classic", "classic_all"):
        mode = "smart"

    try:
        start_dt = (
            datetime.strptime(start_date_str, "%Y-%m-%d")
            if start_date_str else datetime.now()
        )
    except ValueError:
        start_dt = datetime.now()

    job_id = uuid.uuid4().hex
    base_name = os.path.splitext(os.path.basename(f.filename))[0]
    safe_base = re.sub(r"[^\w\-\.]", "_", base_name, flags=re.UNICODE) or "output"
    input_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_{safe_base}{ext}")
    output_path = os.path.join(OUTPUT_FOLDER, f"{job_id}_{safe_base}.xml")

    f.save(input_path)

    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "queued",
            "progress": 0,
            "log": [],
            "output_path": None,
            "output_name": None,
            "error": None,
            "trace": None,
            "created": datetime.now().isoformat(),
        }

    t = threading.Thread(
        target=worker,
        args=(job_id, input_path, output_path, project_title, start_dt, mode),
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def api_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "job not found"}), 404
        return jsonify({
            "status": job["status"],
            "progress": job["progress"],
            "log": list(job["log"]),
            "output_name": job["output_name"],
            "error": job["error"],
        })


@app.route("/api/download/<job_id>")
def api_download(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job or not job["output_path"]:
            abort(404)
        path = job["output_path"]
        name = job["output_name"] or "output.xml"

    if not os.path.exists(path):
        abort(404)

    return send_file(path, as_attachment=True, download_name=name)


# ==============================================================
# HTML
# ==============================================================
HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>تبدیل‌گر هوشمند سند به پروژه</title>
<link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css" rel="stylesheet" />
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }

:root {
    --bg-1: #f5f3ff;
    --card-bg: rgba(255, 255, 255, 0.72);
    --card-border: rgba(255, 255, 255, 0.9);
    --card-shadow: 0 20px 60px -20px rgba(79, 70, 229, 0.25);
    --text: #1e1b4b;
    --text-muted: #6b7280;
    --primary: #6366f1;
    --primary-2: #8b5cf6;
    --primary-soft: rgba(99, 102, 241, 0.12);
    --border: #e5e7eb;
    --input-bg: #ffffff;
    --input-border: #e5e7eb;
    --success: #10b981;
    --error: #ef4444;
    --terminal-bg: #0f172a;
    --terminal-text: #e2e8f0;
    --terminal-dim: #64748b;
    --terminal-green: #4ade80;
    --terminal-red: #f87171;
    --terminal-yellow: #fbbf24;
    --terminal-cyan: #22d3ee;
}

[data-theme="dark"] {
    --bg-1: #0a0e1a;
    --card-bg: rgba(30, 27, 75, 0.55);
    --card-border: rgba(139, 92, 246, 0.25);
    --card-shadow: 0 20px 60px -20px rgba(0, 0, 0, 0.7);
    --text: #f1f5f9;
    --text-muted: #c7d2fe;
    --border: rgba(139, 92, 246, 0.2);
    --input-bg: rgba(15, 23, 42, 0.7);
    --input-border: rgba(139, 92, 246, 0.3);
    --primary-soft: rgba(139, 92, 246, 0.15);
}

body {
    font-family: 'Vazirmatn', system-ui, -apple-system, 'Segoe UI', sans-serif;
    background: var(--bg-1);
    color: var(--text);
    min-height: 100vh;
    overflow-x: hidden;
    transition: background 0.5s ease, color 0.3s ease;
    line-height: 1.6;
}

.bg-blobs { position: fixed; inset: 0; overflow: hidden; z-index: 0; pointer-events: none; }
.blob { position: absolute; border-radius: 50%; filter: blur(90px); opacity: 0.55; animation: float 22s ease-in-out infinite; }
.blob-1 { width: 500px; height: 500px; background: radial-gradient(circle, #a5b4fc, transparent 70%); top: -120px; right: -120px; }
.blob-2 { width: 600px; height: 600px; background: radial-gradient(circle, #c084fc, transparent 70%); bottom: -220px; left: -120px; animation-delay: -7s; }
.blob-3 { width: 420px; height: 420px; background: radial-gradient(circle, #f0abfc, transparent 70%); top: 40%; left: 40%; animation-delay: -14s; }
@keyframes float {
    0%, 100% { transform: translate(0, 0) scale(1); }
    33%      { transform: translate(50px, -50px) scale(1.1); }
    66%      { transform: translate(-50px, 50px) scale(0.95); }
}

.app { position: relative; z-index: 1; max-width: 1200px; margin: 0 auto; padding: 24px 20px 40px; }

.topbar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 16px 22px; background: var(--card-bg);
    backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
    border: 1px solid var(--card-border); border-radius: 20px;
    margin-bottom: 22px; box-shadow: var(--card-shadow);
}
.brand { display: flex; align-items: center; gap: 14px; }
.brand-icon {
    width: 56px; height: 56px; display: grid; place-items: center;
    font-size: 28px; border-radius: 16px;
    background: linear-gradient(135deg, var(--primary), var(--primary-2));
    box-shadow: 0 10px 22px -8px var(--primary); color: #fff; flex-shrink: 0;
}
.brand-text h1 {
    font-size: 17px; font-weight: 800; line-height: 1.35;
    background: linear-gradient(135deg, var(--primary), var(--primary-2));
    -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent;
}
.brand-text p { font-size: 12px; color: var(--text-muted); }
.theme-toggle {
    width: 44px; height: 44px; border-radius: 12px;
    border: 1px solid var(--border); background: var(--input-bg);
    cursor: pointer; font-size: 20px; display: grid; place-items: center;
    transition: all 0.2s; color: var(--text);
}
.theme-toggle:hover { transform: scale(1.05) rotate(-8deg); border-color: var(--primary); box-shadow: 0 0 0 4px var(--primary-soft); }

.main { display: flex; flex-direction: column; gap: 18px; }
.grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
@media (max-width: 900px) {
    .grid { grid-template-columns: 1fr; }
    .brand-text h1 { font-size: 15px; }
    .brand-text p { font-size: 11px; }
}

.card {
    background: var(--card-bg); backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid var(--card-border); border-radius: 20px;
    padding: 20px; box-shadow: var(--card-shadow);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
.card:hover { transform: translateY(-3px); box-shadow: 0 28px 70px -20px rgba(79, 70, 229, 0.35); }

.card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; }
.card-head h2 { font-size: 15px; font-weight: 700; flex: 1; }

.step-badge {
    width: 34px; height: 34px; display: grid; place-items: center;
    border-radius: 11px; font-size: 14px; font-weight: 800;
    background: linear-gradient(135deg, var(--primary), var(--primary-2));
    color: #fff; box-shadow: 0 8px 16px -8px var(--primary); flex-shrink: 0;
}

.dropzone {
    position: relative; border: 2px dashed var(--input-border);
    border-radius: 16px; padding: 34px 20px; text-align: center;
    cursor: pointer; transition: all 0.25s ease;
    background: var(--input-bg); overflow: hidden;
}
.dropzone::before {
    content: ''; position: absolute; inset: 0;
    background: linear-gradient(135deg, var(--primary-soft), transparent 70%);
    opacity: 0; transition: opacity 0.25s; pointer-events: none;
}
.dropzone:hover { border-color: var(--primary); transform: translateY(-2px); }
.dropzone:hover::before { opacity: 1; }
.dropzone.dragover { border-color: var(--primary); border-style: solid; transform: scale(1.02); box-shadow: 0 0 0 6px var(--primary-soft); }
.dz-content { position: relative; z-index: 1; }
.dz-icon { font-size: 46px; margin-bottom: 8px; display: inline-block; transition: transform 0.3s; }
.dropzone:hover .dz-icon { transform: translateY(-5px) scale(1.1); }
.dz-title { font-size: 15px; font-weight: 700; margin-bottom: 4px; }
.dz-sub { font-size: 12px; color: var(--text-muted); margin-bottom: 12px; }
.dz-formats { display: inline-block; padding: 4px 14px; border-radius: 20px; background: var(--primary-soft); color: var(--primary); font-size: 11px; font-weight: 700; }

.file-info {
    display: flex; align-items: center; gap: 10px;
    margin-top: 12px; padding: 12px 14px;
    background: var(--primary-soft); border: 1px solid var(--primary);
    border-radius: 12px; animation: slideDown 0.3s ease;
}
@keyframes slideDown { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: translateY(0); } }
.file-icon { font-size: 18px; }
.file-name { flex: 1; font-size: 13px; font-weight: 700; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--primary); }
.file-size { font-size: 11px; color: var(--text-muted); font-weight: 600; }
.file-remove {
    width: 28px; height: 28px; border: none; border-radius: 8px;
    background: rgba(239, 68, 68, 0.15); color: var(--error);
    cursor: pointer; font-size: 18px; line-height: 1;
    display: grid; place-items: center; transition: all 0.2s; flex-shrink: 0;
}
.file-remove:hover { background: var(--error); color: #fff; transform: rotate(90deg); }

.field { margin-bottom: 14px; }
.field:last-child { margin-bottom: 0; }
.field label { display: block; font-size: 12px; font-weight: 700; margin-bottom: 6px; color: var(--text-muted); }
.field input[type="text"], .field input[type="date"] {
    width: 100%; padding: 10px 14px;
    border: 1.5px solid var(--input-border); border-radius: 12px;
    background: var(--input-bg); color: var(--text);
    font-family: inherit; font-size: 13px;
    transition: all 0.2s; outline: none;
}
.field input:focus { border-color: var(--primary); box-shadow: 0 0 0 4px var(--primary-soft); }

.hint { font-size: 11px; color: var(--text-muted); margin-top: 8px; line-height: 1.7; }

.radio-group { display: flex; gap: 8px; flex-direction: column; }
.radio { position: relative; cursor: pointer; }
.radio input { position: absolute; opacity: 0; pointer-events: none; }
.radio span {
    display: block; padding: 10px 14px;
    border: 1.5px solid var(--input-border); border-radius: 12px;
    background: var(--input-bg); font-size: 13px; font-weight: 700;
    transition: all 0.2s; color: var(--text);
}
.radio span small { display: block; font-size: 11px; font-weight: 400; color: var(--text-muted); margin-top: 2px; }
.radio input:checked + span { border-color: var(--primary); background: var(--primary-soft); color: var(--primary); box-shadow: 0 0 0 4px var(--primary-soft); }
.radio input:checked + span small { color: var(--primary); opacity: 0.85; }
.radio:hover span { border-color: var(--primary); }

.progress-wrap { margin-bottom: 16px; }
.progress-bar {
    height: 12px; border-radius: 20px; background: var(--input-bg);
    border: 1px solid var(--input-border); overflow: hidden; position: relative;
}
.progress-fill {
    height: 100%; width: 0%; border-radius: 20px;
    background: linear-gradient(90deg, var(--primary), var(--primary-2));
    transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative; overflow: hidden;
}
.progress-fill::after {
    content: ''; position: absolute; inset: 0;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.5), transparent);
    animation: shimmer 1.5s infinite;
}
@keyframes shimmer { from { transform: translateX(-100%); } to { transform: translateX(100%); } }
.progress-text { display: flex; justify-content: space-between; margin-top: 8px; font-size: 12px; color: var(--text-muted); font-weight: 700; }

.actions { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; }
.btn {
    display: inline-flex; align-items: center; justify-content: center;
    gap: 8px; padding: 12px 24px; border-radius: 12px; border: none;
    font-family: inherit; font-size: 14px; font-weight: 700;
    cursor: pointer; transition: all 0.2s; text-decoration: none;
    flex: 1; min-width: 160px; color: #fff;
}
.btn-primary { background: linear-gradient(135deg, var(--primary), var(--primary-2)); box-shadow: 0 12px 26px -10px var(--primary); }
.btn-primary:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 16px 32px -10px var(--primary); filter: brightness(1.06); }
.btn-primary:active:not(:disabled) { transform: translateY(0); }
.btn-primary:disabled { opacity: 0.45; cursor: not-allowed; box-shadow: none; }
.btn-success { background: linear-gradient(135deg, var(--success), #14b8a6); box-shadow: 0 12px 26px -10px var(--success); }
.btn-success:hover { transform: translateY(-2px); box-shadow: 0 16px 32px -10px var(--success); filter: brightness(1.06); }
.btn-icon { font-size: 16px; }
.btn-ghost {
    padding: 6px 14px; border: 1px solid var(--border);
    background: transparent; color: var(--text-muted);
    border-radius: 10px; font-family: inherit;
    font-size: 12px; font-weight: 700; cursor: pointer; transition: all 0.2s;
}
.btn-ghost:hover { border-color: var(--primary); color: var(--primary); background: var(--primary-soft); }

.tip-box {
    display: flex; gap: 10px; padding: 12px 14px;
    border-radius: 12px; background: var(--primary-soft);
    border: 1px solid var(--border); font-size: 12px;
    color: var(--text-muted); line-height: 1.7;
}
.tip-icon { font-size: 16px; flex-shrink: 0; }
.tip-box code { background: rgba(99, 102, 241, 0.15); color: var(--primary); padding: 1px 6px; border-radius: 5px; font-family: 'Consolas', monospace; font-size: 11px; direction: ltr; display: inline-block; }
.tip-box strong { color: var(--text); }

.terminal {
    background: var(--terminal-bg); border-radius: 14px; padding: 16px;
    font-family: 'Vazirmatn', 'Consolas', monospace;
    font-size: 12.5px; color: var(--terminal-text);
    min-height: 180px; max-height: 340px;
    overflow-y: auto; line-height: 1.9;
    border: 1px solid rgba(148, 163, 184, 0.15);
    box-shadow: inset 0 2px 12px rgba(0, 0, 0, 0.35);
}
.terminal::-webkit-scrollbar { width: 8px; }
.terminal::-webkit-scrollbar-track { background: transparent; }
.terminal::-webkit-scrollbar-thumb { background: rgba(148, 163, 184, 0.3); border-radius: 4px; }
.terminal::-webkit-scrollbar-thumb:hover { background: rgba(148, 163, 184, 0.5); }
.terminal-line { padding: 2px 0; white-space: pre-wrap; word-break: break-word; animation: fadeIn 0.25s ease; color: var(--terminal-text); }
@keyframes fadeIn { from { opacity: 0; transform: translateX(8px); } to { opacity: 1; transform: translateX(0); } }
.terminal-line.error   { color: var(--terminal-red); }
.terminal-line.success { color: var(--terminal-green); }
.terminal-line.info    { color: var(--terminal-cyan); }
.terminal-line.warn    { color: var(--terminal-yellow); }
.terminal-empty { color: var(--terminal-dim); font-style: italic; text-align: center; padding: 50px 0; }

.footer { text-align: center; padding: 24px 0 0; font-size: 12px; color: var(--text-muted); font-weight: 600; }

.toast-wrap { position: fixed; bottom: 24px; left: 24px; display: flex; flex-direction: column; gap: 10px; z-index: 999; max-width: 340px; }
.toast {
    padding: 14px 18px; border-radius: 14px; background: var(--card-bg);
    backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
    border: 1px solid var(--card-border);
    box-shadow: 0 20px 45px -12px rgba(0, 0, 0, 0.35);
    font-size: 13px; font-weight: 700; display: flex; align-items: center; gap: 10px;
    animation: slideIn 0.35s ease; color: var(--text);
}
@keyframes slideIn { from { opacity: 0; transform: translateY(20px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }
.toast.success { border-right: 4px solid var(--success); }
.toast.error   { border-right: 4px solid var(--error); }
.toast.info    { border-right: 4px solid var(--primary); }

.hidden { display: none !important; }
</style>
</head>
<body>

<div class="bg-blobs">
    <div class="blob blob-1"></div>
    <div class="blob blob-2"></div>
    <div class="blob blob-3"></div>
</div>

<div class="app">

    <header class="topbar">
        <div class="brand">
            <div class="brand-icon">📊</div>
            <div class="brand-text">
                <h1>تبدیل‌گر هوشمند سند به پروژه</h1>
                <p>Word / PDF &nbsp;→&nbsp; Microsoft Project (MSPDI XML)</p>
            </div>
        </div>
        <button class="theme-toggle" id="themeToggle" aria-label="تغییر تم">
            <span class="theme-icon">🌙</span>
        </button>
    </header>

    <main class="main">
        <div class="grid">

            <section class="card">
                <div class="card-head">
                    <span class="step-badge">۱</span>
                    <h2>انتخاب فایل ورودی</h2>
                </div>

                <div class="dropzone" id="dropzone">
                    <input type="file" id="fileInput" accept=".docx,.doc,.pdf" hidden />
                    <div class="dz-content">
                        <div class="dz-icon">📄</div>
                        <div class="dz-title">فایل را اینجا رها کنید</div>
                        <div class="dz-sub">یا کلیک کنید تا انتخاب شود</div>
                        <div class="dz-formats">.docx &nbsp;·&nbsp; .doc &nbsp;·&nbsp; .pdf</div>
                    </div>
                </div>

                <div class="file-info hidden" id="fileInfo">
                    <span class="file-icon">📎</span>
                    <span class="file-name" id="fileName"></span>
                    <span class="file-size" id="fileSize"></span>
                    <button class="file-remove" id="fileRemove" title="حذف">×</button>
                </div>
            </section>

            <section class="card">
                <div class="card-head">
                    <span class="step-badge">۲</span>
                    <h2>تنظیمات خروجی</h2>
                </div>

                <div class="field">
                    <label for="projectTitle">عنوان پروژه</label>
                    <input type="text" id="projectTitle" value="پروژه تبدیل‌شده" placeholder="مثلاً: پروژه توسعه نرم‌افزار" />
                </div>

                <div class="field">
                    <label for="startDate">تاریخ شروع</label>
                    <input type="date" id="startDate" />
                </div>

                <div class="field">
                    <label>حالت تجزیه متن</label>
                    <div class="radio-group">
                        <label class="radio">
                            <input type="radio" name="mode" value="smart" checked />
                            <span>🧠 خلاصه‌سازی هوشمند <small>جدول‌ها، اعداد، فازها و اهداف را استخراج می‌کند — تمیزترین خروجی</small></span>
                        </label>
                        <label class="radio">
                            <input type="radio" name="mode" value="classic" />
                            <span>🎯 تجزیه کلاسیک <small>فقط خطوط شماره‌دار و بولت‌دار را تسک می‌کند</small></span>
                        </label>
                        <label class="radio">
                            <input type="radio" name="mode" value="classic_all" />
                            <span>📋 همه خطوط <small>تمام خطوط متنی را به تسک تبدیل می‌کند (خروجی بزرگ)</small></span>
                        </label>
                    </div>
                    <p class="hint">
                        حالت هوشمند جدول‌ها را به تسک‌های ساختاریافته، اعداد مهم را به آیتم‌های کلیدی و فازها را به گروه‌بندی WBS تبدیل می‌کند.
                    </p>
                </div>
            </section>
        </div>

        <section class="card">
            <div class="card-head">
                <span class="step-badge">۳</span>
                <h2>اجرای تبدیل</h2>
            </div>

            <div class="progress-wrap">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div class="progress-text">
                    <span id="progressLabel">آماده برای شروع</span>
                    <span id="progressPct">۰٪</span>
                </div>
            </div>

            <div class="actions">
                <button class="btn btn-primary" id="convertBtn" disabled>
                    <span class="btn-icon">🚀</span>
                    <span>شروع تبدیل</span>
                </button>
                <a class="btn btn-success hidden" id="downloadBtn" href="#" download>
                    <span class="btn-icon">⬇️</span>
                    <span>دانلود خروجی XML</span>
                </a>
            </div>

            <div class="tip-box">
                <span class="tip-icon">💡</span>
                <div>
                    <strong>تبدیل به MPP:</strong>
                    فایل XML دانلود شده را در Microsoft Project باز کنید و از منوی
                    <code>File → Save As</code> فرمت <code>.mpp</code> را انتخاب کنید.
                </div>
            </div>
        </section>

        <section class="card">
            <div class="card-head">
                <span class="step-badge">📋</span>
                <h2>گزارش عملیات</h2>
                <button class="btn-ghost" id="clearLogBtn">پاک کردن</button>
            </div>
            <div class="terminal" id="terminal">
                <div class="terminal-empty">منتظر شروع عملیات...</div>
            </div>
        </section>
    </main>

    <footer class="footer">
        ساخته‌شده با ❤️ &nbsp;·&nbsp; نسخه وب ۲.۰
    </footer>
</div>

<div class="toast-wrap" id="toastWrap"></div>

<script>
const $ = (s) => document.querySelector(s);

const dropzone     = $('#dropzone');
const fileInput    = $('#fileInput');
const fileInfo     = $('#fileInfo');
const fileName     = $('#fileName');
const fileSize     = $('#fileSize');
const fileRemove   = $('#fileRemove');
const convertBtn   = $('#convertBtn');
const downloadBtn  = $('#downloadBtn');
const progressFill = $('#progressFill');
const progressLbl  = $('#progressLabel');
const progressPct  = $('#progressPct');
const terminal     = $('#terminal');
const projectTitle = $('#projectTitle');
const startDate    = $('#startDate');
const clearLogBtn  = $('#clearLogBtn');
const themeToggle  = $('#themeToggle');

let currentFile  = null;
let currentJobId = null;
let pollTimer    = null;
let lastLogCount = 0;

(function () {
    const today = new Date();
    startDate.value = today.toISOString().split('T')[0];
})();

(function () {
    const saved = localStorage.getItem('projectTitle');
    if (saved) projectTitle.value = saved;
    projectTitle.addEventListener('input', () => {
        localStorage.setItem('projectTitle', projectTitle.value);
    });
})();

(function () {
    const saved = localStorage.getItem('theme') || 'light';
    if (saved === 'dark') {
        document.documentElement.setAttribute('data-theme', 'dark');
        themeToggle.querySelector('.theme-icon').textContent = '☀️';
    }
})();

themeToggle.addEventListener('click', () => {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    if (isDark) {
        document.documentElement.removeAttribute('data-theme');
        themeToggle.querySelector('.theme-icon').textContent = '🌙';
        localStorage.setItem('theme', 'light');
    } else {
        document.documentElement.setAttribute('data-theme', 'dark');
        themeToggle.querySelector('.theme-icon').textContent = '☀️';
        localStorage.setItem('theme', 'dark');
    }
});

dropzone.addEventListener('click', () => fileInput.click());

['dragenter', 'dragover'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
        e.preventDefault(); e.stopPropagation();
        dropzone.classList.add('dragover');
    });
});

['dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
        e.preventDefault(); e.stopPropagation();
        dropzone.classList.remove('dragover');
    });
});

dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files && files[0]) setFile(files[0]);
});

fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
});

function setFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['docx', 'doc', 'pdf'].includes(ext)) {
        toast('فقط فرمت‌های .docx، .doc و .pdf پشتیبانی می‌شوند', 'error');
        return;
    }
    currentFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    fileInfo.classList.remove('hidden');
    convertBtn.disabled = false;
    downloadBtn.classList.add('hidden');
    toast('فایل با موفقیت انتخاب شد', 'success');
}

fileRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    currentFile = null;
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    convertBtn.disabled = true;
    downloadBtn.classList.add('hidden');
});

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

convertBtn.addEventListener('click', async () => {
    if (!currentFile) return;

    const mode = document.querySelector('input[name="mode"]:checked').value;
    const fd = new FormData();
    fd.append('file', currentFile);
    fd.append('project_title', projectTitle.value.trim() || 'پروژه تبدیل‌شده');
    fd.append('start_date', startDate.value);
    fd.append('mode', mode);

    resetUI();
    convertBtn.disabled = true;
    downloadBtn.classList.add('hidden');
    progressLbl.textContent = 'در حال ارسال فایل...';

    try {
        const res = await fetch('/api/convert', { method: 'POST', body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'خطا در ارسال فایل');
        currentJobId = data.job_id;
        progressLbl.textContent = 'در حال پردازش...';
        startPolling();
    } catch (err) {
        toast(err.message, 'error');
        convertBtn.disabled = false;
        progressLbl.textContent = 'خطا در ارسال';
    }
});

function startPolling() {
    if (pollTimer) clearInterval(pollTimer);

    pollTimer = setInterval(async () => {
        try {
            const res = await fetch(`/api/status/${currentJobId}`);
            const data = await res.json();

            updateProgress(data.progress);
            updateLog(data.log);

            if (data.status === 'done') {
                clearInterval(pollTimer);
                pollTimer = null;
                progressLbl.textContent = '✓ تبدیل با موفقیت انجام شد';
                downloadBtn.href = `/api/download/${currentJobId}`;
                downloadBtn.classList.remove('hidden');
                convertBtn.disabled = false;
                toast('فایل آماده دانلود است', 'success');
            } else if (data.status === 'error') {
                clearInterval(pollTimer);
                pollTimer = null;
                progressLbl.textContent = '✗ خطا در تبدیل';
                convertBtn.disabled = false;
                toast('خطا: ' + (data.error || 'ناشناخته'), 'error');
            } else if (data.status === 'running') {
                progressLbl.textContent = 'در حال پردازش...';
            }
        } catch (_) {}
    }, 500);
}

function updateProgress(pct) {
    pct = Math.max(0, Math.min(100, pct || 0));
    progressFill.style.width = pct + '%';
    progressPct.textContent = toFa(pct) + '٪';
}

function updateLog(logs) {
    if (!logs) return;

    if (logs.length === 0) {
        if (lastLogCount !== 0) {
            terminal.innerHTML = '<div class="terminal-empty">منتظر شروع عملیات...</div>';
            lastLogCount = 0;
        }
        return;
    }

    if (lastLogCount === 0) terminal.innerHTML = '';

    for (let i = lastLogCount; i < logs.length; i++) {
        const line = logs[i];
        const div = document.createElement('div');
        div.className = 'terminal-line ' + classify(line);
        div.textContent = line;
        terminal.appendChild(div);
    }

    lastLogCount = logs.length;
    terminal.scrollTop = terminal.scrollHeight;
}

function classify(line) {
    if (line.includes('❌') || line.includes('خطا')) return 'error';
    if (line.includes('✅') || line.includes('موفق')) return 'success';
    if (line.includes('🔍') || line.includes('📝') || line.includes('📖')) return 'info';
    if (line.includes('⚠️')) return 'warn';
    return '';
}

function resetUI() {
    terminal.innerHTML = '<div class="terminal-empty">در حال شروع عملیات...</div>';
    lastLogCount = 0;
    updateProgress(0);
    progressPct.textContent = '۰٪';
    progressLbl.textContent = 'آماده';
}

clearLogBtn.addEventListener('click', () => {
    terminal.innerHTML = '<div class="terminal-empty">گزارش پاک شد</div>';
    lastLogCount = 0;
});

function toFa(n) {
    return String(n).replace(/\d/g, d => '۰۱۲۳۴۵۶۷۸۹'[d]);
}

function toast(msg, type = 'info') {
    const wrap = $('#toastWrap');
    const div = document.createElement('div');
    div.className = 'toast ' + type;
    const icon = type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️';
    div.innerHTML = `<span>${icon}</span><span>${msg}</span>`;
    wrap.appendChild(div);

    setTimeout(() => {
        div.style.transition = 'opacity 0.3s, transform 0.3s';
        div.style.opacity = '0';
        div.style.transform = 'translateY(20px)';
        setTimeout(() => div.remove(), 320);
    }, 3800);
}
</script>
</body>
</html>
"""


# ==============================================================
# Entry point
# ==============================================================
def main():
    host = "127.0.0.1"
    port = 5000
    url = f"http://{host}:{port}"

    print("=" * 60)
    print("  🌐  تبدیل‌گر هوشمند Word/PDF به Microsoft Project  v2.0")
    print(f"  آدرس: {url}")
    print("  برای توقف: Ctrl+C")
    print("=" * 60)

    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host=host, port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()