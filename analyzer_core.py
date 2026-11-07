# pyright: reportMissingImports=false, reportMissingModuleSource=false
# (python-docx / openpyxl are imported lazily; install them with: pip install -r requirements.txt)
import re
import os
import sys
import math
import time
import queue
import shutil
import subprocess
import tempfile
import threading
from collections import Counter
from datetime import datetime
from pathlib import Path

# Optional/third-party readers. The application starts even if one is missing,
# and gives a clear installation message when that feature is used.
try:
    import fitz  # PyMuPDF  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
except ImportError:
    fitz = None

try:
    from docx import Document  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from docx.shared import Pt, RGBColor, Cm  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from docx.enum.text import WD_ALIGN_PARAGRAPH  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from docx.oxml import OxmlElement  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from docx.oxml.ns import qn  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
except ImportError:
    Document = None

try:
    from PIL import Image  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
except ImportError:
    Image = None

try:
    import pytesseract  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
except ImportError:
    pytesseract = None

# Excel export (pip install openpyxl)
try:
    import openpyxl  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.utils import get_column_letter  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.chart import BarChart as XlBarChart, Reference as XlReference  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.worksheet.properties import PageSetupProperties  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
except ImportError:
    openpyxl = None



# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_TITLE = "AI Resume / CV Analyzer"


# ============================================================
# MASTER SKILL DICTIONARY
# ============================================================

SKILLS = {

    "Programming & Technology": [
        "python", "java", "c", "c++", "c#", "javascript",
        "typescript", "r programming", "r language", "golang",
        "ruby", "php", "kotlin", "swift", "scala", "sql",
        "html", "css", "react", "react.js", "angular", "vue",
        "node.js", "nodejs", "django", "flask", "fastapi",
        "spring boot", "asp.net", ".net", "rest api",
        "api development", "git", "github", "gitlab",
        "docker", "kubernetes", "jenkins", "ci/cd", "linux",
        "bash"
    ],

    "Data & Analytics": [
        "data analysis", "data analytics", "data science",
        "data visualization", "data mining", "machine learning",
        "deep learning", "artificial intelligence", "ai",
        "natural language processing", "nlp", "computer vision",
        "predictive analytics", "prescriptive analytics",
        "statistics", "statistical analysis",
        "business intelligence", "power bi", "tableau",
        "qlikview", "looker", "excel", "advanced excel",
        "microsoft excel", "pivot tables", "vlookup", "xlookup",
        "power query", "power pivot", "numpy", "pandas",
        "scipy", "matplotlib", "seaborn", "scikit-learn",
        "tensorflow", "pytorch", "keras", "spark", "pyspark",
        "hadoop", "hive", "mongodb", "mysql", "postgresql",
        "oracle", "sqlite", "etl", "data engineering",
        "data modeling", "data cleaning", "data wrangling",
        "performance metrics", "analytics"
    ],

    "Digital Marketing": [
        "digital marketing", "marketing", "digital advertising",
        "meta ads", "facebook ads", "instagram ads",
        "google ads", "google analytics", "google ads manager",
        "facebook business manager", "media buying",
        "performance marketing", "performance metrics",
        "a/b testing", "ab testing", "cro",
        "conversion rate optimization", "seo", "sem",
        "social media marketing", "content marketing",
        "email marketing", "lead generation", "lead management",
        "marketing analytics", "campaign management",
        "campaign optimization", "customer acquisition",
        "retargeting", "remarketing", "roi", "roas",
        "marketing strategy", "brand management",
        "e-commerce", "ecommerce", "d2c",
        "digital strategy", "crm", "customer relationship management"
    ],

    "Finance & Accounting": [
        "financial analysis", "financial modeling",
        "financial management", "corporate finance",
        "investment analysis", "portfolio management",
        "risk management", "credit analysis", "valuation",
        "business valuation", "company valuation", "dcf",
        "discounted cash flow", "npv", "irr", "wacc",
        "cost of capital", "cost of equity", "cost of debt",
        "accounting", "financial accounting",
        "management accounting", "cost accounting", "taxation",
        "auditing", "internal audit", "statutory audit", "gst",
        "income tax", "balance sheet", "profit and loss",
        "p&l", "cash flow", "budgeting", "forecasting",
        "financial projections", "accounts payable",
        "accounts receivable"
    ],

    "Business & Management": [
        "business analysis", "business analytics",
        "business development", "business strategy",
        "strategic planning", "project management",
        "program management", "product management",
        "operations management", "process improvement",
        "business process management", "bpm",
        "lean", "lean six sigma", "six sigma",
        "quality management", "supply chain management",
        "procurement", "inventory management", "logistics",
        "vendor management", "stakeholder management",
        "change management", "risk assessment", "kpi",
        "performance management", "erp", "sap",
        "crm", "market research", "competitive analysis",
        "pestel", "swot", "strategy"
    ],

    "Human Resources": [
        "human resources", "hr", "human resource management",
        "talent management", "talent acquisition", "recruitment",
        "employee engagement", "employee relations",
        "performance appraisal", "performance management",
        "training and development", "learning and development",
        "organizational behavior", "organizational development",
        "workforce planning", "hr analytics", "compensation",
        "payroll", "labor relations", "succession planning",
        "job analysis"
    ],

    "Soft Skills": [
        "communication", "verbal communication",
        "written communication", "leadership", "teamwork",
        "team player", "collaboration", "problem solving",
        "problem-solving", "critical thinking",
        "analytical thinking", "decision making",
        "decision-making", "time management", "adaptability",
        "flexibility", "creativity", "innovation", "negotiation",
        "conflict resolution", "presentation", "public speaking",
        "interpersonal skills", "emotional intelligence",
        "attention to detail", "organization",
        "organizational skills", "multitasking", "work ethic",
        "customer service", "relationship management",
        "mentoring", "coaching"
    ]
}


# ============================================================
# DISPLAY NAME NORMALIZATION
# ============================================================

DISPLAY_NAMES = {
    "r programming": "R",
    "r language": "R",
    "c++": "C++",
    "c#": "C#",
    "javascript": "JavaScript",
    "typescript": "TypeScript",
    "node.js": "Node.js",
    "nodejs": "Node.js",
    "react.js": "React",
    "react": "React",
    "power bi": "Power BI",
    "aws": "AWS",
    "amazon web services": "AWS",
    "gcp": "Google Cloud",
    "google cloud": "Google Cloud",
    "azure": "Azure",
    "microsoft azure": "Azure",
    "ai": "Artificial Intelligence",
    "artificial intelligence": "Artificial Intelligence",
    "nlp": "Natural Language Processing",
    "natural language processing": "Natural Language Processing",
    "p&l": "Profit & Loss",
    "profit and loss": "Profit & Loss",
    "team player": "Teamwork",
    "problem-solving": "Problem Solving",
    "problem solving": "Problem Solving",
    "decision-making": "Decision Making",
    "decision making": "Decision Making",
    "a/b testing": "A/B Testing",
    "ab testing": "A/B Testing",
    "cro": "CRO",
    "d2c": "D2C",
    "e-commerce": "E-Commerce",
    "ecommerce": "E-Commerce",
    "roi": "ROI",
    "roas": "ROAS",
    "crm": "CRM"
}


def display_skill(skill):
    return DISPLAY_NAMES.get(
        skill.lower(),
        skill.title()
    )


# ============================================================
# TEXT NORMALIZATION
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = text.replace("\x00", " ")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text
    )

    return text.strip()


# ============================================================
# ROBUST DOCUMENT EXTRACTION
# ============================================================

def normalize_text(text):
    """Clean extracted text while preserving useful line breaks."""
    if not text:
        return ""

    text = str(text).replace("\x00", " ")
    if "â€" in text or "Ã" in text:
        try:  # UTF-8 text that was decoded as Windows-1252 ("â€“" -> "–")
            text = text.encode("cp1252").decode("utf-8")
        except Exception:
            for bad, good in (("â€“", "–"), ("â€”", "—"), ("â€™", "’"), ("â€˜", "‘"),
                              ("â€œ", "“"), ("â€\x9d", "”"), ("â€¢", "•")):
                text = text.replace(bad, good)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u00a0", " ")
    text = text.replace("\uf0b7", "•")

    # Fix common PDF soft-hyphen artifacts.
    text = text.replace("\u00ad", "")
    text = re.sub(r"-\n(?=[a-z])", "", text)

    # Remove repeated spaces but keep newlines.
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


# ------------------------------------------------------------
# Automatic installation of missing reader libraries
# ------------------------------------------------------------
# If a reader library (PyMuPDF, Pillow, OCR engine ...) is not installed, the
# app installs it on demand with pip instead of failing. This is skipped for
# packaged (.exe) builds, and can be switched off with the environment
# variable RESUME_ANALYZER_NO_INSTALL=1.

import importlib
import importlib.util
import zlib

_INSTALL_LOCK = threading.RLock()
_PIP_ATTEMPTED = {}
_SUBPROCESS_KW = (
    {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    if os.name == "nt" else {}
)


def _python_executable():
    exe = sys.executable or ""
    if os.name == "nt" and exe.lower().endswith("pythonw.exe"):
        candidate = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(candidate):
            return candidate
    return exe


def _private_libs_dir():
    """Per-user folder for auto-installed helpers (never touches system Python)."""
    try:
        return os.path.join(str(Path.home()), ".resume_analyzer", "libs")
    except Exception:
        return os.path.join(tempfile.gettempdir(), "resume_analyzer_libs")


def _use_private_libs():
    folder = _private_libs_dir()
    if os.path.isdir(folder) and folder not in sys.path:
        sys.path.append(folder)
        try:
            importlib.invalidate_caches()
        except Exception:
            pass


def _pip_install(*packages, timeout=900):
    """Install packages with pip. Returns True on success. Tried once per set."""
    key = tuple(packages)
    with _INSTALL_LOCK:
        if key in _PIP_ATTEMPTED:
            return _PIP_ATTEMPTED[key]
        ok = False
        allowed = (
            not getattr(sys, "frozen", False)
            and os.environ.get("RESUME_ANALYZER_NO_INSTALL") != "1"
            and _python_executable()
        )
        if allowed:
            base = [
                _python_executable(), "-m", "pip", "install", "--quiet",
                "--disable-pip-version-check", "--no-input",
            ]
            target = _private_libs_dir()
            # 1) normal install, 2) private folder (works on locked-down /
            # "externally managed" Python installs such as Linux and Homebrew).
            for extra in ([], ["--target", target, "--upgrade"]):
                try:
                    if extra:
                        os.makedirs(target, exist_ok=True)
                    result = subprocess.run(
                        base + extra + list(packages),
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        timeout=timeout, **_SUBPROCESS_KW
                    )
                    if result.returncode == 0:
                        ok = True
                        break
                except Exception:
                    break
        _use_private_libs()
        try:
            importlib.invalidate_caches()
        except Exception:
            pass
        _PIP_ATTEMPTED[key] = ok
        return ok


def _import_or_install(import_name, pip_name=None):
    """Import a module; if it is missing, pip-install it once and retry."""
    _use_private_libs()
    try:
        return importlib.import_module(import_name)
    except Exception:
        pass
    with _INSTALL_LOCK:
        try:
            return importlib.import_module(import_name)
        except Exception:
            pass
        if _pip_install(pip_name or import_name):
            try:
                return importlib.import_module(import_name)
            except Exception:
                return None
    return None


def _ensure_pillow():
    """Return the PIL.Image module (installing Pillow if needed) or None."""
    global Image
    if Image is None:
        Image = _import_or_install("PIL.Image", "pillow")
    return Image


# ------------------------------------------------------------
# OCR engines: Tesseract (if installed) -> RapidOCR (pip only, no system install)
# ------------------------------------------------------------

def _tesseract_exe():
    exe = shutil.which("tesseract")
    if exe:
        return exe
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        "/usr/bin/tesseract",
    ]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(os.path.join(local, "Programs", "Tesseract-OCR", "tesseract.exe"))
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _ocr_available():
    """True when the Tesseract program and its Python wrapper are both usable."""
    global pytesseract
    exe = _tesseract_exe()
    if not exe:
        return False
    if pytesseract is None:
        pytesseract = _import_or_install("pytesseract")
        if pytesseract is None:
            return False
    try:
        pytesseract.pytesseract.tesseract_cmd = exe
        return True
    except Exception:
        return False


_RAPID = {"engine": None, "failed": False}


def _load_rapid():
    try:
        from rapidocr_onnxruntime import RapidOCR  # type: ignore
        return RapidOCR()
    except Exception:
        pass
    try:
        from rapidocr import RapidOCR  # type: ignore
        return RapidOCR()
    except Exception:
        return None


def _rapid_engine():
    with _INSTALL_LOCK:
        if _RAPID["engine"] is not None:
            return _RAPID["engine"]
        if _RAPID["failed"]:
            return None
        engine = _load_rapid()
        if engine is None and _pip_install("rapidocr-onnxruntime"):
            engine = _load_rapid()
        if engine is None and _pip_install("rapidocr", "onnxruntime"):
            engine = _load_rapid()
        if engine is None:
            _RAPID["failed"] = True
        _RAPID["engine"] = engine
        return engine


def _rapid_text(image):
    engine = _rapid_engine()
    if engine is None:
        raise RuntimeError("no OCR engine available")
    import numpy as np  # installed together with onnxruntime

    output = engine(np.array(image.convert("RGB")))
    items = []
    if isinstance(output, tuple):
        for row in (output[0] or []):
            items.append((row[0], row[1]))
    else:
        boxes = getattr(output, "boxes", None)
        texts = getattr(output, "txts", None)
        if boxes is not None and texts is not None:
            items = list(zip(boxes, texts))

    cells = []
    for box, text in items:
        text = str(text).strip()
        if not text:
            continue
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        cells.append({
            "text": text, "x1": min(xs), "x2": max(xs),
            "cy": (min(ys) + max(ys)) / 2.0, "h": max(max(ys) - min(ys), 1.0),
        })
    if not cells:
        return ""

    cells.sort(key=lambda c: (c["cy"], c["x1"]))
    typical = sorted(c["h"] for c in cells)[len(cells) // 2]
    lines = []
    for cell in cells:
        last = lines[-1] if lines else None
        if (
            last
            and abs(cell["cy"] - last["cy"]) < 0.6 * typical
            and 0 <= cell["x1"] - last["x2"] < 1.5 * typical
        ):
            last["text"] += " " + cell["text"]
            last["x2"] = max(last["x2"], cell["x2"])
        else:
            lines.append(dict(cell))
    return "\n".join(line["text"] for line in lines)


def ocr_image(image):
    """OCR one PIL image with whichever engine is available."""
    if _ensure_pillow() is None:
        raise RuntimeError(
            "Reading scanned files requires Pillow. Install it with: pip install pillow"
        )

    errors = []
    engine_ran = False

    if _ocr_available():
        try:
            engine_ran = True
            text = normalize_text(
                pytesseract.image_to_string(image, config="--psm 6", timeout=180)
            )
            if text:
                return text
        except Exception as error:
            errors.append(f"Tesseract: {error}")

    try:
        text = normalize_text(_rapid_text(image))
        engine_ran = True
        if text:
            return text
    except Exception as error:
        errors.append(f"RapidOCR: {error}")

    if engine_ran:
        return ""
    raise RuntimeError(
        "No OCR engine could be started for this scanned file. Run:  "
        "pip install rapidocr-onnxruntime   (or install Tesseract OCR)."
        + (f"\n\nDetails: {'; '.join(errors)}" if errors else "")
    )


def _text_looks_valid(text):
    """False for empty / garbled extraction (broken font maps, (cid:12) junk)."""
    if not text:
        return False
    sample = text[:20000]
    if sample.count("(cid:") > 5:
        return False
    letters = sum(ch.isalpha() for ch in sample)
    if letters < 20:
        return False
    tokens = sample.split()
    if len(tokens) >= 20:
        singles = sum(1 for t in tokens if len(t) == 1 and t.isalpha())
        if singles / len(tokens) > 0.45:  # "S i d d h a n t" style extraction
            return False
    readable = sum(ch.isalnum() or ch.isspace() or ch in ".,;:-–—/@+()&%'\"•|#_" for ch in sample)
    return readable / max(len(sample), 1) >= 0.80



import zipfile
import html as _html
import xml.etree.ElementTree as ET


# ------------------------------------------------------------
# helpers
# ------------------------------------------------------------

def _pymupdf():
    """Return a working PyMuPDF module, or None (installs it if missing).
    Handles 'import pymupdf' and the wrong 'fitz' package from PyPI."""
    def usable():
        candidates = []
        if fitz is not None:
            candidates.append(fitz)
        try:
            import pymupdf  # type: ignore
            candidates.append(pymupdf)
        except Exception:
            pass
        for module in candidates:
            if hasattr(module, "open"):
                return module
        return None

    module = usable()
    if module is None and _import_or_install("pymupdf") is not None:
        module = usable()
    return module


def _decode_bytes(data):
    """Decode bytes using BOM detection and common encodings."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            return data.decode("utf-16")
        except UnicodeDecodeError:
            pass
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace")
    if len(data) >= 4 and data.count(b"\x00") > len(data) * 0.25:
        # UTF-16 without a byte-order mark
        encoding = "utf-16-le" if data[1:2] == b"\x00" else "utf-16-be"
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    for encoding in ("utf-8", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _strip_html(raw):
    raw = re.sub(r"(?is)<(script|style|head)[^>]*>.*?</\1>", " ", raw)
    raw = re.sub(r"(?i)<br\s*/?>|</(p|div|li|tr|h[1-6]|table|section)>", "\n", raw)
    raw = re.sub(r"(?i)</t[dh]>", " | ", raw)
    raw = re.sub(r"<[^>]+>", " ", raw)
    return _html.unescape(raw)


def _binary_strings(file_path):
    """Last resort for legacy binary files: pull out readable text runs."""
    with open(file_path, "rb") as file:
        data = file.read()
    ascii_runs = re.findall(rb"[\x20-\x7e\xa0-\xff\t\r\n]{6,}", data)
    ascii_text = "\n".join(run.decode("cp1252", errors="ignore") for run in ascii_runs)
    wide_runs = re.findall(rb"(?:[\x20-\x7e\t\r\n]\x00){6,}", data)
    wide_text = "\n".join(run.decode("utf-16le", errors="ignore") for run in wide_runs)
    best = wide_text if len(wide_text) > len(ascii_text) * 0.6 else ascii_text
    return normalize_text(best.replace("\r", "\n"))


# ------------------------------------------------------------
# PDF
# ------------------------------------------------------------

def _order_blocks(blocks, page_width, page_height):
    """
    Reading order for the text blocks of one page. Single-column pages are read
    top to bottom; pages with a side column (sidebar CVs, two-column CVs) are read
    one column after the other, so the columns are not mixed line by line.
    """
    default = sorted(blocks, key=lambda b: (round(b[1], 1), round(b[0], 1)))
    if len(blocks) < 8:
        return default

    best = None
    for edge in sorted({round(b[2]) for b in blocks}):
        left = [i for i, b in enumerate(blocks) if b[2] <= edge + 1]
        right = [i for i, b in enumerate(blocks) if b[0] >= edge + 8]
        if len(left) < 3 or len(right) < 3:
            continue
        side = set(left) | set(right)
        cross = [blocks[i] for i in range(len(blocks)) if i not in side]
        top = max(min(blocks[i][1] for i in left), min(blocks[i][1] for i in right))
        bottom = min(max(blocks[i][3] for i in left), max(blocks[i][3] for i in right))
        if bottom - top < 0.30 * page_height:
            continue
        if any(b[3] > top + 2 and b[1] < bottom - 2 for b in cross):
            continue
        gap = min(blocks[i][0] for i in right) - max(blocks[i][2] for i in left)
        if gap < 8:
            continue
        if best is None or gap > best[0]:
            best = (gap, set(left), set(right), top, bottom)

    if best is None:
        return default

    _gap, left, right, top, bottom = best
    key = lambda b: (round(b[1], 1), round(b[0], 1))
    head = [b for b in blocks if b[3] <= top + 2]
    head_ids = {id(b) for b in head}
    left_col = [blocks[i] for i in left if id(blocks[i]) not in head_ids]
    right_col = [blocks[i] for i in right if id(blocks[i]) not in head_ids]
    placed = head_ids | {id(b) for b in left_col} | {id(b) for b in right_col}
    tail = [b for b in blocks if id(b) not in placed]
    return sorted(head, key=key) + sorted(left_col, key=key) + sorted(right_col, key=key) + sorted(tail, key=key)


def _pdf_pages_pymupdf(file_path, mupdf):
    pages = []
    with mupdf.open(file_path) as pdf:
        if getattr(pdf, "needs_pass", False):
            if not pdf.authenticate(""):
                raise PermissionError(
                    "This PDF is password-protected. Remove the password and try again."
                )
        for page in pdf:
            page_text = ""
            try:
                blocks = [
                    b for b in page.get_text("blocks")
                    if len(b) >= 5 and (len(b) < 7 or b[6] == 0) and str(b[4]).strip()
                ]
                blocks = _order_blocks(blocks, page.rect.width, page.rect.height)
                page_text = normalize_text("\n".join(str(b[4]).strip() for b in blocks))
            except Exception:
                page_text = ""
            if len(page_text) < 20:
                try:
                    plain = normalize_text(page.get_text("text"))
                    if len(plain) > len(page_text):
                        page_text = plain
                except Exception:
                    pass
            pages.append(page_text)
    return pages


def _pdf_text_fallbacks(file_path):
    """Try every secondary PDF text engine; prefer the first one with clean text."""
    results = []

    def keep(value):
        value = normalize_text(value or "")
        if value:
            results.append(value)
        return _text_looks_valid(value) and len(value) >= 100

    for module_name in ("pypdf", "PyPDF2"):
        try:
            module = __import__(module_name)
            reader = module.PdfReader(file_path)
            if getattr(reader, "is_encrypted", False):
                try:
                    reader.decrypt("")
                except Exception:
                    pass
            if keep("\n\n".join((page.extract_text() or "") for page in reader.pages)):
                return results[-1]
        except Exception:
            continue

    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(file_path) as pdf:
            if keep("\n\n".join((page.extract_text() or "") for page in pdf.pages)):
                return results[-1]
    except Exception:
        pass

    exe = shutil.which("pdftotext")
    if exe:
        try:
            result = subprocess.run(
                [exe, "-layout", file_path, "-"],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60,
                **_SUBPROCESS_KW
            )
            if keep(result.stdout.decode("utf-8", errors="replace")):
                return results[-1]
        except Exception:
            pass

    valid = [r_ for r_ in results if _text_looks_valid(r_)]
    pool = valid or results
    return max(pool, key=len) if pool else ""


# ------------------------------------------------------------
# Dependency-free PDF reader (last resort when no PDF library works)
# ------------------------------------------------------------

_STREAM_START = re.compile(rb"stream\r?\n")


def _pdf_inflate(raw):
    try:
        return zlib.decompress(raw)
    except Exception:
        pass
    try:
        return zlib.decompressobj().decompress(raw)
    except Exception:
        return b""


def _pdf_decode_chain(dict_bytes, raw):
    """Apply Flate / ASCII85 / ASCIIHex filters. Returns (data, pending_filter)."""
    if raw is None:
        return b"", None
    value = _pdf_value(dict_bytes, b"/Filter") or b""
    filters = [
        name.decode("ascii", "ignore")
        for name in re.findall(rb"/([A-Za-z0-9]+)", value)
    ]
    data = raw
    for name in filters:
        if name in ("FlateDecode", "Fl"):
            data = _pdf_inflate(data)
        elif name in ("ASCII85Decode", "A85"):
            import base64
            cleaned = data.strip()
            if cleaned.startswith(b"<~"):
                cleaned = cleaned[2:]
            if cleaned.endswith(b"~>"):
                cleaned = cleaned[:-2]
            try:
                data = base64.a85decode(cleaned)
            except Exception:
                return b"", name
        elif name in ("ASCIIHexDecode", "AHx"):
            cleaned = re.sub(rb"[^0-9A-Fa-f]", b"", data.split(b">")[0])
            if len(cleaned) % 2:
                cleaned += b"0"
            data = bytes.fromhex(cleaned.decode("ascii"))
        else:
            return data, name  # image / LZW filters are left to the caller
    return data, None


def _pdf_stream_data(dict_bytes, raw):
    data, pending = _pdf_decode_chain(dict_bytes, raw)
    return b"" if pending else data


def _pdf_objects(data):
    """Parse a PDF into {number: (dict_bytes, raw_stream_or_None)} incl. object streams."""
    objects = {}
    for match in re.finditer(rb"(?<!\d)(\d+)\s+(\d+)\s+obj\b", data):
        number = int(match.group(1))
        start = match.end()
        end_obj = data.find(b"endobj", start)
        stream = _STREAM_START.search(data, start, end_obj if end_obj >= 0 else len(data))
        if stream:
            stop = data.find(b"endstream", stream.end())
            raw = data[stream.end(): stop if stop >= 0 else len(data)]
            if raw.endswith(b"\r\n"):
                raw = raw[:-2]
            elif raw.endswith(b"\n"):
                raw = raw[:-1]
            objects[number] = (data[start:stream.start()], raw)
        else:
            objects[number] = (data[start: end_obj if end_obj >= 0 else len(data)], None)

    # Objects packed inside compressed object streams (PDF 1.5+)
    for number, (dict_bytes, raw) in list(objects.items()):
        if raw is None or b"/ObjStm" not in dict_bytes:
            continue
        body = _pdf_stream_data(dict_bytes, raw)
        count = re.search(rb"/N\s+(\d+)", dict_bytes)
        first = re.search(rb"/First\s+(\d+)", dict_bytes)
        if not body or not count or not first:
            continue
        count, first = int(count.group(1)), int(first.group(1))
        header = [int(x) for x in re.findall(rb"\d+", body[:first])][: count * 2]
        pairs = list(zip(header[0::2], header[1::2]))
        for index, (obj_number, offset) in enumerate(pairs):
            stop = pairs[index + 1][1] if index + 1 < len(pairs) else len(body) - first
            objects.setdefault(obj_number, (body[first + offset: first + stop], None))
    return objects


def _pdf_value(blob, key):
    """Return the raw value following a /Key inside a PDF dictionary."""
    if not blob:
        return None
    match = re.search(re.escape(key) + rb"(?![A-Za-z0-9])\s*", blob)
    if not match:
        return None
    i = match.end()
    if blob[i:i + 2] == b"<<":
        depth, j = 0, i
        while j < len(blob) - 1:
            pair = blob[j:j + 2]
            if pair == b"<<":
                depth += 1
                j += 2
            elif pair == b">>":
                depth -= 1
                j += 2
                if depth == 0:
                    return blob[i:j]
            else:
                j += 1
        return blob[i:]
    if blob[i:i + 1] == b"[":
        j = blob.find(b"]", i)
        return blob[i: j + 1 if j >= 0 else len(blob)]
    ref = re.match(rb"\d+\s+\d+\s+R", blob[i:i + 40])
    if ref:
        return ref.group(0)
    return blob[i:i + 60]


def _pdf_resolve(value, objects):
    if value is None:
        return None
    ref = re.match(rb"\s*(\d+)\s+\d+\s+R", value)
    if ref:
        item = objects.get(int(ref.group(1)))
        return item[0] if item else None
    return value


def _parse_tounicode(stream):
    mapping, width = {}, 1

    def utf16(hex_bytes):
        try:
            raw = bytes.fromhex(hex_bytes.decode("ascii"))
            return raw.decode("utf-16-be", errors="ignore")
        except Exception:
            return ""

    for block in re.findall(rb"beginbfchar(.*?)endbfchar", stream, re.S):
        for src, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", block):
            width = max(width, (len(src) + 1) // 2)
            mapping[int(src, 16)] = utf16(dst)
    for block in re.findall(rb"beginbfrange(.*?)endbfrange", stream, re.S):
        for m in re.finditer(
            rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])", block
        ):
            width = max(width, (len(m.group(1)) + 1) // 2)
            low, high = int(m.group(1), 16), int(m.group(2), 16)
            target = m.group(3)
            if target.startswith(b"["):
                for k, dst in enumerate(re.findall(rb"<([0-9A-Fa-f]*)>", target)):
                    mapping[low + k] = utf16(dst)
            else:
                base = utf16(target[1:-1])
                if base:
                    for k in range(min(high - low + 1, 65536)):
                        mapping[low + k] = base[:-1] + chr(ord(base[-1]) + k)
    return mapping, width


_PDF_ESCAPES = {
    ord("n"): b"\n", ord("r"): b"\r", ord("t"): b"\t", ord("b"): b"\b",
    ord("f"): b"\f", ord("("): b"(", ord(")"): b")", ord("\\"): b"\\",
}


def _pdf_literal_bytes(token):
    body, out, i = token[1:-1], bytearray(), 0
    while i < len(body):
        ch = body[i]
        if ch == 0x5C and i + 1 < len(body):  # backslash
            nxt = body[i + 1]
            if nxt in _PDF_ESCAPES:
                out += _PDF_ESCAPES[nxt]
                i += 2
            elif 0x30 <= nxt <= 0x37:
                digits = re.match(rb"[0-7]{1,3}", body[i + 1:i + 4]).group(0)
                out.append(int(digits, 8) & 0xFF)
                i += 1 + len(digits)
            elif nxt in (0x0A, 0x0D):
                i += 2
            else:
                out.append(nxt)
                i += 2
        else:
            out.append(ch)
            i += 1
    return bytes(out)


def _pdf_hex_bytes(token):
    digits = re.sub(rb"\s+", b"", token[1:-1])
    if len(digits) % 2:
        digits += b"0"
    try:
        return bytes.fromhex(digits.decode("ascii"))
    except Exception:
        return b""


def _pdf_show(raw, font):
    if font and font["map"]:
        width = font["width"]
        pieces = []
        for i in range(0, len(raw), width):
            code = int.from_bytes(raw[i:i + width], "big")
            if code in font["map"]:
                pieces.append(font["map"][code])
            elif width == 1:
                pieces.append(raw[i:i + 1].decode("cp1252", errors="ignore"))
        return "".join(pieces)
    return raw.decode("cp1252", errors="replace")


_PDF_TOKEN = re.compile(
    rb"\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]*>|/[^\s/<>\[\]()]*|\[|\]|[-+]?\d*\.?\d+|[A-Za-z'\"*]+",
    re.S,
)


def _mat_mul(m, n):
    """Row-vector matrix product m x n (PDF convention)."""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (
        a * A + b * C, a * B + b * D,
        c * A + d * C, c * B + d * D,
        e * A + f * C + E, e * B + f * D + F,
    )


_MAT_ID = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def _pdf_content_segments(content, fonts, form=None, base=_MAT_ID):
    """Interpret a content stream; returns positioned text segments (page space)."""
    import math

    segments = []          # dicts: x, y, fs, text
    ctm, ctm_stack = base, []
    tm = tlm = _MAT_ID
    lead, font, font_size = 0.0, None, 1.0
    seg = None
    operands, array = [], None

    def number(value):
        try:
            return float(value)
        except Exception:
            return 0.0

    def nums(count):
        values = [v for kind, v in operands if kind == "num"]
        return values[-count:] if len(values) >= count else None

    def new_segment():
        nonlocal seg
        m = _mat_mul(tm, ctm)
        size = abs(font_size) * (math.hypot(m[0], m[1]) or 1.0)
        seg = {"x": m[4], "y": m[5], "fs": max(size, 0.01), "text": ""}
        segments.append(seg)

    def emit(text):
        if seg is None:
            new_segment()
        seg["text"] += text

    def move(dx, dy):
        nonlocal tm, tlm
        tlm = _mat_mul((1.0, 0.0, 0.0, 1.0, dx, dy), tlm)
        tm = tlm
        new_segment()

    for token in _PDF_TOKEN.findall(content):
        head = token[:1]
        if token == b"[":
            array = []
        elif token == b"]":
            if array is not None:
                operands.append(("arr", array))
            array = None
        elif head == b"(":
            (array if array is not None else operands).append(("str", _pdf_literal_bytes(token)))
        elif head == b"<":
            (array if array is not None else operands).append(("str", _pdf_hex_bytes(token)))
        elif head == b"/":
            operands.append(("name", token[1:]))
        elif re.match(rb"[-+]?\d*\.?\d+$", token):
            (array if array is not None else operands).append(("num", number(token)))
        else:
            op = token
            if op == b"q":
                ctm_stack.append(ctm)
            elif op == b"Q":
                if ctm_stack:
                    ctm = ctm_stack.pop()
            elif op == b"cm":
                v = nums(6)
                if v:
                    ctm = _mat_mul(tuple(v), ctm)
            elif op == b"BT":
                tm = tlm = _MAT_ID
                seg = None
            elif op == b"Tf" and len(operands) >= 2 and operands[-2][0] == "name":
                font = fonts.get(operands[-2][1])
                if operands[-1][0] == "num":
                    font_size = operands[-1][1] or 1.0
            elif op == b"TL":
                v = nums(1)
                if v:
                    lead = v[0]
            elif op == b"Tm":
                v = nums(6)
                if v:
                    tm = tlm = tuple(v)
                    new_segment()
            elif op in (b"Td", b"TD"):
                v = nums(2)
                if v:
                    if op == b"TD":
                        lead = -v[1]
                    move(v[0], v[1])
            elif op == b"T*":
                move(0.0, -lead)
            elif op == b"Tj" and operands and operands[-1][0] == "str":
                emit(_pdf_show(operands[-1][1], font))
            elif op in (b"'", b'"') and operands and operands[-1][0] == "str":
                move(0.0, -lead)
                emit(_pdf_show(operands[-1][1], font))
            elif op == b"TJ" and operands and operands[-1][0] == "arr":
                for kind, value in operands[-1][1]:
                    if kind == "str":
                        emit(_pdf_show(value, font))
                    elif kind == "num" and value < -250:
                        emit(" ")
            elif op == b"ET":
                seg = None
            elif op == b"Do" and form and operands and operands[-1][0] == "name":
                segments.extend(form(operands[-1][1], ctm))
            operands = []

    return segments


def _pdf_segments_to_text(segments):
    """Order positioned segments top-to-bottom / left-to-right and join them."""
    items = [s_ for s_ in segments if s_["text"]]
    # Visual order: top to bottom (PDF y grows upwards), then left to right.
    items.sort(key=lambda s_: -s_["y"])
    rows = []
    for item in items:
        if rows and abs(rows[-1][0]["y"] - item["y"]) <= 0.5 * max(item["fs"], rows[-1][0]["fs"]):
            rows[-1].append(item)
        else:
            rows.append([item])

    lines = []
    for row in rows:
        row.sort(key=lambda s_: s_["x"])
        line, previous = "", None
        for item in row:
            text = item["text"]
            if previous is not None and not line.endswith(" ") and not text.startswith(" "):
                prev_text = previous["text"].strip()
                glyph_by_glyph = len(prev_text) <= 1 and prev_text not in ("•", "▪", "●", "-", "–", "·")
                if glyph_by_glyph:
                    gap = item["x"] - previous["x"]
                    if gap > 0.9 * item["fs"] * 1.6:
                        line += " "
                else:
                    line += " "
            line += text
            previous = item
        if line.strip():
            lines.append(line.strip())
    return "\n".join(lines)


def _pdf_page_order(objects):
    """Page object numbers in reading order (falls back to numeric order)."""
    pages = [
        n for n, (d, _) in objects.items()
        if re.search(rb"/Type\s*/Page(?![A-Za-z])", d)
    ]
    roots = [
        n for n, (d, _) in objects.items()
        if re.search(rb"/Type\s*/Pages(?![A-Za-z])", d) and b"/Parent" not in d
    ]
    ordered, seen = [], set()

    def walk(number, depth=0):
        if number in seen or depth > 30 or number not in objects:
            return
        seen.add(number)
        dict_bytes = objects[number][0]
        if re.search(rb"/Type\s*/Page(?![A-Za-z])", dict_bytes):
            ordered.append(number)
            return
        kids = _pdf_value(dict_bytes, b"/Kids")
        kids = _pdf_resolve(kids, objects) if kids and not kids.startswith(b"[") else kids
        for ref in re.findall(rb"(\d+)\s+\d+\s+R", kids or b""):
            walk(int(ref), depth + 1)

    for root in roots:
        walk(root)
    return ordered if len(ordered) >= len(pages) else sorted(pages)


def _pdf_raw_text(file_path):
    with open(file_path, "rb") as file:
        data = file.read()
    objects = _pdf_objects(data)
    if not objects:
        return ""

    font_cache = {}

    def load_font(number):
        if number not in font_cache:
            info = {"map": {}, "width": 1}
            item = objects.get(number)
            if item:
                ref = re.search(rb"/ToUnicode\s+(\d+)\s+\d+\s+R", item[0])
                if ref and int(ref.group(1)) in objects:
                    d, raw = objects[int(ref.group(1))]
                    mapping, width = _parse_tounicode(_pdf_stream_data(d, raw))
                    info = {"map": mapping, "width": width}
            font_cache[number] = info
        return font_cache[number]

    def name_refs(dict_bytes):
        return {
            name: int(number)
            for name, number in re.findall(
                rb"/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R", dict_bytes or b""
            )
        }

    def render(content, resources, depth=0, base=_MAT_ID):
        fonts = {
            name: load_font(number)
            for name, number in name_refs(
                _pdf_resolve(_pdf_value(resources, b"/Font"), objects)
            ).items()
        }
        xobjects = name_refs(_pdf_resolve(_pdf_value(resources, b"/XObject"), objects))

        def form(name, ctm):
            number = xobjects.get(name)
            item = objects.get(number) if number is not None else None
            if (
                not item or depth > 6
                or not re.search(rb"/Subtype\s*/Form", item[0])
            ):
                return []
            inner = _pdf_resolve(_pdf_value(item[0], b"/Resources"), objects) or resources
            matrix = _MAT_ID
            raw_matrix = _pdf_value(item[0], b"/Matrix")
            if raw_matrix:
                values = re.findall(rb"[-+]?\d*\.?\d+", raw_matrix)
                if len(values) >= 6:
                    matrix = tuple(float(v) for v in values[:6])
            return render(
                _pdf_stream_data(item[0], item[1]), inner, depth + 1,
                _mat_mul(matrix, ctm)
            )

        return _pdf_content_segments(content, fonts, form, base)

    pages_text = []
    for page_number in _pdf_page_order(objects):
        page = objects[page_number][0]

        node, resources, hops = page, None, 0
        while node is not None and hops < 12 and resources is None:
            value = _pdf_value(node, b"/Resources")
            if value:
                resources = _pdf_resolve(value, objects)
                break
            parent = _pdf_value(node, b"/Parent")
            ref = re.match(rb"(\d+)\s+\d+\s+R", parent or b"")
            node = objects[int(ref.group(1))][0] if ref and int(ref.group(1)) in objects else None
            hops += 1

        contents = _pdf_value(page, b"/Contents")
        refs = [int(x) for x in re.findall(rb"(\d+)\s+\d+\s+R", contents or b"")]
        chunks = []
        for number in refs:
            item = objects.get(number)
            if not item:
                continue
            if item[1] is None and item[0].strip().startswith(b"["):
                for inner in re.findall(rb"(\d+)\s+\d+\s+R", item[0]):
                    sub = objects.get(int(inner))
                    if sub:
                        chunks.append(_pdf_stream_data(sub[0], sub[1]))
            else:
                chunks.append(_pdf_stream_data(item[0], item[1]))
        text = re.sub(
            r"[ \t]*\n[\s]*", "\n",
            _pdf_segments_to_text(render(b"\n".join(chunks), resources))
        )
        if text.strip():
            pages_text.append(text)
    return normalize_text("\n\n".join(pages_text))


def _pdf_raw_images(file_path, max_images=15):
    """Pull embedded page images (scans) out of a PDF without any PDF library."""
    if _ensure_pillow() is None:
        return []
    import io

    with open(file_path, "rb") as file:
        data = file.read()
    images = []
    objects = _pdf_objects(data)
    for number in sorted(objects):
        dict_bytes, raw = objects[number]
        if raw is None or not re.search(rb"/Subtype\s*/Image", dict_bytes):
            continue
        try:
            picture = None
            data, pending = _pdf_decode_chain(dict_bytes, raw)
            if pending in ("DCTDecode", "DCT", "JPXDecode"):
                picture = Image.open(io.BytesIO(data))
                picture.load()
            elif pending is None and data and b"/Predictor" not in dict_bytes:
                width = int(re.search(rb"/Width\s+(\d+)", dict_bytes).group(1))
                height = int(re.search(rb"/Height\s+(\d+)", dict_bytes).group(1))
                bits = int(re.search(rb"/BitsPerComponent\s+(\d+)", dict_bytes).group(1))
                if bits == 1 and b"DeviceGray" in dict_bytes:
                    picture = Image.frombytes("1", (width, height), data)
                elif bits == 8 and b"DeviceGray" in dict_bytes:
                    picture = Image.frombytes("L", (width, height), data)
                elif bits == 8 and b"DeviceRGB" in dict_bytes:
                    picture = Image.frombytes("RGB", (width, height), data)
            if picture is not None and min(picture.size) >= 200:
                images.append(picture.convert("RGB"))
        except Exception:
            continue
        if len(images) >= max_images:
            break
    return images


def _pdf_render_pages(file_path, max_pages=15):
    """Render PDF pages to images using the first rendering backend that works."""
    if _ensure_pillow() is None:
        return []

    mupdf = _pymupdf()
    if mupdf is not None:
        try:
            images = []
            with mupdf.open(file_path) as pdf:
                for index, page in enumerate(pdf):
                    if index >= max_pages:
                        break
                    pix = page.get_pixmap(matrix=mupdf.Matrix(2.5, 2.5), alpha=False)
                    images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
            if images:
                return images
        except Exception:
            pass

    pdfium = _import_or_install("pypdfium2")
    if pdfium is not None:
        try:
            document = pdfium.PdfDocument(file_path)
            images = []
            for index in range(min(len(document), max_pages)):
                images.append(document[index].render(scale=2.5).to_pil().convert("RGB"))
            if images:
                return images
        except Exception:
            pass

    exe = shutil.which("pdftoppm")
    if exe:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                subprocess.run(
                    [exe, "-r", "200", "-l", str(max_pages), "-png", file_path,
                     os.path.join(temp_dir, "page")],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
                    **_SUBPROCESS_KW
                )
                images = []
                for name in sorted(os.listdir(temp_dir)):
                    with Image.open(os.path.join(temp_dir, name)) as page_image:
                        images.append(page_image.convert("RGB"))
                if images:
                    return images
        except Exception:
            pass

    return _pdf_raw_images(file_path, max_pages)


def extract_pdf_text(file_path):
    """
    PDF reader. Tries, in order, and keeps the best result:
      PyMuPDF -> pypdf / PyPDF2 / pdfplumber / pdftotext -> built-in parser
      -> OCR of the rendered pages (scanned or image-only PDFs).
    Missing libraries are installed automatically when possible.
    """
    candidates = []

    def consider(value):
        value = normalize_text(value or "")
        if value:
            candidates.append(value)
        return value

    def best():
        valid = [c for c in candidates if _text_looks_valid(c)]
        return max(valid or candidates, key=len) if (valid or candidates) else ""

    mupdf = _pymupdf()
    if mupdf is not None:
        try:
            consider("\n\n".join(_pdf_pages_pymupdf(file_path, mupdf)))
        except PermissionError:
            raise
        except Exception:
            pass

    def good_enough():
        text = best()
        return len(text) >= 100 and _text_looks_valid(text)

    if not good_enough():
        _import_or_install("pypdf")
        consider(_pdf_text_fallbacks(file_path))

    if not good_enough():
        try:
            consider(_pdf_raw_text(file_path))
        except Exception:
            pass

    ocr_error = None
    if not good_enough():
        try:
            pages = _pdf_render_pages(file_path)
            ocr_pages = []
            for page_image in pages:
                try:
                    ocr_pages.append(ocr_image(page_image))
                except RuntimeError as error:
                    ocr_error = error
                    break
            consider("\n\n".join(p for p in ocr_pages if p))
        except Exception as error:
            ocr_error = error

    text = best()
    if text:
        return text

    detail = f"\n\n{ocr_error}" if ocr_error else ""
    raise ValueError(
        "No readable text could be extracted from this PDF. It may be a scan "
        "or password-protected." + detail
    )


# ------------------------------------------------------------
# DOCX / PPTX / ODT (XML based)
# ------------------------------------------------------------

def _docx_xml_text(file_path):
    """Read a DOCX straight from its XML (includes text boxes, headers, footers)."""
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    def collect(node, buffer):
        for child in node:
            if child.tag == W + "p":
                continue  # nested paragraphs are visited separately
            if child.tag == W + "t" and child.text:
                buffer.append(child.text)
            elif child.tag == W + "tab":
                buffer.append(" ")
            elif child.tag in (W + "br", W + "cr"):
                buffer.append("\n")
            collect(child, buffer)

    parts = []
    with zipfile.ZipFile(file_path) as archive:
        names = [
            n for n in archive.namelist()
            if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)
        ]
        names.sort(key=lambda n: 0 if n.endswith("document.xml") else 1)
        for name in names:
            root = ET.fromstring(archive.read(name))
            for parent in list(root.iter()):
                for child in list(parent):
                    if child.tag.endswith("}Fallback"):
                        parent.remove(child)
            for paragraph in root.iter(W + "p"):
                buffer = []
                collect(paragraph, buffer)
                line = "".join(buffer).strip()
                if line:
                    parts.append(line)
    return normalize_text("\n".join(parts))


def extract_docx_text(file_path):
    """DOCX reader: python-docx, with a direct-XML fallback."""
    docx_text = ""
    if Document is not None:
        try:
            document = Document(file_path)
            parts = []
            for paragraph in document.paragraphs:
                value = paragraph.text.strip()
                if value:
                    parts.append(value)
            # Many resumes store contact details or skills in tables.
            for table in document.tables:
                for row in table.rows:
                    cells = []
                    for cell in row.cells:
                        value = normalize_text(cell.text)
                        if value and value not in cells:
                            cells.append(value)
                    if cells:
                        parts.append(" | ".join(cells))
            docx_text = normalize_text("\n".join(parts))
        except Exception:
            docx_text = ""

    try:
        xml_text = _docx_xml_text(file_path)
    except Exception:
        xml_text = ""

    # Prefer python-docx output (keeps table structure) unless it missed
    # a large part of the content (e.g. text boxes).
    if len(docx_text) >= 0.5 * len(xml_text):
        xml_text = ""

    text = xml_text if len(xml_text) > len(docx_text) else docx_text
    if not text:
        raise ValueError("The DOCX file was opened, but no readable text was found.")
    return text


def extract_odt_xml_text(file_path):
    with zipfile.ZipFile(file_path) as archive:
        root = ET.fromstring(archive.read("content.xml"))
    lines = []
    for element in root.iter():
        if element.tag.endswith("}p") or element.tag.endswith("}h"):
            line = "".join(element.itertext()).strip()
            if line:
                lines.append(line)
    return normalize_text("\n".join(lines))


def extract_pptx_text(file_path):
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    lines = []
    with zipfile.ZipFile(file_path) as archive:
        slides = sorted(
            (n for n in archive.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1))
        )
        for name in slides:
            root = ET.fromstring(archive.read(name))
            for paragraph in root.iter(A + "p"):
                line = "".join(t.text or "" for t in paragraph.iter(A + "t")).strip()
                if line:
                    lines.append(line)
    text = normalize_text("\n".join(lines))
    if not text:
        raise ValueError("No readable text was found in the presentation.")
    return text


# ------------------------------------------------------------
# LibreOffice / legacy formats
# ------------------------------------------------------------

def _find_office_converter():
    """Find LibreOffice/soffice for legacy Office/RTF conversion."""
    for executable in ("soffice", "libreoffice"):
        path = shutil.which(executable)
        if path:
            return path

    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def extract_via_libreoffice(file_path, target="docx"):
    """Convert DOC/RTF/ODT (and friends) to DOCX (or CSV/TXT) with LibreOffice, then read it."""
    office = _find_office_converter()
    if not office:
        raise RuntimeError("LibreOffice is not installed.")

    with tempfile.TemporaryDirectory() as temp_dir:
        profile = Path(temp_dir) / "lo_profile"
        out_dir = Path(temp_dir) / "out"
        out_dir.mkdir()
        command = [
            office,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--convert-to", target,
            "--outdir", str(out_dir),
            file_path,
        ]
        result = subprocess.run(
            command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=180, **_SUBPROCESS_KW
        )
        converted = list(out_dir.glob("*." + target.split(":")[0]))
        if result.returncode != 0 or not converted:
            details = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                "LibreOffice could not convert the document."
                + (f"\n\nDetails: {details}" if details else "")
            )
        if target == "docx":
            return extract_docx_text(str(converted[0]))
        if target == "pptx":
            return extract_pptx_text(str(converted[0]))
        return normalize_text(_decode_bytes(converted[0].read_bytes()))


def _rtf_to_text(raw):
    """Small built-in RTF-to-text converter (used when LibreOffice is missing)."""
    skip_words = {
        "fonttbl", "colortbl", "stylesheet", "info", "pict", "header", "footer",
        "themedata", "colorschememapping", "latentstyles", "datastore",
        "generator", "listtable", "listoverridetable", "rsidtbl", "xmlnstbl",
    }
    out, stack, skip = [], [], False
    i, n = 0, len(raw)
    while i < n:
        ch = raw[i]
        if ch == "{":
            stack.append(skip)
            i += 1
        elif ch == "}":
            skip = stack.pop() if stack else False
            i += 1
        elif ch == "\\":
            i += 1
            if i >= n:
                break
            c = raw[i]
            if c in "\\{}":
                if not skip:
                    out.append(c)
                i += 1
            elif c == "'":
                try:
                    value = bytes([int(raw[i + 1:i + 3], 16)]).decode("cp1252", errors="replace")
                    if not skip:
                        out.append(value)
                except ValueError:
                    pass
                i += 3
            elif c == "*":
                skip = True
                i += 1
            elif c == "~":
                out.append(" ")
                i += 1
            elif c in "\r\n":
                out.append("\n")
                i += 1
            else:
                match = re.match(r"([a-zA-Z]+)(-?\d+)? ?", raw[i:])
                if not match:
                    i += 1
                    continue
                word, param = match.group(1), match.group(2)
                i += match.end()
                if word in skip_words:
                    skip = True
                elif skip:
                    continue
                elif word in ("par", "line", "sect", "page", "row"):
                    out.append("\n")
                elif word == "tab":
                    out.append(" ")
                elif word == "cell":
                    out.append(" | ")
                elif word == "u" and param:
                    code = int(param)
                    out.append(chr(code + 65536 if code < 0 else code))
                    if raw[i:i + 2] == "\\'":
                        i += 4
                    elif i < n:
                        i += 1
        else:
            if ch not in "\r\n" and not skip:
                out.append(ch)
            i += 1
    return normalize_text("".join(out))


def extract_rtf_text(file_path):
    try:
        return extract_via_libreoffice(file_path)
    except Exception:
        pass
    with open(file_path, "rb") as file:
        text = _rtf_to_text(file.read().decode("latin-1"))
    if not text:
        raise ValueError("The RTF file contains no readable text.")
    return text


def extract_odt_text(file_path):
    try:
        text = extract_odt_xml_text(file_path)
        if text:
            return text
    except Exception:
        pass
    return extract_via_libreoffice(file_path)


def extract_doc_text(file_path):
    """Legacy .doc: LibreOffice -> antiword/catdoc -> raw text extraction."""
    try:
        return extract_via_libreoffice(file_path)
    except Exception:
        pass

    for tool in ("antiword", "catdoc"):
        exe = shutil.which(tool)
        if exe:
            try:
                result = subprocess.run(
                    [exe, file_path], stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE, timeout=60
                )
                text = normalize_text(result.stdout.decode("utf-8", errors="replace"))
                if text:
                    return text
            except Exception:
                pass

    text = _binary_strings(file_path)
    if len(text) < 30:
        raise RuntimeError(
            "This legacy Word file could not be read. Install LibreOffice "
            "or save the file as DOCX/PDF and try again."
        )
    return text


# ------------------------------------------------------------
# TXT / HTML / images
# ------------------------------------------------------------

def extract_txt_text(file_path):
    """Read plain text / HTML / Markdown / CSV / JSON / XML in any common encoding."""
    with open(file_path, "rb") as file:
        data = file.read()
    raw = _decode_bytes(data)
    if re.search(r"(?i)<\s*(html|body|div|p|table)\b", raw[:5000]) or re.match(
        r"(?is)\s*(<\?xml|<svg)", raw[:200]
    ):
        raw = _strip_html(raw)
    text = normalize_text(raw)
    if not text:
        raise ValueError("The file could be opened, but it contains no readable text.")
    return text


def extract_image_text(file_path):
    """OCR PNG/JPG/BMP/TIFF/WebP/GIF resume images (all frames of multi-page TIFFs)."""
    if _ensure_pillow() is None:
        raise RuntimeError(
            "Image support requires Pillow. Install it with:\n"
            "pip install pillow"
        )

    if Path(file_path).suffix.lower() in (".heic", ".heif"):
        heif = _import_or_install("pillow_heif", "pillow-heif")
        if heif is not None:
            try:
                heif.register_heif_opener()
            except Exception:
                pass

    texts = []
    with Image.open(file_path) as image:
        frames = getattr(image, "n_frames", 1)
        for index in range(min(frames, 20)):
            try:
                image.seek(index)
            except EOFError:
                break
            frame = image.convert("RGB")
            if max(frame.size) < 1500:
                scale = 1500 / max(frame.size)
                frame = frame.resize(
                    (int(frame.width * scale), int(frame.height * scale)),
                    Image.LANCZOS
                )
            texts.append(ocr_image(frame))

    text = normalize_text("\n\n".join(t for t in texts if t))
    if not text:
        raise ValueError("OCR ran on the image but found no readable text.")
    return text


# ------------------------------------------------------------
# Spreadsheets, e-books, other ZIP-based and legacy binary files
# ------------------------------------------------------------

def extract_xlsx_text(file_path):
    """Read text from XLSX/XLSM (shared strings + inline strings + numbers)."""
    lines = []
    with zipfile.ZipFile(file_path) as archive:
        names = archive.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.iter():
                if item.tag.endswith("}si"):
                    shared.append("".join(t.text or "" for t in item.iter() if t.tag.endswith("}t")))
        sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        for name in sheets:
            root = ET.fromstring(archive.read(name))
            for row in root.iter():
                if not row.tag.endswith("}row"):
                    continue
                cells = []
                for cell in row:
                    kind = cell.get("t")
                    value = ""
                    for child in cell:
                        if child.tag.endswith("}v") and child.text:
                            value = child.text
                            if kind == "s" and value.isdigit() and int(value) < len(shared):
                                value = shared[int(value)]
                        elif child.tag.endswith("}is"):
                            value = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
                    if value.strip():
                        cells.append(value.strip())
                if cells:
                    lines.append(" | ".join(cells))
    text = normalize_text("\n".join(lines))
    if not text:
        raise ValueError("No readable text was found in the spreadsheet.")
    return text


def extract_zip_text(file_path):
    """Generic ZIP container (EPUB, Pages export, unknown packages): read its text/HTML/XML parts."""
    parts = []
    with zipfile.ZipFile(file_path) as archive:
        for info in archive.infolist():
            lower = info.filename.lower()
            if info.is_dir() or info.file_size > 20_000_000:
                continue
            if lower.endswith((".html", ".xhtml", ".htm")):
                parts.append(_strip_html(_decode_bytes(archive.read(info))))
            elif lower.endswith((".txt", ".md", ".rtf")):
                parts.append(_decode_bytes(archive.read(info)))
            elif lower.endswith(".xml") and "content" in lower:
                try:
                    root = ET.fromstring(archive.read(info))
                    parts.append("\n".join(t.strip() for t in root.itertext() if t.strip()))
                except Exception:
                    pass
    text = normalize_text("\n".join(parts))
    if not text:
        raise ValueError("No readable text was found inside this archive.")
    return text


def extract_binary_text(file_path):
    """Legacy binary files (WPS, unknown): recover readable text runs."""
    text = _binary_strings(file_path)
    lines = [
        line for line in text.splitlines()
        if re.search(r"[A-Za-z]{3,}", line)
        and sum(ch.isascii() for ch in line) >= 0.8 * len(line)
    ]
    text = normalize_text("\n".join(lines))
    if len(re.findall(r"[A-Za-z]{3,}", text)) < 15:
        raise ValueError(
            "This file type could not be read. Save the CV as PDF, DOCX or TXT and try again."
        )
    return text


def extract_xls_text(file_path):
    """Legacy XLS: LibreOffice (to CSV) -> raw text recovery."""
    try:
        text = extract_via_libreoffice(file_path, "csv")
        if text.strip():
            return text
    except Exception:
        pass
    return extract_binary_text(file_path)


def extract_ppt_text(file_path):
    """Legacy PPT: LibreOffice (to PPTX) -> raw text recovery."""
    try:
        return extract_via_libreoffice(file_path, "pptx")
    except Exception:
        pass
    return extract_binary_text(file_path)


# ------------------------------------------------------------
# FILE TYPE DETECTION + UNIVERSAL READER
# ------------------------------------------------------------

_EXT_KIND = {
    ".pdf": "pdf", ".ai": "pdf",
    ".docx": "docx", ".docm": "docx", ".dotx": "docx", ".dotm": "docx",
    ".doc": "doc", ".dot": "doc", ".wps": "doc",
    ".rtf": "rtf",
    ".odt": "odt", ".ott": "odt", ".fodt": "text",
    ".pptx": "pptx", ".pptm": "pptx", ".ppt": "ppt", ".pps": "ppt",
    ".xlsx": "xlsx", ".xlsm": "xlsx", ".xltx": "xlsx", ".xls": "xls", ".xlt": "xls",
    ".epub": "zip", ".zip": "zip", ".pages": "zip",
    ".txt": "text", ".text": "text", ".md": "text", ".csv": "text", ".tsv": "text",
    ".json": "text", ".xml": "text", ".html": "text", ".htm": "text",
    ".xhtml": "text", ".svg": "text", ".log": "text", ".rst": "text",
    ".png": "image", ".jpg": "image", ".jpeg": "image", ".jpe": "image",
    ".jfif": "image", ".bmp": "image", ".tif": "image", ".tiff": "image",
    ".webp": "image", ".gif": "image", ".heic": "image", ".heif": "image",
    ".ppm": "image", ".pgm": "image", ".jp2": "image",
}


def _detect_kind(file_path):
    """Decide the real file type from its content first, then its extension."""
    extension = Path(file_path).suffix.lower()
    try:
        with open(file_path, "rb") as file:
            head = file.read(2048)
    except OSError:
        head = b""

    stripped = head.lstrip(b"\xef\xbb\xbf \t\r\n")

    if b"%PDF" in head[:1024]:
        return "pdf"
    if stripped.startswith(b"{\\rtf"):
        return "rtf"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        # Old OLE container: Word, Excel or PowerPoint
        if extension in (".xls", ".xlt"):
            return "xls"
        if extension in (".ppt", ".pps"):
            return "ppt"
        return "doc"
    if (
        head.startswith(b"\x89PNG")
        or head.startswith(b"\xff\xd8")
        or head.startswith(b"GIF8")
        or (head.startswith(b"BM") and extension == ".bmp")
        or head[:4] in (b"II*\x00", b"MM\x00*")
        or (head[:4] == b"RIFF" and head[8:12] == b"WEBP")
        or (head[4:8] == b"ftyp" and head[8:12] in (b"heic", b"heix", b"mif1", b"msf1", b"hevc"))
    ):
        return "image"
    if head.startswith(b"PK"):
        try:
            with zipfile.ZipFile(file_path) as archive:
                names = set(archive.namelist())
        except Exception:
            names = set()
        if any(n.startswith("word/") for n in names):
            return "docx"
        if any(n.startswith("xl/") for n in names):
            return "xlsx"
        if any(n.startswith("ppt/slides/") for n in names):
            return "pptx"
        if "content.xml" in names:
            return "odt"
        if names:
            return "zip"

    if extension in _EXT_KIND:
        guess = _EXT_KIND[extension]
        # A ".pdf" without any PDF signature that is plain text is really text.
        if guess == "pdf" and b"%PDF" not in head and head and b"\x00" not in head:
            return "text"
        return guess

    # Unknown extension: text when it decodes cleanly, otherwise binary recovery.
    if head[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" not in head:
        return "text"
    return None


def _image_openable(file_path):
    if _ensure_pillow() is None:
        return False
    try:
        with Image.open(file_path) as picture:
            picture.verify()
        return True
    except Exception:
        return False


def extract_file_text(file_path):
    """
    Universal resume reader. The real file type is detected from the file
    content, so wrong or missing extensions do not matter.

      PDF   -> PyMuPDF / pypdf / pdfplumber / pdftotext / built-in parser,
               OCR for scanned or image-only PDFs
      DOCX  -> python-docx or direct XML (text boxes, headers, footers)
      DOC / RTF / ODT -> LibreOffice, or built-in / antiword / catdoc fallbacks
      PPTX, XLSX, EPUB, ZIP, TXT, MD, HTML, CSV, JSON, XML -> text extraction
      Images (PNG, JPG, BMP, TIFF, WebP, GIF, HEIC) -> OCR
      Anything else -> best-effort text recovery

    Missing helper libraries are installed automatically when possible.
    """
    file_path = os.fspath(file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    if os.path.getsize(file_path) == 0:
        raise ValueError("The file is empty.")

    kind = _detect_kind(file_path)
    readers = {
        "pdf": extract_pdf_text,
        "docx": extract_docx_text,
        "doc": extract_doc_text,
        "rtf": extract_rtf_text,
        "odt": extract_odt_text,
        "pptx": extract_pptx_text,
        "xlsx": extract_xlsx_text,
        "zip": extract_zip_text,
        "xls": extract_xls_text,
        "ppt": extract_ppt_text,
        "binary": extract_binary_text,
        "text": extract_txt_text,
        "image": extract_image_text,
    }

    def usable(value):
        return bool(value and value.strip())

    primary_error = None
    if kind is not None:
        try:
            text = readers[kind](file_path)
            if usable(text):
                return text
            primary_error = ValueError("No readable text was found in this file.")
        except PermissionError:
            raise
        except Exception as error:
            primary_error = error

    # Fallbacks for mislabeled, damaged or unusual files.
    attempts = []
    if kind not in ("text", None):
        attempts.append(extract_txt_text if kind in ("docx", "odt", "pptx", "rtf", "zip", "xlsx") else None)
    if kind != "image":
        attempts.append(lambda p: extract_image_text(p) if _image_openable(p) else "")
    if kind not in ("pdf", "image", "binary", "xls", "ppt", "docx", "xlsx", "pptx", "odt", "zip"):
        attempts.append(extract_binary_text)
    if kind is None:
        attempts.append(extract_binary_text)

    for attempt in attempts:
        if attempt is None:
            continue
        try:
            text = attempt(file_path)
            if usable(text):
                return text
        except PermissionError:
            raise
        except Exception:
            continue

    if primary_error is not None:
        raise primary_error
    raise ValueError(
        f"Unsupported or unreadable file type: {Path(file_path).suffix or 'unknown'}\n\n"
        "Try saving the CV as PDF, DOCX or TXT."
    )


# ============================================================
# CONTACT INFORMATION
# ============================================================

EMAIL_PATTERN = re.compile(
    r"\b[A-Za-z0-9._%+-]+"
    r"@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
)


PHONE_PATTERN = re.compile(
    r"(?<!\d)"
    r"(?:\+?\d{1,3}[\s.-]?)?"
    r"(?:\(?\d{2,5}\)?[\s.-]?)?"
    r"\d{3,5}[\s.-]?\d{4}"
    r"(?!\d)"
)


def extract_email(text):

    match = EMAIL_PATTERN.search(
        text
    )

    return (
        match.group(0)
        if match
        else ""
    )


def extract_phone(text):

    matches = PHONE_PATTERN.findall(
        text
    )

    for match in matches:

        digits = re.sub(
            r"\D",
            "",
            match
        )

        if 10 <= len(digits) <= 15:
            return match.strip()

    return ""


# ============================================================
# NAME DETECTION
# ============================================================

IGNORED_NAME_LINES = {
    "resume",
    "curriculum vitae",
    "cv",
    "profile",
    "professional profile",
    "summary",
    "professional summary",
    "objective",
    "education",
    "experience",
    "skills",
    "projects",
    "certifications",
    "courses",
    "contact",
    "languages"
}


NON_NAME_WORDS = {
    # headings / labels
    "summary", "qualification", "qualifications", "education", "educational",
    "experience", "experiences", "skills", "skill", "objective", "profile",
    "achievement", "achievements", "achievements", "leadership", "internship",
    "internships", "training", "professional", "certification", "certifications",
    "courses", "course", "programme", "program", "languages", "language",
    "interests", "interest", "hobbies", "projects", "project", "references",
    "declaration", "contact", "details", "detail", "information", "address",
    "personal", "academic", "exposure", "industry", "curriculum", "vitae",
    "resume", "cv", "roll", "number", "work", "history", "career", "overview",
    "technical", "core", "key", "strengths", "responsibilities", "awards",
    "activities", "extracurricular", "publications", "board", "passing", "year",
    # organisations / places
    "university", "college", "school", "institute", "institution", "academy",
    "ltd", "limited", "pvt", "private", "inc", "llp", "corp", "company",
    "technologies", "solutions", "services", "developers", "media", "bank",
    "department", "faculty", "campus",
}


def looks_like_name(line):

    line = line.strip()

    if not line:
        return False

    if len(line) > 70:
        return False

    if "@" in line:
        return False

    if re.search(
        r"\d",
        line
    ):
        return False

    if line.lower() in IGNORED_NAME_LINES:
        return False

    # Section headings, labels and organisation names are never a person's name.
    if any(
        re.sub(r"[^a-z]", "", word.lower()) in NON_NAME_WORDS
        for word in line.split()
    ):
        return False

    words = line.split()

    if not 2 <= len(words) <= 5:
        return False

    for word in words:

        if not re.match(
            r"^[A-Za-zÀ-ÖØ-öø-ÿ'.-]+$",
            word
        ):
            return False

    return True


def extract_name(text):

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # Explicit name field
    for line in lines[:30]:

        match = re.match(
            r"^(?:name|candidate name|full name)"
            r"\s*[:\-]\s*(.+)$",
            line,
            re.IGNORECASE
        )

        if match:

            candidate = match.group(1).strip()

            if looks_like_name(candidate):
                return candidate

    # A line starting with Mr./Ms./Dr. ... is a very reliable name marker, even
    # when the PDF layout (side columns) pushes it far down the extracted text.
    for line in lines[:80]:

        match = re.match(
            r"^(?:mr|mrs|ms|miss|mx|dr|shri|smt|sri|prof)\.?\s+(.+)$",
            line,
            re.IGNORECASE
        )

        if match:

            candidate = match.group(1).strip()

            if looks_like_name(candidate):
                return candidate

    # First plausible name
    for line in lines[:15]:

        if looks_like_name(line):
            return line

    return ""


# ============================================================
# SKILL EXTRACTION
# ============================================================

def contains_skill(text, skill):

    escaped = re.escape(
        skill
    )

    escaped = escaped.replace(
        r"\ ",
        r"\s+"
    )

    pattern = (
        r"(?<![a-zA-Z0-9])"
        + escaped
        + r"(?![a-zA-Z0-9])"
    )

    return bool(
        re.search(
            pattern,
            text,
            re.IGNORECASE
        )
    )


def extract_skills_by_category(text):

    result = {}

    for category, skills in SKILLS.items():

        found = []

        for skill in skills:

            if contains_skill(
                text,
                skill
            ):

                found.append(
                    display_skill(skill)
                )

        if found:

            result[category] = sorted(
                set(found)
            )

    return result


# ============================================================
# SECTION DETECTION
# ============================================================

SECTION_ALIASES = {

    "Education": [
        "education",
        "academic background",
        "educational qualification",
        "qualifications"
    ],

    "Experience": [
        "experience",
        "work experience",
        "professional experience",
        "employment",
        "career history",
        "internship",
        "internships",
        "internship experience",
        "work history",
        "employment history",
        "professional background",
        "relevant experience"
    ],

    "Skills": [
        "skills",
        "technical skills",
        "core skills",
        "key skills",
        "competencies"
    ],

    "Certifications": [
        "certifications",
        "certificates",
        "courses",
        "courses and certifications",
        "certification"
    ],

    "Projects": [
        "projects",
        "academic projects",
        "personal projects"
    ],

    "Languages": [
        "languages",
        "language"
    ],

    "Summary": [
        "summary",
        "professional summary",
        "profile",
        "career objective",
        "objective"
    ]
}


def normalize_heading(line):

    clean = line.strip().lower()

    clean = re.sub(
        r"[^a-z0-9& ]",
        "",
        clean
    )

    clean = re.sub(
        r"\s+",
        " ",
        clean
    )

    return clean.strip()


# Headings are recognised by meaning, not only by an exact list of titles, so
# "INTERNSHIP / PROFESSIONAL TRAINING (Total Duration: 16 Weeks)" counts as
# Experience and "ACHIEVEMENT DETAIL" correctly ends it.
_HEADING_RULES = (
    ("Experience", {
        "words": {
            "work", "professional", "relevant", "industrial", "practical", "summer",
            "winter", "internship", "internships", "experience", "experiences",
            "training", "trainings", "employment", "history", "details", "detail",
            "and", "of", "record", "apprenticeship", "traineeship", "field",
            "corporate", "industry", "on", "the", "job", "professionals",
        },
        "keys": {
            "experience", "experiences", "internship", "internships", "employment",
            "training", "trainings", "apprenticeship", "traineeship",
        },
    }),
    ("Education", {
        "words": {
            "summary", "of", "educational", "education", "qualification",
            "qualifications", "academic", "academics", "background", "details",
            "detail", "record", "and", "profile",
        },
        "keys": {"education", "educational", "qualification", "qualifications", "academics"},
    }),
    ("Other", {
        "words": {
            "achievement", "achievements", "detail", "details", "leadership", "awards",
            "honours", "honors", "accomplishments", "interests", "interest", "hobbies",
            "extracurricular", "activities", "declaration", "references", "personal",
            "information", "permanent", "address", "academic", "industry", "exposure",
            "and", "positions", "responsibility", "volunteer", "volunteering",
            "social", "contact", "strengths", "co", "curricular", "of",
        },
        "keys": {
            "achievement", "achievements", "accomplishments", "awards", "honours",
            "honors", "interests", "interest", "hobbies", "extracurricular",
            "declaration", "references", "address", "exposure", "volunteer",
            "volunteering", "personal", "contact", "strengths", "curricular",
        },
    }),
)


def classify_heading(line):
    """Return the section a heading line starts, or None if it is not a heading."""
    clean = normalize_heading(line)
    for section, aliases in SECTION_ALIASES.items():
        if clean in aliases:
            return section

    raw = line.strip()
    if not raw or len(raw) > 80 or len(raw.split()) > 9 or raw.endswith("."):
        return None
    base = re.sub(r"\([^)]*\)", " ", raw)          # drop "(Total Duration: 16 Weeks)"
    if re.search(r"\d", base):
        return None
    words = re.findall(r"[a-z]+", base.lower())
    if not words:
        return None
    for section, rule in _HEADING_RULES:
        if all(w in rule["words"] for w in words) and any(w in rule["keys"] for w in words):
            return section
    return None


def detect_sections(text):

    lines = text.splitlines()

    sections = {}

    current_section = "General"

    sections[current_section] = []

    for line in lines:

        detected = classify_heading(line)

        if detected:

            current_section = detected

            if current_section not in sections:
                sections[current_section] = []

            continue

        sections.setdefault(
            current_section,
            []
        ).append(
            line
        )

    final_sections = {}

    for section, content in sections.items():

        value = normalize_text(
            "\n".join(content)
        )

        if value:
            final_sections[section] = value

    return final_sections


# ============================================================
# EDUCATION EXTRACTION
# ============================================================

def extract_education(text):

    sections = detect_sections(
        text
    )

    education_text = sections.get(
        "Education",
        ""
    )

    if education_text:
        return education_text

    # Fallback: search lines containing education keywords
    lines = text.splitlines()

    matches = []

    education_keywords = [
        "bachelor",
        "master",
        "mba",
        "bba",
        "b.com",
        "m.com",
        "b.tech",
        "m.tech",
        "degree",
        "university",
        "college",
        "school",
        "cgpa",
        "gpa",
        "diploma"
    ]

    for line in lines:

        lower = line.lower()

        if any(
            keyword in lower
            for keyword in education_keywords
        ):

            matches.append(
                line.strip()
            )

    return "\n".join(
        dict.fromkeys(matches)
    )


# ============================================================
# EXPERIENCE EXTRACTION
# ============================================================

_MONTH_RE = (
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|"
    r"Dec(?:ember)?)"
)
_DATE_TOKEN_RE = (
    r"(?:\d{1,2}[-/. ]\s*" + _MONTH_RE + r"\.?[-/. ,]*\s*"      # 12-May-2025
    r"|" + _MONTH_RE + r"\.?[-/. ,]*\s*"                          # May 2025 / May-2025
    r"|\d{1,2}[-/.])?\d{4}"                                        # 05/2025 / 2025
)
DATE_RANGE_RE = re.compile(
    r"(?<![\d/.-])" + _DATE_TOKEN_RE + r"\s*(?:-|–|—|to)\s*"
    r"(?:Present|Current|Till\s+date|Now|Ongoing|" + _DATE_TOKEN_RE + r")",
    re.IGNORECASE,
)

DURATION_RE = re.compile(
    r"^\(?\s*\d+\s*(?:weeks?|wks?|months?|mos?|days?|years?|yrs?)\s*\)?$",
    re.IGNORECASE,
)

ROLE_WORDS = (
    "manager", "analyst", "associate", "intern", "consultant",
    "executive", "developer", "engineer", "specialist", "coordinator",
    "lead", "director", "assistant", "advisor", "trainee", "buyer",
    "mentor", "officer", "administrator", "designer", "relations"
)

ACTION_WORDS = (
    "led ", "leading ", "acting ", "owning ", "driving ", "developed ",
    "regularly analysed ", "regularly analyzed ", "analysed ", "analyzed ",
    "brainstormed ", "maintained ", "managed ", "implemented ",
    "collaborated ", "communicated ", "utilized ", "assisted ",
    "handled ", "scaled ", "planned ", "executed ", "learned ",
    "conducted ", "supported ", "contributed ", "gained ", "worked ", "created ",
    "designed ", "built ", "prepared ", "organized ", "organised ", "coordinated ",
    "performed ", "participated ", "helped ", "generated ", "delivered ",
    "researched ", "trained ", "monitored ", "tracked ", "optimized ", "optimised ",
    "launched ", "supervised ", "mentored ", "established ", "improved ",
    "increased ", "reduced ", "identified ", "achieved ", "presented ",
    "evaluated ", "studied ", "compiled ", "updated ", "ensured ", "responsible ",
    "interacted ", "negotiated ", "managed ", "wrote ", "ran "
)


def _clean_experience_line(line):
    line = normalize_text(line).strip(" |•-\t")
    return re.sub(r"\s+", " ", line).strip()


def _shorten_experience_point(text, max_words=14):
    """Create a short, grammatical experience bullet from a source bullet."""
    text = _clean_experience_line(text)
    text = text.replace("fullfunnel", "full-funnel")
    text = re.sub(r"\s*;\s*", "; ", text)

    words = text.split()
    if len(words) <= max_words:
        return text.rstrip(".") + "."

    # Prefer a concise measurable achievement when two values are present,
    # e.g. "scaled ... from ₹70–80 lakhs to ₹1.5 crore".
    metric = r"(?:₹\s?[\d,.]+(?:[–-][\d,.]+)?\s?(?:lakh|lakhs|crore|crores)?|\$\s?[\d,.]+(?:\s?(?:million|billion|k|m))?|\b\d+(?:[–-]\d+)?%)"
    achievement = re.search(
        rf"\b(?:scaled|increased|reduced|grew|generated|achieved|improved|"
        rf"managed|drove|delivered|raised|lowered)\b.*?"
        rf"{metric}.*?\bto\s+{metric}",
        text,
        re.IGNORECASE,
    )
    if achievement:
        candidate = achievement.group(0).strip()
        if len(candidate.split()) <= max_words:
            return candidate.rstrip(".,;:") + "."

    # Semicolon-separated clauses often contain the useful achievement.
    semicolon_parts = [p.strip() for p in text.split(";") if p.strip()]
    if len(semicolon_parts) > 1:
        candidates = [
            p for p in semicolon_parts
            if re.search(r"(?:₹|\$|%)", p)
        ]
        candidate = candidates[0] if candidates else semicolon_parts[0]
        if len(candidate.split()) <= max_words:
            return candidate.rstrip(".") + "."

    # Preserve a concise opening clause instead of cutting a sentence midway.
    comma_parts = [p.strip() for p in re.split(r",\s+", text) if p.strip()]
    if comma_parts and 6 <= len(comma_parts[0].split()) <= max_words:
        return comma_parts[0].rstrip(".") + "."

    # Final fallback: word-limit and remove dangling conjunctions/prepositions.
    shortened = " ".join(words[:max_words])
    shortened = re.sub(
        r"\s+(?:and|or|but|for|to|by|with|from|of|in|on)$",
        "",
        shortened,
        flags=re.IGNORECASE,
    )
    return shortened.rstrip(".,;:") + "."


def _parse_experience_blocks(text):
    """Parse experience entries using date ranges as reliable record boundaries."""
    lines = [_clean_experience_line(x) for x in text.splitlines()]
    lines = [x for x in lines if x]

    date_indexes = [i for i, line in enumerate(lines) if DATE_RANGE_RE.search(line)]
    blocks = []

    for position, date_index in enumerate(date_indexes):
        start = date_indexes[position - 1] + 1 if position > 0 else 0
        end = date_indexes[position + 1] if position + 1 < len(date_indexes) else len(lines)
        block = lines[start:end]
        if block:
            blocks.append(block)

    return blocks


def _clean_role_label(role):
    """Make multi-part CV job titles read naturally in the Experience tab."""
    role = re.sub(r"\s*\|\s*", " ", role).strip()
    role = re.sub(r"\s+", " ", role)
    role = re.sub(r"\bmanager\b", "Manager", role, flags=re.IGNORECASE)
    role = re.sub(r"\bpart time\b", "(Part-time)", role, flags=re.IGNORECASE)
    role = re.sub(r"\bintern\b", "Intern", role, flags=re.IGNORECASE)
    return role


def _undated_experience(text):
    """Internships / roles written without dates: still show them, grouped per entry."""
    lines = [_clean_experience_line(x) for x in (text or "").splitlines()]
    lines = [x for x in lines if x and not DURATION_RE.match(x.strip())]
    entries, header, bullets = [], [], []

    def heading_like(line):
        lower = line.lower()
        return (
            len(line.split()) <= 9
            and not line.rstrip().endswith((".", "!", "?"))
            and not any(lower.startswith(w) for w in ACTION_WORDS)
        )

    def flush():
        if header or bullets:
            head = header[0] if header else "Experience"
            sub = header[1] if len(header) > 1 else ""
            entry = [head] + ([sub] if sub else [])
            entry += [f"• {b[:1].upper() + b[1:]}" for b in bullets[:2]]
            entries.append("\n".join(entry))

    pending = ""
    for line in lines:
        if heading_like(line) and not pending:
            if bullets:
                flush()
                header, bullets = [], []
            header.append(re.sub(r"(?i)^profile\s*[:\-–]\s*", "", line))
        else:
            pending = (pending + " " + line).strip()
            if line.endswith((".", "!", "?")):
                bullets.append(_shorten_experience_point(pending))
                pending = ""
    if pending:
        bullets.append(_shorten_experience_point(pending))
    flush()
    return "\n\n".join(entries)


def format_experience_short(text):
    """
    Convert the full Experience section into a compact, consistent format.

    Each entry contains:
      Company — Job title
      Location | Date range
      1–2 concise achievement/responsibility bullets

    Company descriptions and repeated marketing copy from the source CV are
    intentionally omitted from the Experience tab.
    """
    blocks = _parse_experience_blocks(text)
    formatted = []

    for block in blocks:
        date_line = next((line for line in block if DATE_RANGE_RE.search(line)), "")
        if not date_line:
            continue

        date_match = DATE_RANGE_RE.search(date_line)
        display_date = date_match.group(0) if date_match else date_line

        # Header is everything before the date line.
        date_pos = block.index(date_line)
        header = block[:date_pos]
        body = block[date_pos + 1:]

        location = ""
        role = ""
        company = ""

        # Internship-style CVs: "PROFILE: Digital Marketing Intern" below the date,
        # and a duration line such as "8 Weeks".
        duration = ""
        cleaned_body = []
        for line in body:
            stripped = line.strip()
            profile = re.match(
                r"(?i)^(?:profile|role|designation|position|title|job title)\s*[:\-–]\s*(.+)$",
                stripped,
            )
            if DURATION_RE.match(stripped):
                duration = duration or stripped.strip("() ")
                continue
            if profile and not role:
                role = profile.group(1).strip()
                continue
            cleaned_body.append(line)
        body = cleaned_body
        if role:
            # Drop a line that only restates the title, e.g. "Intern – Company".
            body = [
                line for line in body
                if not (
                    line.lower().startswith(role.lower())
                    and len(line.split()) <= 8
                    and not line.rstrip().endswith(".")
                )
            ]

        # Example: | Pune | Digital Marketing Analyst
        # Also handles: | Pune | Customer Relations | Intern
        def _has_role(value):
            return any(word in value.lower() for word in ROLE_WORDS)

        def _plain_company_line(line):
            candidate = line.strip("|").strip()
            return bool(
                candidate and "|" not in line
                and len(candidate.split()) <= 8
                and not candidate.endswith((".", "!", "?"))
                and not any(candidate.lower().startswith(w) for w in ACTION_WORDS)
                and not DATE_RANGE_RE.search(candidate)
            )

        place_words = {"remote", "hybrid", "onsite", "on-site", "work from home", "wfh"}

        # "Company | Role" or "Company | City | Role" on ONE line is only used when
        # the CV has no separate company line (otherwise the first piece is a city).
        if not role and not any(_plain_company_line(l) for l in header):
            for line in header:
                if line.lstrip().startswith("|"):
                    continue
                pieces = [p.strip() for p in line.split("|") if p.strip()]
                if pieces and pieces[0].lower() in place_words:
                    continue
                if len(pieces) == 2 and _has_role(pieces[1]) and len(pieces[0].split()) <= 8:
                    company, role = pieces[0], pieces[1]
                    break
                if (
                    len(pieces) >= 3 and not _has_role(pieces[0])
                    and len(pieces[0].split()) <= 8
                    and any(_has_role(p) for p in pieces[2:])
                ):
                    company, location, role = pieces[0], pieces[1], " | ".join(pieces[2:])
                    break

        for line in ([] if role else header):
            pieces = [p.strip() for p in line.split("|") if p.strip()]
            if len(pieces) >= 2:
                possible_location = pieces[0]
                possible_role = pieces[-1]
                role_candidate = possible_role
                # For formats such as "Pune | Customer Relations | Intern",
                # keep the complete role phrase instead of only "Intern".
                if len(pieces) >= 3:
                    role_candidate = " | ".join(pieces[1:])
                if len(possible_location) <= 40 and any(
                    word in role_candidate.lower() for word in ROLE_WORDS
                ):
                    location = possible_location
                    role = role_candidate
                    break

        # If the role line wasn't found in pipe-separated form, use a line
        # containing a job-title keyword.
        if not role:
            for line in header:
                if any(word in line.lower() for word in ROLE_WORDS):
                    role = line.strip("|").strip()
                    break

        if role and not company and "|" not in role:
            for separator in (" at ", " - ", " – ", " — ", ", "):
                if separator in role:
                    left, right = [p.strip() for p in role.split(separator, 1)]
                    if left and right and len(left.split()) <= 8 and len(right.split()) <= 8:
                        if _has_role(left) and not _has_role(right):
                            role, company = left, right
                        elif _has_role(right) and not _has_role(left):
                            company, role = left, right
                        else:
                            continue
                        break

        # Company is a separate short header line. Prefer lines without pipes
        # so a full location/title line is never mistaken for the company.
        for line in ([] if company else reversed(header)):
            candidate = line.strip("|").strip()
            if not candidate or candidate == role or candidate == location:
                continue
            if "|" in line:
                continue
            if candidate.endswith((".", "!", "?")) or any(
                candidate.lower().startswith(w) for w in ACTION_WORDS
            ):
                continue  # a sentence from the previous entry, not a company name
            if len(candidate.split()) <= 8 and not DATE_RANGE_RE.search(candidate):
                company = candidate
                break

        # Rebuild wrapped PDF lines into complete paragraphs/bullets first.
        # This prevents a sentence split across two PDF text blocks from
        # being displayed as an incomplete experience point.
        paragraphs = []
        current = ""
        for line in body:
            if not current:
                current = line
            else:
                current += " " + line

            if line.endswith((".", "!", "?")):
                paragraphs.append(current)
                current = ""

        if current:
            paragraphs.append(current)

        # Action-oriented paragraphs are the useful part of the experience section.
        action_lines = [
            paragraph for paragraph in paragraphs
            if any(paragraph.lower().startswith(w) for w in ACTION_WORDS)
        ]

        # Keep only the two most relevant points for a short and precise tab.
        bullets = []
        for point in action_lines:
            point = _shorten_experience_point(point)
            if point not in bullets:
                bullets.append(point)
            if len(bullets) == 2:
                break

        # Fallback if a non-standard CV has no recognizable action verb.
        if not bullets:
            fallback = [
                _shorten_experience_point(x)
                for x in body
                if len(x.split()) >= 6
            ]
            bullets = fallback[:2]

        if not role:
            role = "Role not detected"
        if not company:
            company = "Company not detected"

        role = _clean_role_label(role)
        company = company.strip()

        # Keep bullet capitalization consistent even when the PDF extracted
        # a lowercase action verb.
        bullets = [
            (bullet[:1].upper() + bullet[1:]) if bullet else bullet
            for bullet in bullets
        ]

        heading = f"{company} — {role}"
        if duration:
            display_date = f"{display_date} ({duration})"
        if location:
            subheading = f"{location} | {display_date}"
        else:
            subheading = display_date

        entry = [heading, subheading]
        entry.extend(f"• {bullet}" for bullet in bullets)
        formatted.append("\n".join(entry))

    if not formatted:
        undated = _undated_experience(text)
        if undated:
            return undated
        return "No structured experience entries were detected."

    return "\n\n".join(formatted)


def extract_experience(text):
    sections = detect_sections(text)
    return sections.get("Experience", "")


def extract_job_roles(text):
    """Extract likely job titles, especially from common CV separator formats."""
    roles = []
    role_words = (
        "manager", "analyst", "associate", "intern", "consultant",
        "executive", "developer", "engineer", "specialist", "coordinator",
        "lead", "director", "assistant", "advisor", "trainee", "buyer",
        "mentor", "officer", "administrator", "designer"
    )

    for raw_line in text.splitlines():
        line = normalize_text(raw_line).strip(" |•-\t")
        if not line or len(line) > 100:
            continue

        # CV format such as:
        # | Pune | Digital Marketing Analyst | Company |
        pieces = [p.strip() for p in line.split("|") if p.strip()]
        candidates = pieces if pieces else [line]

        for candidate in candidates:
            lower = candidate.lower()

            if not any(word in lower for word in role_words):
                continue

            # Avoid long bullet/sentence descriptions.
            if len(candidate.split()) > 8:
                continue

            # Avoid obvious non-title phrases.
            if any(prefix in lower for prefix in (
                "utilized ", "developed ", "handled ", "analysed ",
                "analyzed ", "managed ", "assisted ", "communicated ",
                "leading ", "acting ", "owning ", "driving ",
                "major in ", "research ", "marketing core"
            )):
                continue

            # Ignore academic/course phrases that happen to contain
            # words such as "management".
            if any(marker in lower for marker in (
                "international business", "case studies", "csr."
            )):
                continue

            if candidate not in roles:
                roles.append(candidate)

    return roles[:15]


# ============================================================
# CERTIFICATION EXTRACTION
# ============================================================

def extract_certifications(text):

    sections = detect_sections(
        text
    )

    certification_text = sections.get(
        "Certifications",
        ""
    )

    if certification_text:
        return certification_text

    lines = text.splitlines()

    matches = []

    keywords = [
        "certification",
        "certified",
        "udemy",
        "coursera",
        "course",
        "certificate"
    ]

    for line in lines:

        if any(
            keyword in line.lower()
            for keyword in keywords
        ):

            if len(line.strip()) > 3:

                matches.append(
                    line.strip()
                )

    return "\n".join(
        dict.fromkeys(matches)
    )


# ============================================================
# LANGUAGE EXTRACTION
# ============================================================

COMMON_LANGUAGES = [
    "English",
    "Hindi",
    "Marathi",
    "Tamil",
    "Telugu",
    "Kannada",
    "Malayalam",
    "Gujarati",
    "Bengali",
    "Punjabi",
    "Urdu",
    "French",
    "German",
    "Spanish",
    "Italian",
    "Japanese",
    "Chinese",
    "Arabic"
]


def extract_languages(text):

    sections = detect_sections(
        text
    )

    language_text = sections.get(
        "Languages",
        ""
    )

    found = []

    search_text = (
        language_text
        if language_text
        else text
    )

    for language in COMMON_LANGUAGES:

        if re.search(
            r"\b"
            + re.escape(language)
            + r"\b",
            search_text,
            re.IGNORECASE
        ):

            found.append(
                language
            )

    return found


# ============================================================
# CAREER AREA DETECTION
# ============================================================

CAREER_AREAS = {

    "Digital Marketing": [
        "digital marketing",
        "meta ads",
        "google ads",
        "media buying",
        "performance marketing",
        "social media marketing",
        "seo",
        "sem",
        "roas",
        "campaign"
    ],

    "Business Analytics": [
        "business analytics",
        "business analyst",
        "data analysis",
        "analytics",
        "sql",
        "excel",
        "data visualization"
    ],

    "Finance": [
        "financial analysis",
        "financial modeling",
        "accounting",
        "finance",
        "valuation",
        "investment",
        "audit"
    ],

    "Human Resources": [
        "human resources",
        "recruitment",
        "talent acquisition",
        "employee engagement",
        "hr analytics"
    ],

    "Operations & Management": [
        "operations",
        "project management",
        "supply chain",
        "process improvement",
        "business management",
        "operations management"
    ],

    "Technology": [
        "python",
        "java",
        "software development",
        "programming",
        "machine learning",
        "artificial intelligence",
        "cloud computing"
    ]
}


def detect_career_areas(text):

    areas = []

    for area, keywords in CAREER_AREAS.items():

        matches = 0

        for keyword in keywords:

            if contains_skill(
                text,
                keyword
            ):

                matches += 1

        if matches >= 1:

            areas.append(
                (
                    area,
                    matches
                )
            )

    areas.sort(
        key=lambda x: x[1],
        reverse=True
    )

    return [
        area
        for area, score in areas
    ]


# ============================================================
# RESUME COMPLETENESS
# ============================================================

def calculate_completeness(
    text,
    name,
    email,
    phone,
    education,
    experience,
    skills,
    certifications,
    languages
):

    checks = {

        "Name": bool(name),

        "Email": bool(email),

        "Phone": bool(phone),

        "Education": bool(education),

        "Experience": bool(experience),

        "Skills": bool(skills),

        "Certifications/Courses":
            bool(certifications),

        "Languages":
            bool(languages)

    }

    score = (
        sum(checks.values())
        /
        len(checks)
    ) * 100

    return round(
        score,
        1
    ), checks


# ============================================================
# RESUME ANALYSIS
# ============================================================

def analyze_resume(
    file_path
):

    text = extract_file_text(
        file_path
    )

    if not text.strip():

        raise ValueError(
            "No readable text was found in the resume."
        )

    name = extract_name(
        text
    )

    email = extract_email(
        text
    )

    phone = extract_phone(
        text
    )

    education = extract_education(
        text
    )

    experience = extract_experience(
        text
    )

    certifications = extract_certifications(
        text
    )

    languages = extract_languages(
        text
    )

    skills_by_category = extract_skills_by_category(
        text
    )

    all_skills = sorted(
        set(
            skill
            for skills in skills_by_category.values()
            for skill in skills
        )
    )

    roles = extract_job_roles(
        text
    )

    career_areas = detect_career_areas(
        text
    )

    completeness, section_status = calculate_completeness(
        text,
        name,
        email,
        phone,
        education,
        experience,
        all_skills,
        certifications,
        languages
    )

    return {

        "file": file_path,

        "filename":
            os.path.basename(
                file_path
            ),

        "text":
            text,

        "name":
            name,

        "email":
            email,

        "phone":
            phone,

        "education":
            education,

        "experience":
            experience,

        "certifications":
            certifications,

        "languages":
            languages,

        "skills_by_category":
            skills_by_category,

        "all_skills":
            all_skills,

        "roles":
            roles,

        "career_areas":
            career_areas,

        "completeness":
            completeness,

        "section_status":
            section_status
    }



# ============================================================
# MULTI-CV SUPPORT: FILE COLLECTION
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".pdf", ".docx", ".docm", ".dotx", ".doc", ".dot", ".rtf", ".odt", ".ott",
    ".txt", ".text", ".md", ".html", ".htm", ".xml", ".csv", ".json", ".pptx",
    ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif",
    ".xlsx", ".xlsm", ".epub", ".heic", ".heif", ".jfif", ".wps", ".dotm", ".pptm"
}

SUPPORTED_PATTERN = (
    "*.pdf *.docx *.docm *.dotx *.doc *.dot *.rtf *.odt *.ott *.txt *.text "
    "*.md *.html *.htm *.xml *.csv *.json *.pptx "
    "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.gif "
    "*.xlsx *.xlsm *.epub *.heic *.heif *.jfif *.wps *.dotm *.pptm"
)


def collect_supported_files(paths, limit=500):
    """
    Expand files / folders into a flat list of CV files.
    Files picked explicitly are always accepted (any type); only files found
    while scanning a folder are filtered by known extension.
    """
    files = []
    skipped = 0

    for raw in paths:
        path = Path(raw)

        if path.is_dir():
            for folder, _dirs, names in os.walk(path):
                for name in sorted(names):
                    if name.startswith(("~$", ".")):
                        continue
                    candidate = Path(folder) / name
                    if candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                        files.append(str(candidate))
        elif path.is_file():
            if not path.name.startswith("~$"):
                files.append(str(path))
            else:
                skipped += 1

        if len(files) >= limit:
            break

    return files[:limit], skipped


# ============================================================
# MULTI-CV SUPPORT: COMBINED ANALYSIS
# ============================================================

def candidate_label(analysis):
    """Readable label for a candidate: detected name, else the file name."""
    name = (analysis.get("name") or "").strip()
    return name or Path(analysis["filename"]).stem


def _unique_labels(analyses):
    labels = [candidate_label(a) for a in analyses]
    counts = Counter(label.lower() for label in labels)
    result = []
    for analysis, label in zip(analyses, labels):
        if counts[label.lower()] > 1:
            label = f"{label} ({analysis['filename']})"
        result.append(label)
    return result


def parse_required_skills(text):
    """Turn 'Python, SQL; Power BI' into ['Python', 'SQL', 'Power BI']."""
    found = []
    for part in re.split(r"[,;\n]+", text or ""):
        part = part.strip()
        if part and part.lower() not in [x.lower() for x in found]:
            found.append(part)
    return found


def skill_in_resume(analysis, skill):
    """True when a (possibly custom) skill appears anywhere in the CV."""
    if any(skill.lower() == known.lower() for known in analysis["all_skills"]):
        return True
    return contains_skill(analysis["text"], skill)


# Acronyms and product names that str.title() would spell wrongly
# (e.g. "Sql", "Seo", "Rest Api"). Purely a display fix; detection is unchanged.
for _key, _label in {
    "sql": "SQL", "mysql": "MySQL", "postgresql": "PostgreSQL", "nosql": "NoSQL",
    "etl": "ETL", "seo": "SEO", "sem": "SEM", "hr": "HR", "hr analytics": "HR Analytics",
    "hris": "HRIS", "ats": "ATS", "okr": "OKR", "gst": "GST", "dcf": "DCF", "irr": "IRR",
    "npv": "NPV", "wacc": "WACC", "kpi": "KPI", "swot": "SWOT", "bpm": "BPM",
    "erp": "ERP", "sap": "SAP", "vba": "VBA", "api": "API", "api development": "API Development",
    "rest api": "REST API", "ci/cd": "CI/CD", "html": "HTML", "css": "CSS", "php": "PHP",
    "numpy": "NumPy", "scipy": "SciPy", ".net": ".NET", "ui/ux": "UI/UX",
    "tensorflow": "TensorFlow", "pytorch": "PyTorch", "mongodb": "MongoDB", "github": "GitHub",
    "powerpoint": "PowerPoint", "linkedin": "LinkedIn", "wordpress": "WordPress",
    "bigquery": "BigQuery", "hubspot": "HubSpot", "salesforce": "Salesforce", "jira": "Jira",
    "devops": "DevOps", "fastapi": "FastAPI", "graphql": "GraphQL", "oop": "OOP",
}.items():
    DISPLAY_NAMES.setdefault(_key, _label)


# Reverse lookup: display skill name -> category (first category wins).
SKILL_CATEGORY = {}
for _category, _skills in SKILLS.items():
    for _skill in _skills:
        SKILL_CATEGORY.setdefault(display_skill(_skill), _category)


def build_combined_analysis(analyses, required=None):
    """
    Analyse several CVs together.

    Returns a dictionary holding ranked comparison rows, skill frequency
    across the group, skills shared by every CV, skills only one CV has,
    career-area distribution, languages and (optionally) a match score
    against a list of required skills.
    """
    required = [r for r in (required or []) if r]
    labels = _unique_labels(analyses)
    count = len(analyses)

    rows = []
    skill_holders = {}
    area_holders = {}
    language_holders = {}

    for analysis, label in zip(analyses, labels):
        for skill in analysis["all_skills"]:
            skill_holders.setdefault(skill, []).append(label)
        for area in analysis["career_areas"]:
            area_holders.setdefault(area, []).append(label)
        for language in analysis["languages"]:
            language_holders.setdefault(language, []).append(label)

        row = {
            "entry_id": analysis.get("_id"),
            "label": label,
            "name": analysis["name"],
            "filename": analysis["filename"],
            "email": analysis["email"],
            "phone": analysis["phone"],
            "completeness": analysis["completeness"],
            "skill_count": len(analysis["all_skills"]),
            "role_count": len(analysis["roles"]),
            "top_area": analysis["career_areas"][0] if analysis["career_areas"] else "—",
            "career_areas": list(analysis["career_areas"]),
            "skills": list(analysis["all_skills"]),
            "roles": list(analysis["roles"]),
            "match_pct": None,
            "matched": [],
            "missing": [],
            "analysis": analysis,
        }

        if required:
            matched = [r for r in required if skill_in_resume(analysis, r)]
            row["matched"] = matched
            row["missing"] = [r for r in required if r not in matched]
            row["match_pct"] = len(matched) / len(required) * 100

        rows.append(row)

    rows.sort(key=lambda r: (
        -(r["match_pct"] or 0),
        -r["completeness"],
        -r["skill_count"],
        r["label"].lower(),
    ))
    for rank, row in enumerate(rows, start=1):
        row["rank"] = rank

    skill_freq = sorted(
        ((skill, len(holders), holders) for skill, holders in skill_holders.items()),
        key=lambda item: (-item[1], item[0].lower())
    )
    career_freq = sorted(
        ((area, len(holders), holders) for area, holders in area_holders.items()),
        key=lambda item: (-item[1], item[0].lower())
    )
    language_freq = sorted(
        ((lang, len(holders), holders) for lang, holders in language_holders.items()),
        key=lambda item: (-item[1], item[0].lower())
    )

    common_skills = []
    unique_skills = {}
    if count > 1:
        common_skills = sorted(
            (s for s, holders in skill_holders.items() if len(holders) == count),
            key=str.lower
        )
        for skill, holders in skill_holders.items():
            if len(holders) == 1:
                unique_skills.setdefault(holders[0], []).append(skill)
        for label in unique_skills:
            unique_skills[label].sort(key=str.lower)

    return {
        "count": count,
        "rows": rows,
        "required": required,
        "avg_completeness": (sum(r["completeness"] for r in rows) / count) if count else 0.0,
        "avg_skills": (sum(r["skill_count"] for r in rows) / count) if count else 0.0,
        "total_unique_skills": len(skill_holders),
        "skill_freq": skill_freq,
        "career_freq": career_freq,
        "language_freq": language_freq,
        "common_skills": common_skills,
        "unique_skills": unique_skills,
        "top_area": career_freq[0][0] if career_freq else "—",
        "generated": datetime.now(),
    }


def combined_summary_sentences(combined):
    """Short plain-English findings used at the top of generated reports."""
    rows = combined["rows"]
    count = combined["count"]
    sentences = []

    if count == 0:
        return ["No CVs have been analyzed yet."]

    if count == 1:
        row = rows[0]
        sentences.append(
            f"{row['label']}'s CV is {row['completeness']:.1f}% complete "
            f"and contains {row['skill_count']} recognized skills."
        )
        return sentences

    sentences.append(
        f"{count} CVs were analyzed together. The average completeness is "
        f"{pct1(combined['avg_completeness'])}% and {combined['total_unique_skills']} "
        f"different skills were recognized across the group."
    )

    if combined["skill_freq"]:
        top = [s for s, _c, _h in combined["skill_freq"][:5]]
        sentences.append("The most common skills are " + ", ".join(top) + ".")

    if combined["career_freq"]:
        sentences.append(
            f"The most frequent career area is {combined['career_freq'][0][0]}."
        )

    best = max(rows, key=lambda r: (r["completeness"], r["skill_count"]))
    sentences.append(
        f"The most complete CV belongs to {best['label']} "
        f"({best['completeness']:.1f}%)."
    )

    if combined["required"]:
        top = rows[0]
        sentences.append(
            f"Against the target skills ({', '.join(combined['required'])}), "
            f"the best match is {top['label']} at {top['match_pct']:.0f}%."
        )

    if combined["common_skills"]:
        sentences.append(
            "Skills shared by every CV: " + ", ".join(combined["common_skills"]) + "."
        )

    return sentences


def experience_entries(experience_text):
    """Split the compact experience format into structured entries."""
    if not experience_text:
        return []

    formatted = format_experience_short(experience_text)
    entries = []

    for block in formatted.split("\n\n"):
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if not lines or lines[0].startswith("No structured experience"):
            continue
        entries.append({
            "heading": lines[0],
            "subheading": lines[1] if len(lines) > 1 else "",
            "bullets": [line.lstrip("•").strip() for line in lines[2:]],
        })

    return entries


def pct1(value):
    """Format with one decimal, rounding halves up (matches Excel)."""
    return f"{value + 1e-9:.1f}"


def _text_lines(value):
    return [line.strip() for line in (value or "").splitlines() if line.strip()]


# ============================================================
# REPORT EXPORT (WORD / EXCEL / TEXT)
# ============================================================

_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean_office_text(value):
    """Remove control characters that Word/Excel files cannot store."""
    if value is None:
        return ""
    return _CONTROL_CHARS.sub("", str(value))


def default_report_name(analyses_count):
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    prefix = "CV_Analysis_Report" if analyses_count != 1 else "CV_Analysis"
    return f"{prefix}_{stamp}"


# ------------------------------------------------------------
# PLAIN TEXT
# ------------------------------------------------------------

def build_text_report(analyses, combined):
    out = []
    out.append("AI RESUME / CV ANALYSIS REPORT")
    out.append("=" * 70)
    out.append(f"Generated: {combined['generated'].strftime('%d %B %Y, %H:%M')}")
    out.append("")

    if combined["count"] > 1:
        out.append("COMBINED ANALYSIS")
        out.append("-" * 70)
        for sentence in combined_summary_sentences(combined):
            out.append(sentence)
        out.append("")
        out.append("CANDIDATE RANKING")
        for row in combined["rows"]:
            line = (
                f"{row['rank']}. {row['label']} — {row['completeness']:.1f}% complete, "
                f"{row['skill_count']} skills, {row['top_area']}"
            )
            if row["match_pct"] is not None:
                line += f", {row['match_pct']:.0f}% match"
            out.append(line)
        out.append("")
        out.append("MOST COMMON SKILLS")
        for skill, number, _holders in combined["skill_freq"][:15]:
            out.append(f"  {skill}: {number} of {combined['count']} CVs")
        out.append("")

    for analysis in analyses:
        out.append("=" * 70)
        out.append(f"Candidate: {analysis['name'] or 'Not detected'}")
        out.append(f"Email: {analysis['email'] or 'Not detected'}")
        out.append(f"Phone: {analysis['phone'] or 'Not detected'}")
        out.append(f"Resume: {analysis['filename']}")
        out.append("")
        out.append(f"Completeness Score: {analysis['completeness']:.1f}%")
        out.append("")
        out.append("CAREER AREAS")
        for index, area in enumerate(analysis["career_areas"], start=1):
            out.append(f"{index}. {area}")
        out.append("")
        out.append("SKILLS")
        for category, skills in analysis["skills_by_category"].items():
            out.append(f"\n{category}:")
            for index, skill in enumerate(skills, start=1):
                out.append(f"  {index}. {skill}")
        out.append("")
        out.append("JOB ROLES")
        for index, role in enumerate(analysis["roles"], start=1):
            out.append(f"{index}. {role}")
        out.append("")
        out.append("EDUCATION")
        out.append(analysis["education"] or "Not detected.")
        out.append("")
        out.append("EXPERIENCE")
        out.append(analysis["experience"] or "Not detected.")
        out.append("")
        out.append("COURSES / CERTIFICATIONS")
        out.append(analysis["certifications"] or "Not detected.")
        out.append("")
        out.append("LANGUAGES")
        for index, language in enumerate(analysis["languages"], start=1):
            out.append(f"{index}. {language}")
        out.append("")

    return "\n".join(out)


def export_text(path, analyses, combined):
    with open(path, "w", encoding="utf-8") as file:
        file.write(build_text_report(analyses, combined))


# ------------------------------------------------------------
# WORD (.docx)
# ------------------------------------------------------------

_W_NAVY = "1F2D4D"
_W_PRIMARY = "2F5597"
_W_SOFT = "EAF0F8"
_W_ZEBRA = "F6F8FB"
_W_BORDER = "CBD3E1"
_W_MUTED = "5B667A"
_W_TEXT = "1F2937"
_W_GOOD = "1E7B4F"
_W_WARN = "B45F06"
_W_BAD = "B3261E"


def _w_level_color(value):
    if value >= 75:
        return _W_GOOD
    if value >= 50:
        return _W_WARN
    return _W_BAD


def _w_style_font(style, name, size=None, bold=None, color=None):
    style.font.name = name
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = RGBColor.from_string(color)

    # Remove theme-font attributes so the font name above really applies.
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def _w_shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _w_table_borders(table, color=_W_BORDER, size=4):
    tbl_pr = table._tbl.tblPr
    for old in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)
    look = tbl_pr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(borders)
    else:
        tbl_pr.append(borders)


def _w_cell_margins(table, top=60, bottom=60, left=100, right=100):
    tbl_pr = table._tbl.tblPr
    margins = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        element = OxmlElement(f"w:{side}")
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")
        margins.append(element)
    look = tbl_pr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(margins)
    else:
        tbl_pr.append(margins)


def _w_cell(cell, text, bold=False, italic=False, size=9.5, color=_W_TEXT,
            align=None, fill=None):
    if fill:
        _w_shade(cell, fill)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    if align:
        paragraph.alignment = align
    run = paragraph.add_run(_clean_office_text(text))
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return paragraph


def _w_widths(table, widths_cm):
    table.autofit = False
    for column, width in zip(table.columns, widths_cm):
        column.width = Cm(width)          # table grid (used by LibreOffice / Word layout)
    for row in table.rows:
        for cell, width in zip(row.cells, widths_cm):
            cell.width = Cm(width)


def _w_para(container, text="", bold=False, italic=False, size=None,
            color=None, align=None, after=None, before=None, style=None,
            keep_next=False):
    paragraph = container.add_paragraph(style=style) if style else container.add_paragraph()
    if text:
        run = paragraph.add_run(_clean_office_text(text))
        run.bold = bold
        run.italic = italic
        if size:
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
    if align is not None:
        paragraph.alignment = align
    if after is not None:
        paragraph.paragraph_format.space_after = Pt(after)
    if before is not None:
        paragraph.paragraph_format.space_before = Pt(before)
    if keep_next:
        paragraph.paragraph_format.keep_with_next = True
    return paragraph


def _w_heading(doc, text, level=1):
    heading = doc.add_heading(_clean_office_text(text), level=level)
    heading.paragraph_format.keep_with_next = True
    return heading


def _w_page_number(paragraph, field="PAGE"):
    """Insert a live PAGE (or NUMPAGES) field; every piece is its own small run."""
    def piece(element):
        run = paragraph.add_run()
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor.from_string(_W_MUTED)
        run._r.append(element)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = f" {field} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, placeholder, end):
        piece(element)


def _w_repeat_header(table):
    """Repeat the first row on every page and keep rows from splitting."""
    for index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
        if index == 0:
            header = OxmlElement("w:tblHeader")
            tr_pr.append(header)


# ------------------------------------------------------------
# Report libraries: loaded (and installed automatically) when first needed
# ------------------------------------------------------------

def _install_possible():
    """True when the app is allowed to pip-install missing helper libraries."""
    return (
        not getattr(sys, "frozen", False)
        and os.environ.get("RESUME_ANALYZER_NO_INSTALL") != "1"
    )


def _ensure_docx_support():
    """Make python-docx available (installing it if needed). Returns True/False."""
    global Document, Pt, RGBColor, Cm, WD_ALIGN_PARAGRAPH, WD_TABLE_ALIGNMENT
    global WD_CELL_VERTICAL_ALIGNMENT, OxmlElement, qn
    if Document is not None and "Pt" in globals():
        return True
    if _import_or_install("docx", "python-docx") is None:
        return False
    try:
        from docx import Document as _Document  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from docx.shared import Pt as _Pt, RGBColor as _RGB, Cm as _Cm  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from docx.enum.table import WD_TABLE_ALIGNMENT as _TA, WD_CELL_VERTICAL_ALIGNMENT as _VA  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from docx.oxml import OxmlElement as _Ox  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from docx.oxml.ns import qn as _qn  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    except Exception:
        return False
    Document, Pt, RGBColor, Cm = _Document, _Pt, _RGB, _Cm
    WD_ALIGN_PARAGRAPH, WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT = _AL, _TA, _VA
    OxmlElement, qn = _Ox, _qn
    return True


def _ensure_xlsx_support():
    """Make openpyxl available (installing it if needed). Returns True/False."""
    global openpyxl, Font, PatternFill, Alignment, Border, Side, get_column_letter
    global DataBarRule, ColorScaleRule, XlBarChart, XlReference
    global PageSetupProperties, ILLEGAL_CHARACTERS_RE
    if openpyxl is not None and "Font" in globals():
        return True
    if _import_or_install("openpyxl", "openpyxl") is None:
        return False
    try:
        import openpyxl as _op  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _Al, Border as _B, Side as _S  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.utils import get_column_letter as _gcl  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.formatting.rule import DataBarRule as _DB, ColorScaleRule as _CS  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.chart import BarChart as _BC, Reference as _Ref  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.worksheet.properties import PageSetupProperties as _PSP  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE as _ILL  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    except Exception:
        return False
    openpyxl = _op
    Font, PatternFill, Alignment, Border, Side = _F, _PF, _Al, _B, _S
    get_column_letter, DataBarRule, ColorScaleRule = _gcl, _DB, _CS
    XlBarChart, XlReference, PageSetupProperties, ILLEGAL_CHARACTERS_RE = _BC, _Ref, _PSP, _ILL
    return True


def _w_bar(paragraph, fraction, cells=10, size=8):
    """Small text progress bar (decorative; the number is always printed beside it)."""
    fraction = max(0.0, min(1.0, fraction))
    filled = int(round(fraction * cells))
    for count_, color in ((filled, "7FA3D9"), (cells - filled, "E1E6F0")):
        if count_ > 0:
            run = paragraph.add_run("█" * count_)
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor.from_string(color)


def _w_row_cant_split(table):
    """Stop every row of a table from being split across two pages."""
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))


def _w_keep_table_together(table):
    """Keep a small table on one page by chaining its paragraphs together."""
    rows = table.rows
    for row in rows[:-1]:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = True


def _w_cell_edges(cell, **edges):
    """Set individual cell borders, e.g. left=('24', '2F5597') or top=None."""
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:tcBorders")):
        tc_pr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        spec = edges.get(edge)
        if spec:
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), spec[0])
            element.set(qn("w:color"), spec[1])
        else:
            element.set(qn("w:val"), "nil")
        element.set(qn("w:space"), "0")
        borders.append(element)
    # Word expects tcBorders before shading / margins / vertical alignment.
    later = [qn(f"w:{name}") for name in
             ("shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark")]
    anchor = next((child for child in tc_pr if child.tag in later), None)
    if anchor is not None:
        anchor.addprevious(borders)
    else:
        tc_pr.append(borders)


def _w_tiles(doc, tiles, total_cm=17.0):
    """Row of summary tiles: (value, caption, hint)."""
    table = doc.add_table(rows=1, cols=len(tiles))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _w_table_borders(table, "FFFFFF", 18)
    _w_widths(table, [total_cm / len(tiles)] * len(tiles))
    for cell, (value, caption, hint) in zip(table.rows[0].cells, tiles):
        _w_shade(cell, _W_SOFT)
        cell.text = ""
        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(8)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(_clean_office_text(value))
        r1.bold = True
        r1.font.size = Pt(20 if len(str(value)) < 10 else 12)
        r1.font.color.rgb = RGBColor.from_string(_W_PRIMARY)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0 if hint else 8)
        r2 = p2.add_run(_clean_office_text(caption))
        r2.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = RGBColor.from_string(_W_TEXT)
        if hint:
            p3 = cell.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p3.paragraph_format.space_after = Pt(8)
            r3 = p3.add_run(_clean_office_text(hint))
            r3.italic = True
            r3.font.size = Pt(8)
            r3.font.color.rgb = RGBColor.from_string(_W_MUTED)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _w_row_cant_split(table)
    return table


def _w_callout(doc, title, items, fill=None, accent=None, width_cm=17.0):
    """Shaded note box: items are (bold lead-in, text) pairs."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _w_widths(table, [width_cm])
    cell = table.rows[0].cells[0]
    _w_shade(cell, fill or _W_SOFT)
    _w_cell_edges(cell, left=("24", accent or _W_PRIMARY))
    _w_cell_margins(table, 90, 90, 160, 140)
    cell.text = ""
    _w_row_cant_split(table)
    head = cell.paragraphs[0]
    head.paragraph_format.space_after = Pt(3)
    run = head.add_run(_clean_office_text(title))
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor.from_string(_W_NAVY)
    for index, (lead, text) in enumerate(items):
        paragraph = cell.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(2)
        if lead:
            lead_run = paragraph.add_run(_clean_office_text(lead) + " ")
            lead_run.bold = True
            lead_run.font.size = Pt(9.5)
            lead_run.font.color.rgb = RGBColor.from_string(_W_NAVY)
        body = paragraph.add_run(_clean_office_text(text))
        body.font.size = Pt(9.5)
    return table


def _w_score_guide(doc, has_match, required):
    items = [
        ("Profile completeness:",
         "how many of the 8 standard CV sections were found (Name, Email, Phone, Education, "
         "Experience, Skills, Certifications/Courses, Languages). It shows how complete the "
         "CV is - not the quality of the candidate or how well they would perform."),
        ("Skills found:",
         "the number of skills from the application's built-in list that appear in the CV. "
         "Other skills, or unusual wording, are not counted. More skills does not mean a "
         "better fit for a role."),
    ]
    if has_match:
        items.append((
            "Target-skill match:",
            "the share of the target skills (" + ", ".join(required) + ") found anywhere in the "
            "CV. It is a keyword check and does not show how well or how recently a skill was used."))
    items.append((
        "Please note:",
        "the application reads each CV automatically and can miss information held in images "
        "or unusual layouts. \"Not detected\" means not found by the application - please check "
        "the original CV. This report supports human review; it does not accept or reject "
        "any candidate."))
    return _w_callout(doc, "How to read the scores", items)


def export_docx(path, analyses, combined):
    if not _ensure_docx_support():
        raise RuntimeError(
            "Word export requires python-docx, and it could not be installed "
            "automatically. Install it manually with:\n"
            "pip install python-docx"
        )

    doc = Document()

    # ---- page + base styles ------------------------------------------
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.9)
    section.header_distance = Cm(0.9)
    section.footer_distance = Cm(0.9)

    normal = doc.styles["Normal"]
    _w_style_font(normal, "Calibri", 10.5, color=_W_TEXT)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.08
    _w_style_font(doc.styles["Heading 1"], "Calibri", 17, True, _W_NAVY)
    _w_style_font(doc.styles["Heading 2"], "Calibri", 12.5, True, _W_PRIMARY)
    _w_style_font(doc.styles["Heading 3"], "Calibri", 11, True, _W_TEXT)
    doc.styles["Heading 1"].paragraph_format.space_before = Pt(16)
    doc.styles["Heading 1"].paragraph_format.space_after = Pt(6)
    doc.styles["Heading 2"].paragraph_format.space_before = Pt(12)
    doc.styles["Heading 2"].paragraph_format.space_after = Pt(4)
    _w_style_font(doc.styles["List Bullet"], "Calibri", 10.5, color=_W_TEXT)

    # thin rule under Heading 1 (inserted before <w:spacing> to keep the XML order valid)
    h1_ppr = doc.styles["Heading 1"].element.get_or_add_pPr()
    rule = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), _W_PRIMARY)
    rule.append(bottom)
    spacing = h1_ppr.find(qn("w:spacing"))
    if spacing is not None:
        spacing.addprevious(rule)
    else:
        h1_ppr.append(rule)

    # ---- header + footer ----------------------------------------------------
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_run = header.add_run("CV Analysis Report  ·  Decision-support information for human review")
    header_run.font.size = Pt(8.5)
    header_run.font.color.rgb = RGBColor.from_string(_W_MUTED)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note = footer.add_run(f"{APP_TITLE}  ·  {combined['generated'].strftime('%d %b %Y')}  ·  Page ")
    note.font.size = Pt(8.5)
    note.font.color.rgb = RGBColor.from_string(_W_MUTED)
    _w_page_number(footer)
    of_run = footer.add_run(" of ")
    of_run.font.size = Pt(8.5)
    of_run.font.color.rgb = RGBColor.from_string(_W_MUTED)
    _w_page_number(footer, "NUMPAGES")

    rows = combined["rows"]
    many = combined["count"] > 1
    required = combined["required"]
    has_match = bool(required)
    generated_text = combined["generated"].strftime("%d %B %Y, %H:%M")

    # ---- title banner -----------------------------------------------------
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    _w_widths(banner, [17.0])
    cell = banner.rows[0].cells[0]
    _w_shade(cell, _W_NAVY)
    cell.text = ""
    title_p = cell.paragraphs[0]
    title_p.paragraph_format.space_before = Pt(16)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.paragraph_format.left_indent = Cm(0.3)
    title_run = title_p.add_run("CV Analysis Report" if many else "CV Analysis")
    title_run.bold = True
    title_run.font.size = Pt(26)
    title_run.font.color.rgb = RGBColor.from_string("FFFFFF")
    sub_p = cell.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(2)
    sub_p.paragraph_format.left_indent = Cm(0.3)
    subtitle = (
        f"{combined['count']} CVs analyzed together" if many
        else _clean_office_text(rows[0]["label"]) if rows else ""
    )
    sub_run = sub_p.add_run(f"{subtitle}   ·   Generated {generated_text}")
    sub_run.font.size = Pt(10.5)
    sub_run.font.color.rgb = RGBColor.from_string("D6E0F2")
    tag_p = cell.add_paragraph()
    tag_p.paragraph_format.space_after = Pt(14)
    tag_p.paragraph_format.left_indent = Cm(0.3)
    tag_run = tag_p.add_run("HR decision-support report  ·  Automated analysis for human reviewers")
    tag_run.italic = True
    tag_run.font.size = Pt(9)
    tag_run.font.color.rgb = RGBColor.from_string("AFC0DC")

    _w_para(doc, "", after=2)

    # ---- combined overview --------------------------------------------------
    if many:
        tiles = [
            (str(combined["count"]), "CVs analyzed", "Total in this report"),
            (f"{pct1(combined['avg_completeness'])}%", "Average profile completeness",
             "Share of CV sections found"),
            (str(combined["total_unique_skills"]), "Unique skills found", "Across all CVs"),
            (combined["top_area"], "Top career area", "Most frequent in the group"),
        ]
        _w_tiles(doc, tiles)
        if has_match:
            _w_para(doc, "", after=0, size=4)
            full = sum(1 for r in rows if r["match_pct"] is not None and r["match_pct"] >= 100)
            _w_tiles(doc, [
                (str(len(required)), "Target skills entered", "Used for the match"),
                (f"{rows[0]['match_pct']:.0f}%", "Highest target-skill match", rows[0]["label"]),
                (str(full), "CVs with all target skills", "Every target skill found"),
                (rows[0]["label"], "Top ranked candidate", "First in the ranking order"),
            ])

        _w_heading(doc, "Executive Summary", 1)
        _w_heading(doc, "Key Findings", 2)
        for sentence in combined_summary_sentences(combined):
            _w_para(doc, sentence, style="List Bullet", after=3)
        _w_para(doc, "", after=2)
        _w_score_guide(doc, has_match, required)

        # Ranking table
        _w_heading(doc, "Candidate Comparison", 1)
        _w_para(
            doc,
            "Candidates are listed in the ranking order set by the application "
            + ("(target-skill match, then profile completeness, then number of skills found). "
               if has_match else "(profile completeness, then number of skills found). ")
            + "The order is a sorting aid, not a hiring recommendation.",
            italic=True, size=9, color=_W_MUTED, after=4, keep_next=True
        )
        headers = ["Rank", "Candidate", "Profile completeness", "Skills found", "Main career area"]
        widths = [1.3, 4.6, 4.4, 1.9, 4.8]
        if has_match:
            headers.append("Target-skill match")
            widths = [1.2, 3.7, 3.7, 1.6, 3.2, 3.6]
        table = doc.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _w_table_borders(table)
        _w_cell_margins(table)
        for cell, header_text in zip(table.rows[0].cells, headers):
            _w_cell(cell, header_text, bold=True, size=9.5, color="FFFFFF", fill=_W_PRIMARY,
                    align=WD_ALIGN_PARAGRAPH.CENTER if header_text in ("Rank", "Skills found") else None)
        for index, row in enumerate(rows):
            cells = table.add_row().cells
            fill = _W_ZEBRA if index % 2 else None
            _w_cell(cells[0], str(row["rank"]), bold=True, fill=fill, align=WD_ALIGN_PARAGRAPH.CENTER)
            _w_cell(cells[1], row["label"], bold=True, fill=fill)
            score_cell = cells[2]
            _w_cell(score_cell, f"{row['completeness']:.1f}%", bold=True, color=_W_NAVY, fill=fill)
            score_cell.paragraphs[0].add_run("  ")
            _w_bar(score_cell.paragraphs[0], row["completeness"] / 100, 8)
            _w_cell(cells[3], str(row["skill_count"]), fill=fill, align=WD_ALIGN_PARAGRAPH.CENTER)
            _w_cell(cells[4], row["top_area"], fill=fill)
            if has_match:
                match_cell = cells[5]
                _w_cell(match_cell, f"{row['match_pct']:.0f}%  ({len(row['matched'])}/{len(required)})",
                        bold=True, color=_W_NAVY, fill=fill)
                match_cell.paragraphs[0].add_run("  ")
                _w_bar(match_cell.paragraphs[0], row["match_pct"] / 100, 6)
        _w_widths(table, widths)
        _w_repeat_header(table)

        if has_match:
            _w_para(
                doc,
                "Target skills: " + ", ".join(required) + ".  The bracket shows target skills "
                "found out of target skills entered.",
                italic=True, size=9, color=_W_MUTED, before=4
            )

        # Skill frequency
        if combined["skill_freq"]:
            _w_heading(doc, "Skills Overview", 1)
            _w_heading(doc, "Most Common Skills", 2)
            shared = [item for item in combined["skill_freq"] if item[1] >= 2]
            top = shared[:15] if len(shared) >= 3 else combined["skill_freq"][:10]
            table = doc.add_table(rows=1, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            _w_table_borders(table)
            _w_cell_margins(table, 40, 40)
            for cell, header_text in zip(table.rows[0].cells, ["Skill", "CVs with this skill", "Share of CVs"]):
                _w_cell(cell, header_text, bold=True, color="FFFFFF", fill=_W_PRIMARY)
            for index, (skill, number, _holders) in enumerate(top):
                cells = table.add_row().cells
                fill = _W_ZEBRA if index % 2 else None
                share = number / combined["count"]
                _w_cell(cells[0], skill, bold=True, fill=fill)
                _w_cell(cells[1], f"{number} of {combined['count']}", fill=fill)
                bar_cell = cells[2]
                if fill:
                    _w_shade(bar_cell, fill)
                bar_cell.text = ""
                bp = bar_cell.paragraphs[0]
                bp.paragraph_format.space_after = Pt(0)
                _w_bar(bp, share, 14, 9)
                pct = bp.add_run(f"  {share * 100:.0f}%")
                pct.font.size = Pt(9)
                pct.font.color.rgb = RGBColor.from_string(_W_MUTED)
                bar_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _w_widths(table, [6.0, 3.6, 7.4])
            _w_repeat_header(table)

        if combined["career_freq"]:
            _w_heading(doc, "Career Areas", 2)
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            _w_table_borders(table)
            _w_cell_margins(table, 40, 40)
            for cell, header_text in zip(table.rows[0].cells, ["Career area", "Candidates"]):
                _w_cell(cell, header_text, bold=True, color="FFFFFF", fill=_W_PRIMARY)
            for index, (area, number, holders) in enumerate(combined["career_freq"]):
                cells = table.add_row().cells
                fill = _W_ZEBRA if index % 2 else None
                _w_cell(cells[0], f"{area}  ({number})", bold=True, fill=fill)
                _w_cell(cells[1], ", ".join(holders), fill=fill)
            _w_widths(table, [6.0, 11.0])
            _w_repeat_header(table)

        if combined["common_skills"] or combined["unique_skills"]:
            _w_heading(doc, "Shared & Distinctive Skills", 2)
            if combined["common_skills"]:
                _w_para(doc, "Skills present in every CV", bold=True, color=_W_PRIMARY, after=2, keep_next=True)
                _w_para(doc, ", ".join(combined["common_skills"]), after=8)
            if combined["unique_skills"]:
                _w_para(doc, "Skills only one candidate has", bold=True, color=_W_PRIMARY, after=2, keep_next=True)
                for label, skills in combined["unique_skills"].items():
                    p = _w_para(doc, "", after=2)
                    r1 = p.add_run(_clean_office_text(label) + ": ")
                    r1.bold = True
                    p.add_run(_clean_office_text(", ".join(skills)))

    # ---- one section per candidate -------------------------------------------
    NOT_FOUND = "Not detected - please verify in the original CV."
    for position, row in enumerate(rows):
        analysis = row["analysis"]
        statuses = analysis["section_status"]
        found = sum(1 for present in statuses.values() if present)
        missing = [name for name, present in statuses.items() if not present]

        heading = _w_heading(doc, row["label"] if many else f"Candidate Profile: {row['label']}", 1)
        if many or position > 0:
            heading.paragraph_format.page_break_before = True
        if many:
            _w_para(doc, f"Rank {row['rank']} of {combined['count']}  ·  Resume file: {analysis['filename']}",
                    italic=True, size=9.5, color=_W_MUTED, after=6, keep_next=True)

        tiles = [
            (f"{analysis['completeness']:.1f}%", "Profile completeness",
             f"{found} of {len(statuses)} CV sections found"),
            (str(len(analysis["all_skills"])), "Skills found", "Recognized in the CV"),
            (str(len(analysis["roles"])), "Job roles detected", "Titles found in the CV"),
        ]
        if row["match_pct"] is not None:
            tiles.append((f"{row['match_pct']:.0f}%", "Target-skill match",
                          f"{len(row['matched'])} of {len(required)} target skills"))
        _w_tiles(doc, tiles)

        _w_heading(doc, "Contact & File Details", 2)
        info = doc.add_table(rows=3, cols=2)
        info.alignment = WD_TABLE_ALIGNMENT.CENTER
        _w_table_borders(info)
        _w_cell_margins(info)
        for index, (label, value) in enumerate([
            ("Email", analysis["email"] or "Not detected - please verify"),
            ("Phone", analysis["phone"] or "Not detected - please verify"),
            ("Resume file", analysis["filename"]),
        ]):
            _w_cell(info.rows[index].cells[0], label, bold=True, size=9.5, color=_W_PRIMARY, fill=_W_SOFT)
            _w_cell(info.rows[index].cells[1], value, size=10)
        _w_widths(info, [3.6, 13.4])
        _w_keep_table_together(info)

        if row["match_pct"] is not None:
            _w_heading(doc, "Target-Skill Match", 2)
            match_table = doc.add_table(rows=2, cols=2)
            match_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            _w_table_borders(match_table)
            _w_cell_margins(match_table)
            _w_cell(match_table.rows[0].cells[0], "Target skills found", bold=True, color="FFFFFF", fill=_W_PRIMARY)
            _w_cell(match_table.rows[0].cells[1], "Target skills not found", bold=True, color="FFFFFF", fill=_W_PRIMARY)
            _w_cell(match_table.rows[1].cells[0], ", ".join(row["matched"]) or "None", size=10)
            _w_cell(match_table.rows[1].cells[1], ", ".join(row["missing"]) or "None", size=10)
            _w_widths(match_table, [8.5, 8.5])
            _w_keep_table_together(match_table)

        _w_heading(doc, "CV Sections Checked", 2)
        names = list(statuses.items())
        grid_rows = (len(names) + 3) // 4
        grid = doc.add_table(rows=max(grid_rows, 1), cols=4)
        grid.alignment = WD_TABLE_ALIGNMENT.CENTER
        _w_table_borders(grid)
        _w_cell_margins(grid, 50, 50)
        for index in range(max(grid_rows, 1) * 4):
            target = grid.rows[index // 4].cells[index % 4]
            if index < len(names):
                name, present = names[index]
                _w_cell(target, ("✓  " if present else "✗  ") + name, size=9.5,
                        bold=not present, color=_W_GOOD if present else _W_WARN,
                        fill=None if present else "FDF3E1")
            else:
                _w_cell(target, "", size=9.5)
        _w_widths(grid, [4.25] * 4)
        _w_keep_table_together(grid)
        _w_para(doc, "✓ found by the application      ✗ not detected - please verify in the original CV",
                italic=True, size=8.5, color=_W_MUTED, before=3, after=4)

        if missing:
            _w_callout(
                doc, "Please verify manually",
                [("", "The application did not find: " + ", ".join(missing) +
                  ". This may be a limitation of automatic reading rather than missing "
                  "information, so please check the original CV.")],
                fill="FDF3E1", accent=_W_WARN
            )

        _w_heading(doc, "Career Areas & Job Roles", 2)
        roles_table = doc.add_table(rows=2, cols=2)
        roles_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _w_table_borders(roles_table)
        _w_cell_margins(roles_table)
        _w_cell(roles_table.rows[0].cells[0], "Career areas", bold=True, color="FFFFFF", fill=_W_PRIMARY)
        _w_cell(roles_table.rows[0].cells[1], "Job roles detected", bold=True, color="FFFFFF", fill=_W_PRIMARY)
        for cell, items in ((roles_table.rows[1].cells[0], analysis["career_areas"]),
                            (roles_table.rows[1].cells[1], analysis["roles"])):
            cell.text = ""
            first_line = True
            for item in (items or [NOT_FOUND]):
                paragraph = cell.paragraphs[0] if first_line else cell.add_paragraph()
                first_line = False
                paragraph.paragraph_format.space_after = Pt(1)
                run = paragraph.add_run(_clean_office_text(("•  " if items else "") + item))
                run.font.size = Pt(10)
                if not items:
                    run.italic = True
                    run.font.color.rgb = RGBColor.from_string(_W_MUTED)
        _w_widths(roles_table, [8.5, 8.5])
        _w_keep_table_together(roles_table)

        _w_heading(doc, "Skills", 2)
        if analysis["skills_by_category"]:
            skills_table = doc.add_table(rows=0, cols=2)
            skills_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            _w_table_borders(skills_table)
            _w_cell_margins(skills_table)
            for category, skills in analysis["skills_by_category"].items():
                cells = skills_table.add_row().cells
                _w_cell(cells[0], category, bold=True, size=9.5, color=_W_PRIMARY, fill=_W_SOFT)
                _w_cell(cells[1], ", ".join(skills), size=10)
            _w_widths(skills_table, [4.6, 12.4])
            _w_row_cant_split(skills_table)
        else:
            _w_para(doc, "No recognized skills were detected - please verify in the original CV.",
                    italic=True, color=_W_MUTED)

        _w_heading(doc, "Experience", 2)
        entries = experience_entries(analysis["experience"])
        if entries:
            for entry in entries:
                _w_para(doc, entry["heading"], bold=True, size=11, after=0, before=6, keep_next=True)
                if entry["subheading"]:
                    _w_para(doc, entry["subheading"], italic=True, size=9.5, color=_W_MUTED,
                            after=2, keep_next=bool(entry["bullets"]))
                for index, bullet in enumerate(entry["bullets"]):
                    bullet_p = _w_para(doc, bullet, style="List Bullet", after=1)
                    bullet_p.paragraph_format.keep_together = True
        else:
            _w_para(doc, "No structured experience entries were detected - please verify in the original CV.",
                    italic=True, color=_W_MUTED)

        _w_heading(doc, "Education", 2)
        lines = _text_lines(analysis["education"])
        if lines:
            for line in lines:
                _w_para(doc, line, after=1)
        else:
            _w_para(doc, "No dedicated Education information was detected - please verify in the original CV.",
                    italic=True, color=_W_MUTED)

        _w_heading(doc, "Courses & Certifications", 2)
        lines = _text_lines(analysis["certifications"])
        if lines:
            for line in lines:
                _w_para(doc, line, style="List Bullet", after=1)
        else:
            _w_para(doc, "No courses or certifications were detected - please verify in the original CV.",
                    italic=True, color=_W_MUTED)

        _w_heading(doc, "Languages", 2)
        if analysis["languages"]:
            _w_para(doc, ", ".join(analysis["languages"]))
        else:
            _w_para(doc, "No languages were confidently detected - please verify in the original CV.",
                    italic=True, color=_W_MUTED)

        if not many:
            _w_para(doc, "", after=4)
            _w_score_guide(doc, has_match, required)

    doc.core_properties.title = "CV Analysis Report"
    doc.core_properties.author = APP_TITLE
    doc.save(path)



# ------------------------------------------------------------
# EXCEL (.xlsx)
# ------------------------------------------------------------

def _xl(value):
    """Make a value safe for an Excel cell."""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def export_xlsx(path, analyses, combined):
    if not _ensure_xlsx_support():
        raise RuntimeError(
            "Excel export requires openpyxl, and it could not be installed "
            "automatically. Install it manually with:\n"
            "pip install openpyxl"
        )

    import math
    from openpyxl.worksheet.hyperlink import Hyperlink  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.worksheet.table import Table, TableStyleInfo  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.worksheet.pagebreak import Break  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.worksheet.page import PageMargins  # pyright: ignore[reportMissingImports, reportMissingModuleSource]
    from openpyxl.chart.label import DataLabelList  # pyright: ignore[reportMissingImports, reportMissingModuleSource]

    # ---- corporate palette (shared with the Word report) ---------------------
    NAVY, BLUE, SOFT, ZEBRA = "1F2D4D", "2F5597", "EAF0F8", "F6F8FB"
    BORDER_C, TEXT, MUTED = "CBD3E1", "1F2937", "5B667A"
    GOOD_BG, GOOD_TX = "DFF1E6", "1E7B4F"
    NOTE_BG = "FFF6E0"
    FONT = "Calibri"

    thin = Side(style="thin", color=BORDER_C)
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def font(size=10, bold=False, color=TEXT, italic=False, underline=None):
        return Font(name=FONT, size=size, bold=bold, color=color, italic=italic,
                    underline=underline)

    def solid(color):
        return PatternFill("solid", fgColor=color)

    def lines_needed(text, width):
        chars = max(int(width * 1.05), 8)
        return sum(max(1, math.ceil(len(part) / chars)) for part in str(text).split("\n"))

    def height_for(text, width, minimum=18.0):
        return max(minimum, lines_needed(text, width) * 13.5 + 6)

    def set_widths(ws, widths):
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width

    def banner(ws, title, subtitle, last_col):
        for col in range(1, last_col + 1):
            ws.cell(row=1, column=col).fill = solid(NAVY)
            ws.cell(row=2, column=col).fill = solid(BLUE)
        top = ws.cell(row=1, column=1, value=title)
        top.font = font(18, True, "FFFFFF")
        top.alignment = Alignment(vertical="center", indent=1)
        sub = ws.cell(row=2, column=1, value=_xl(subtitle))
        sub.font = font(10, False, "FFFFFF")
        sub.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[1].height = 34
        ws.row_dimensions[2].height = 20

    def header_row(ws, row, labels, start_col=1, height=36):
        for offset, label in enumerate(labels):
            cell = ws.cell(row=row, column=start_col + offset, value=label)
            cell.font = font(10, True, "FFFFFF")
            cell.fill = solid(BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = box
        ws.row_dimensions[row].height = height

    def section_title(ws, row, text, first_col=1, last_col=7):
        cell = ws.cell(row=row, column=first_col, value=text)
        cell.font = font(13, True, NAVY)
        cell.alignment = Alignment(vertical="center")
        for col in range(first_col, last_col + 1):
            ws.cell(row=row, column=col).border = Border(
                bottom=Side(style="medium", color=BLUE))
        ws.row_dimensions[row].height = 24

    def finish_sheet(ws, landscape, tab, title_rows=None, title_cols=None,
                     wide_pages=1, area=None):
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = tab
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.orientation = "landscape" if landscape else "portrait"
        ws.page_setup.paperSize = 9  # A4
        ws.page_setup.fitToWidth = wide_pages
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.7)
        ws.print_options.horizontalCentered = True
        if title_rows:
            ws.print_title_rows = title_rows
        if title_cols:
            ws.print_title_cols = title_cols
        if area:
            ws.print_area = area
        ws.oddFooter.left.text = APP_TITLE
        ws.oddFooter.left.size = 8
        ws.oddFooter.center.text = "Page &P of &N"
        ws.oddFooter.center.size = 8
        ws.oddFooter.right.text = combined["generated"].strftime("%d %b %Y")
        ws.oddFooter.right.size = 8

    rows = combined["rows"]
    count = len(rows)
    required = combined["required"]
    has_match = bool(required)
    generated = combined["generated"].strftime("%d %B %Y, %H:%M")
    n_checks = max((len(r["analysis"]["section_status"]) for r in rows), default=8)

    def missing_sections(analysis):
        return [name for name, present in analysis["section_status"].items() if not present]

    def verify_text(analysis):
        missing = missing_sections(analysis)
        if not missing:
            return "None - all CV sections were found"
        return "Not detected: " + ", ".join(missing) + ". Please check the original CV."

    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "Summary"
    comparison = wb.create_sheet("Comparison")
    matrix = wb.create_sheet("Skill Matrix")
    skills_sheet = wb.create_sheet("Skills Detail")
    details = wb.create_sheet("Candidate Details")

    # =====================================================================
    # Sheet: Comparison
    # =====================================================================
    cols = [
        ("Rank", 7), ("Candidate", 26),
        ("Profile completeness\n(share of CV sections found)", 20),
        ("Skills found\n(count)", 12), ("Job roles found\n(count)", 12),
        ("Main career area", 24),
    ]
    if has_match:
        cols += [("Target-skill match\n(share of target skills found)", 20),
                 ("Target skills found", 30), ("Target skills not found", 30)]
    cols += [("Skills detected", 72), ("Items to verify manually", 32),
             ("Email", 30), ("Phone", 18), ("Resume file", 34)]
    last_col = len(cols)

    banner(comparison, "Candidate Comparison",
           f"{count} candidate{'s' if count != 1 else ''}  ·  Generated {generated}  ·  "
           "Listed in the ranking order set by the application", last_col)
    note = comparison.cell(
        row=3, column=1,
        value="Scores are decision-support information only. Use the filter arrows in the "
              "header row to sort or filter. The Summary sheet explains what each score means.")
    note.font = font(9, False, MUTED, True)
    note.alignment = Alignment(vertical="center", indent=1)
    comparison.row_dimensions[3].height = 18

    header_row(comparison, 4, [c[0] for c in cols], height=46)
    set_widths(comparison, [c[1] for c in cols])
    first = 5
    last = first + count - 1
    col_index = {c[0].split("\n")[0]: i for i, c in enumerate(cols, start=1)}
    center_cols = {1, 3, 4, 5} | ({7} if has_match else set())
    wrap_cols = {2, 6} | ({8, 9} if has_match else set()) | {last_col - 4, last_col - 3}

    for offset, row in enumerate(rows):
        r = first + offset
        analysis = row["analysis"]
        values = [
            row["rank"], row["label"], row["completeness"] / 100, row["skill_count"],
            row["role_count"], row["top_area"],
        ]
        if has_match:
            values += [row["match_pct"] / 100,
                       ", ".join(row["matched"]) or "-",
                       ", ".join(row["missing"]) or "-"]
        values += [", ".join(row["skills"]) or "None detected", verify_text(analysis),
                   row["email"] or "Not detected", row["phone"] or "Not detected",
                   row["filename"]]
        tallest = 20.0
        for col, value in enumerate(values, start=1):
            cell = comparison.cell(row=r, column=col, value=_xl(value))
            cell.font = font(10, bold=(col in (1, 2)))
            cell.border = box
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if col in center_cols else "left",
                wrap_text=col in wrap_cols or col == last_col - 2 or col == last_col,
            )
            if offset % 2:
                cell.fill = solid(ZEBRA)
            if isinstance(value, str) and col in wrap_cols | {last_col}:
                tallest = max(tallest, height_for(value, cols[col - 1][1]))
        comparison.cell(row=r, column=3).number_format = "0.0%"
        if has_match:
            comparison.cell(row=r, column=7).number_format = "0%"
        comparison.row_dimensions[r].height = min(tallest, 160)

    if count:
        comparison.conditional_formatting.add(
            f"C{first}:C{last}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                        color="8FB0DE", showValue=True))
        comparison.conditional_formatting.add(
            f"D{first}:D{last}",
            ColorScaleRule(start_type="min", start_color="FFFFFF",
                           end_type="max", end_color="BCD3EE"))
        if has_match:
            comparison.conditional_formatting.add(
                f"G{first}:G{last}",
                DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                            color="8FD0B0", showValue=True))
        comparison.auto_filter.ref = f"A4:{get_column_letter(last_col)}{last}"
    comparison.freeze_panes = "C5"
    finish_sheet(comparison, True, BLUE, title_rows="4:4", title_cols="A:B",
                 wide_pages=1 if last_col <= 12 else 2)

    # =====================================================================
    # Sheet: Skill Matrix
    # =====================================================================
    skill_names = [skill for skill, _n, _h in combined["skill_freq"]]
    m_last_col = max(2 + len(skill_names), 8)
    banner(matrix, "Skill Matrix",
           f"Which recognized skills were detected in which CV  ·  Generated {generated}",
           m_last_col)

    legend_tick = matrix.cell(row=3, column=2, value="✓")
    legend_tick.font = font(10, True, GOOD_TX)
    legend_tick.fill = solid(GOOD_BG)
    legend_tick.alignment = Alignment(horizontal="center")
    legend_tick.border = box
    matrix.cell(row=3, column=3, value="Skill detected in the CV").font = font(9, False, TEXT)
    legend_dash = matrix.cell(row=4, column=2, value="–")
    legend_dash.font = font(10, False, "8A93A8")
    legend_dash.fill = solid("F1F3F7")
    legend_dash.alignment = Alignment(horizontal="center")
    legend_dash.border = box
    matrix.cell(row=4, column=3, value=(
        "Not detected (the CV may use different wording - verify manually if the skill is important)"
    )).font = font(9, False, TEXT)
    matrix.cell(row=3, column=1, value="How to read").font = font(9, True, NAVY)
    matrix.cell(row=3, column=1).alignment = Alignment(horizontal="right", indent=1)
    matrix.row_dimensions[3].height = 18
    matrix.row_dimensions[4].height = 18

    m_header = 6
    header_row(matrix, m_header, ["Candidate", "Skills detected"] + skill_names)
    matrix.cell(row=m_header, column=1).alignment = Alignment(
        horizontal="left", vertical="center", indent=1)
    longest = max([len(s) for s in skill_names] or [8])
    matrix.row_dimensions[m_header].height = min(190, 24 + longest * 6.0)
    for col in range(3, len(skill_names) + 3):
        matrix.cell(row=m_header, column=col).alignment = Alignment(
            horizontal="center", vertical="bottom", text_rotation=90)

    for offset, row in enumerate(rows):
        r = m_header + 1 + offset
        name_cell = matrix.cell(row=r, column=1, value=_xl(row["label"]))
        name_cell.font = font(10, True)
        name_cell.border = box
        name_cell.alignment = Alignment(vertical="center", indent=1)
        total_cell = matrix.cell(row=r, column=2, value=row["skill_count"])
        total_cell.font = font(10, True, NAVY)
        total_cell.fill = solid(SOFT)
        total_cell.border = box
        total_cell.alignment = Alignment(horizontal="center", vertical="center")
        owned = set(row["skills"])
        for c_index, skill in enumerate(skill_names, start=3):
            cell = matrix.cell(row=r, column=c_index)
            cell.border = box
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if skill in owned:
                cell.value = "✓"
                cell.font = font(10, True, GOOD_TX)
                cell.fill = solid(GOOD_BG)
            else:
                cell.value = "–"
                cell.font = font(9, False, "8A93A8")
                cell.fill = solid("F1F3F7")
        matrix.row_dimensions[r].height = 20

    total_row = m_header + 1 + count
    label = matrix.cell(row=total_row, column=1, value="Candidates with this skill")
    label.font = font(10, True, "FFFFFF")
    label.fill = solid(NAVY)
    label.border = box
    label.alignment = Alignment(vertical="center", indent=1)
    blank = matrix.cell(row=total_row, column=2)
    blank.fill = solid(NAVY)
    blank.border = box
    holders_by_skill = {skill: number for skill, number, _h in combined["skill_freq"]}
    for c_index, skill in enumerate(skill_names, start=3):
        cell = matrix.cell(row=total_row, column=c_index, value=holders_by_skill[skill])
        cell.font = font(10, True, "FFFFFF")
        cell.fill = solid(NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = box
    matrix.row_dimensions[total_row].height = 20
    matrix.column_dimensions["A"].width = 30
    matrix.column_dimensions["B"].width = 11
    for c_index in range(3, len(skill_names) + 3):
        matrix.column_dimensions[get_column_letter(c_index)].width = 4.6
    matrix.freeze_panes = f"C{m_header + 1}"
    finish_sheet(matrix, True, BLUE, title_rows=f"{m_header}:{m_header}", title_cols="A:B",
                 wide_pages=max(1, math.ceil(len(skill_names) / 40)))

    # =====================================================================
    # Sheet: Skills Detail
    # =====================================================================
    banner(skills_sheet, "Skills Detail",
           f"Skills grouped by how many CVs contain them  ·  Generated {generated}", 11)
    intro = skills_sheet.cell(
        row=3, column=1,
        value="Left: every recognized skill and the candidates who have it. "
              "Right: the skills found in each CV, grouped by category. "
              "Use the filter arrows to narrow either list.")
    intro.font = font(9, False, MUTED, True)
    intro.alignment = Alignment(vertical="center", indent=1)

    skill_categories = {}
    for row in rows:
        for category, skills in row["analysis"]["skills_by_category"].items():
            for skill in skills:
                seen = skill_categories.setdefault(skill, [])
                if category not in seen:
                    seen.append(category)

    set_widths(skills_sheet, [28, 30, 14, 13, 22, 56, 3, 28, 30, 30, 3])
    section_title(skills_sheet, 5, "Skill overview", 1, 6)
    section_title(skills_sheet, 5, "Skills listed by candidate", 8, 10)

    if skill_names:
        header_row(skills_sheet, 6, ["Skill", "Skill category", "CVs with this skill",
                                     "Share of CVs", "How common", "Candidates"], 1, 34)
        for offset, (skill, number, holders) in enumerate(combined["skill_freq"]):
            r = 7 + offset
            if count > 1 and number == count:
                common = "Found in all CVs"
            elif count > 1 and number == 1:
                common = "Found in 1 CV only"
            elif count > 1:
                common = f"Found in {number} of {count} CVs"
            else:
                common = "-"
            values = [skill, "; ".join(skill_categories.get(skill, [])) or "-",
                      number, number / count if count else 0, common, ", ".join(holders)]
            for col, value in enumerate(values, start=1):
                cell = skills_sheet.cell(row=r, column=col, value=_xl(value))
                cell.font = font(10, bold=(col == 1))
                cell.border = box
                cell.alignment = Alignment(
                    vertical="top", wrap_text=col in (2, 6),
                    horizontal="center" if col in (3, 4) else "left")
                if offset % 2:
                    cell.fill = solid(ZEBRA)
            skills_sheet.cell(row=r, column=4).number_format = "0%"
            skills_sheet.row_dimensions[r].height = max(
                height_for(values[5], 56), height_for(values[1], 30), 18)
        overview_last = 6 + len(combined["skill_freq"])
        skills_sheet.conditional_formatting.add(
            f"D7:D{overview_last}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                        color="8FB0DE", showValue=True))
        table = Table(displayName="SkillOverview", ref=f"A6:F{overview_last}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        skills_sheet.add_table(table)

        header_row(skills_sheet, 6, ["Candidate", "Skill category", "Skill"], 8, 34)
        line = 7
        for row in rows:
            for category, skills in row["analysis"]["skills_by_category"].items():
                for skill in skills:
                    for col, value in enumerate((row["label"], category, skill), start=8):
                        cell = skills_sheet.cell(row=line, column=col, value=_xl(value))
                        cell.font = font(10, bold=(col == 8))
                        cell.border = box
                        cell.alignment = Alignment(vertical="center")
                        if (line - 7) % 2:
                            cell.fill = solid(ZEBRA)
                    line += 1
        if line > 7:
            table2 = Table(displayName="SkillsByCandidate", ref=f"H6:J{line - 1}")
            table2.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
            skills_sheet.add_table(table2)
        skills_sheet.freeze_panes = "A7"
    else:
        skills_sheet.cell(row=7, column=1, value=(
            "No recognized skills were detected in the CVs. Please check the original "
            "documents manually.")).font = font(10, False, MUTED, True)
    finish_sheet(skills_sheet, True, BLUE, title_rows="6:6", wide_pages=1)

    # =====================================================================
    # Sheet: Candidate Details (one profile block per candidate)
    # =====================================================================
    banner(details, "Candidate Details",
           f"One profile per candidate, in ranking order  ·  Generated {generated}", 2)
    set_widths(details, [30, 100])
    details.cell(row=3, column=1, value=(
        "'Not detected' means the application did not find the information in the CV text "
        "- please check the original CV before drawing any conclusion.")).font = font(
        9, False, MUTED, True)
    details.cell(row=3, column=1).alignment = Alignment(vertical="center", indent=1)
    details.row_dimensions[3].height = 18

    cursor = 5
    NOT_FOUND = "Not detected - please check the original CV"

    def put(label_text, value_text, label_fill=SOFT, bold_value=False, color=TEXT):
        nonlocal cursor
        label_cell = details.cell(row=cursor, column=1, value=label_text)
        label_cell.font = font(10, True, NAVY)
        label_cell.fill = solid(label_fill)
        label_cell.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        label_cell.border = box
        value_cell = details.cell(row=cursor, column=2, value=_xl(value_text))
        value_cell.font = font(10, bold_value, color)
        value_cell.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        value_cell.border = box
        details.row_dimensions[cursor].height = max(
            height_for(value_text, 100), height_for(label_text, 30))
        cursor += 1

    def group_title(text):
        nonlocal cursor
        cell = details.cell(row=cursor, column=1, value=text.upper())
        cell.font = font(9, True, BLUE)
        cell.alignment = Alignment(vertical="bottom")
        details.row_dimensions[cursor].height = 22
        cursor += 1

    for position, row in enumerate(rows):
        analysis = row["analysis"]
        for col in (1, 2):
            details.cell(row=cursor, column=col).fill = solid(NAVY)
        title = details.cell(
            row=cursor, column=1,
            value=_xl(f"{row['label']}   ·   Rank {row['rank']} of {count}"))
        title.font = font(14, True, "FFFFFF")
        title.alignment = Alignment(vertical="center", indent=1)
        details.row_dimensions[cursor].height = 30
        cursor += 1

        group_title("Contact and file")
        put("Email", analysis["email"] or NOT_FOUND)
        put("Phone", analysis["phone"] or NOT_FOUND)
        put("Resume file", analysis["filename"])

        group_title("Scores (see the Summary sheet for what they mean)")
        found = sum(1 for present in analysis["section_status"].values() if present)
        put("Profile completeness",
            f"{analysis['completeness']:.1f}%  ({found} of {len(analysis['section_status'])} "
            "standard CV sections found)", bold_value=True)
        put("Skills found", f"{len(analysis['all_skills'])} recognized skills")
        if row["match_pct"] is not None:
            put("Target-skill match",
                f"{row['match_pct']:.0f}%  ({len(row['matched'])} of {len(required)} target "
                "skills found)", bold_value=True)
            put("Target skills found", ", ".join(row["matched"]) or "None")
            put("Target skills not found", ", ".join(row["missing"]) or "None")
        sections_text = "     ".join(
            ("✓ " if present else "✗ ") + name
            for name, present in analysis["section_status"].items())
        put("CV sections checked", sections_text)
        put("Items to verify manually", verify_text(analysis),
            color=TEXT if not missing_sections(analysis) else "8A4B00")

        group_title("Career profile")
        put("Career areas", "\n".join(analysis["career_areas"]) or NOT_FOUND)
        put("Job roles detected", "\n".join(analysis["roles"]) or NOT_FOUND)

        group_title("Skills by category")
        if analysis["skills_by_category"]:
            for category, skills in analysis["skills_by_category"].items():
                put(category, ", ".join(skills))
        else:
            put("Skills", "No recognized skills were detected - please check the original CV")

        group_title("Experience")
        entries = experience_entries(analysis["experience"])
        if entries:
            for number, entry in enumerate(entries, start=1):
                block = [entry["heading"]]
                if entry["subheading"]:
                    block.append(entry["subheading"])
                block.extend("• " + b for b in entry["bullets"])
                put(f"Experience {number}", "\n".join(block))
        else:
            put("Experience", "No structured experience entries were detected - "
                              "please check the original CV")

        group_title("Education, courses and languages")
        put("Education", analysis["education"] or NOT_FOUND)
        put("Courses and certifications", analysis["certifications"] or NOT_FOUND)
        put("Languages", ", ".join(analysis["languages"]) or NOT_FOUND)

        cursor += 1
        if position < len(rows) - 1:
            details.row_breaks.append(Break(id=cursor - 1))
            cursor += 1
    details.freeze_panes = "A5"
    finish_sheet(details, False, BLUE)

    # =====================================================================
    # Sheet: Summary (management dashboard)
    # =====================================================================
    set_widths(summary, [30, 14, 14, 3, 30, 14, 14])
    banner(summary, "CV Analysis Report",
           f"HR decision-support summary  ·  {count} CV{'s' if count != 1 else ''} analyzed  ·  "
           f"Generated {generated}", 7)
    summary.row_dimensions[3].height = 8

    white_edge = Side(style="thick", color="FFFFFF")

    def card(top, first_col, last_col, caption, value, number_format, hint):
        for r in (top, top + 1, top + 2):
            for c in range(first_col, last_col + 1):
                cell = summary.cell(row=r, column=c)
                cell.fill = solid(SOFT)
                cell.border = Border(left=white_edge, right=white_edge)
            if last_col > first_col:
                summary.merge_cells(start_row=r, start_column=first_col,
                                    end_row=r, end_column=last_col)
        cap = summary.cell(row=top, column=first_col, value=caption)
        cap.font = font(9, True, MUTED)
        cap.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        val = summary.cell(row=top + 1, column=first_col, value=_xl(value))
        val.font = font(24 if len(str(value)) < 14 else 14, True, BLUE)
        val.number_format = number_format
        val.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        hnt = summary.cell(row=top + 2, column=first_col, value=_xl(hint))
        hnt.font = font(8.5, False, MUTED, True)
        hnt.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    def set_card_heights(top):
        summary.row_dimensions[top].height = 24
        summary.row_dimensions[top + 1].height = 40
        summary.row_dimensions[top + 2].height = 30

    card(4, 1, 1, "CVs ANALYZED", count, "0", "Total CVs in this report")
    card(4, 2, 3, "AVERAGE PROFILE COMPLETENESS", combined["avg_completeness"] / 100, "0.0%",
         f"Share of the {n_checks} standard CV sections found")
    card(4, 5, 5, "UNIQUE SKILLS RECOGNIZED", combined["total_unique_skills"], "0",
         "Different skills found across all CVs")
    card(4, 6, 7, "AVERAGE SKILLS PER CV", combined["avg_skills"], "0.0",
         "Recognized skills in a typical CV")
    set_card_heights(4)
    next_row = 7
    if has_match:
        full = sum(1 for r in rows if r["match_pct"] is not None and r["match_pct"] >= 100)
        card(7, 1, 1, "TARGET SKILLS ENTERED", len(required), "0", "Skills used for the match")
        card(7, 2, 3, "HIGHEST TARGET-SKILL MATCH", rows[0]["match_pct"] / 100, "0%",
             rows[0]["label"])
        card(7, 5, 5, "CVS WITH ALL TARGET SKILLS", full, "0", "Every target skill was found")
        card(7, 6, 7, "TOP RANKED CANDIDATE", rows[0]["label"], "General",
             "First in the application's ranking")
        set_card_heights(7)
        next_row = 10
    summary.row_dimensions[next_row].height = 10

    row_cursor = next_row + 1
    section_title(summary, row_cursor, "Key findings")
    row_cursor += 1
    for sentence in combined_summary_sentences(combined)[:6]:
        summary.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=7)
        cell = summary.cell(row=row_cursor, column=1, value=_xl("•  " + sentence))
        cell.font = font(10)
        cell.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        summary.row_dimensions[row_cursor].height = height_for(sentence, 100, 20)
        row_cursor += 1

    row_cursor += 1
    section_title(summary, row_cursor, "How to read the scores")
    row_cursor += 1
    ranking_text = (
        "Candidates are ordered by the application: highest target-skill match first, then "
        "profile completeness, then number of skills found."
        if has_match else
        "Candidates are ordered by the application: highest profile completeness first, then "
        "number of skills found."
    ) + " This is a sorting aid only - it is not a hiring recommendation."
    guide = [
        ("Profile completeness",
         f"How many of the {n_checks} standard CV sections the application found: Name, Email, "
         "Phone, Education, Experience, Skills, Certifications/Courses and Languages. Each "
         "section found adds an equal share. It shows how complete the CV is. It does not "
         "measure the quality of the candidate or how well they would perform in a role."),
        ("Skills found",
         "The number of skills from the application's built-in skills list that appear in the "
         "CV. Skills outside the list, or described in unusual wording, are not counted. A "
         "higher number does not make a candidate better suited to a role."),
    ]
    if has_match:
        guide.append((
            "Target-skill match",
            "Shown because target skills were entered: " + ", ".join(required) + ". It is the "
            "share of those skills found anywhere in the CV (a keyword check). It does not show "
            "how well, how recently or how often a skill was used."))
    guide += [
        ("Ranking order", ranking_text),
        ("Manual verification",
         "The application reads each CV automatically. Information held in images, unusual "
         "layouts or unusual wording can be missed. 'Not detected' means the application did "
         "not find it - not that the candidate lacks it. Please check the original CV."),
        ("Decision support",
         "This report supports human review and does not replace it. No candidate has been "
         "accepted or rejected by the application."),
    ]
    for label_text, text in guide:
        label_cell = summary.cell(row=row_cursor, column=1, value=label_text)
        label_cell.font = font(10, True, NAVY)
        label_cell.fill = solid(SOFT)
        label_cell.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        label_cell.border = box
        summary.merge_cells(start_row=row_cursor, start_column=2, end_row=row_cursor, end_column=7)
        text_cell = summary.cell(row=row_cursor, column=2, value=text)
        text_cell.font = font(10)
        text_cell.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        for c in range(2, 8):
            summary.cell(row=row_cursor, column=c).border = box
        summary.row_dimensions[row_cursor].height = height_for(text, 84, 24)
        row_cursor += 1

    row_cursor += 1
    section_title(summary, row_cursor, "What is in this workbook")
    row_cursor += 1
    contents = [
        ("Comparison", "All candidates side by side with scores, skills and items to verify."),
        ("Skill Matrix", "A grid showing which skills were detected in which CV."),
        ("Skills Detail", "Skills grouped by how many CVs contain them, and skills per candidate."),
        ("Candidate Details", "One profile per candidate: contact, experience, education and more."),
    ]
    for sheet_name, description in contents:
        link = summary.cell(row=row_cursor, column=1, value=sheet_name)
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{sheet_name}'!A1",
                                   display=sheet_name)
        link.font = font(10, True, BLUE, underline="single")
        link.alignment = Alignment(indent=1, vertical="center")
        summary.merge_cells(start_row=row_cursor, start_column=2, end_row=row_cursor, end_column=7)
        desc = summary.cell(row=row_cursor, column=2, value=description)
        desc.font = font(10, False, TEXT)
        desc.alignment = Alignment(vertical="center", indent=1)
        summary.row_dimensions[row_cursor].height = 20
        row_cursor += 1

    row_cursor += 1
    chart_row = None
    if combined["skill_freq"] and count > 1:
        summary.row_breaks.append(Break(id=row_cursor - 1))
        section_title(summary, row_cursor, "Most common skills")
        chart_row = row_cursor + 1
        row_cursor = chart_row + 17

    skill_title = row_cursor
    section_title(summary, skill_title, "Skills across all CVs", 1, 3)
    section_title(summary, skill_title, "Career areas", 5, 7)
    header_row(summary, skill_title + 1, ["Skill", "CVs", "Share of CVs"], 1, 24)
    header_row(summary, skill_title + 1, ["Career area", "CVs", "Share of CVs"], 5, 24)
    skill_first = skill_title + 2
    for offset, (skill, number, _holders) in enumerate(combined["skill_freq"]):
        r = skill_first + offset
        values = [(1, skill), (2, number), (3, number / count if count else 0)]
        for col, value in values:
            cell = summary.cell(row=r, column=col, value=_xl(value))
            cell.font = font(10, bold=(col == 1))
            cell.border = box
            cell.alignment = Alignment(horizontal="left" if col == 1 else "center")
            if offset % 2:
                cell.fill = solid(ZEBRA)
        summary.cell(row=r, column=3).number_format = "0%"
    skill_last = skill_first + len(combined["skill_freq"]) - 1
    area_first = skill_first
    for offset, (area, number, _holders) in enumerate(combined["career_freq"]):
        r = area_first + offset
        values = [(5, area), (6, number), (7, number / count if count else 0)]
        for col, value in values:
            cell = summary.cell(row=r, column=col, value=_xl(value))
            cell.font = font(10, bold=(col == 5))
            cell.border = box
            cell.alignment = Alignment(horizontal="left" if col == 5 else "center",
                                       wrap_text=(col == 5), vertical="center")
            if offset % 2:
                cell.fill = solid(ZEBRA)
        summary.cell(row=r, column=7).number_format = "0%"
    area_last = area_first + len(combined["career_freq"]) - 1
    if combined["skill_freq"]:
        summary.conditional_formatting.add(
            f"C{skill_first}:C{skill_last}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                        color="8FB0DE", showValue=True))
    if combined["career_freq"]:
        summary.conditional_formatting.add(
            f"G{area_first}:G{area_last}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                        color="8FB0DE", showValue=True))

    if chart_row is not None:
        chart = XlBarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = None
        top_n = min(10, len(combined["skill_freq"]))
        data = XlReference(summary, min_col=2, min_row=skill_first - 1,
                           max_row=skill_first - 1 + top_n)
        cats = XlReference(summary, min_col=1, min_row=skill_first,
                           max_row=skill_first - 1 + top_n)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.legend = None
        chart.height = 8.2
        chart.width = 22
        chart.gapWidth = 60
        chart.x_axis.scaling.orientation = "maxMin"
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.y_axis.majorUnit = 1
        chart.y_axis.title = "Number of CVs"
        chart.y_axis.majorGridlines = None
        chart.series[0].graphicalProperties.solidFill = BLUE
        chart.series[0].graphicalProperties.line.solidFill = BLUE
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.showCatName = False
        chart.dataLabels.showLegendKey = False
        summary.add_chart(chart, f"A{chart_row}")

    summary_last = max(skill_last if combined["skill_freq"] else skill_first,
                       area_last if combined["career_freq"] else skill_first)
    finish_sheet(summary, False, NAVY, area=f"A1:G{summary_last}")

    wb.active = 0
    wb.properties.title = "CV Analysis Report"
    wb.properties.creator = APP_TITLE
    wb.save(path)


EXPORTERS = {
    "docx": export_docx,
    "xlsx": export_xlsx,
    "txt": export_text,
}


# ============================================================
# MODERN UI TOOLKIT
# ============================================================
