"""
╔══════════════════════════════════════════════════════════════════════════════╗
║          CV MAKER — CORPORATES GUIDE  •  Your Guide to Success              ║
║  Generates professional PDFs matching the official Corporates Guide format  ║
╚══════════════════════════════════════════════════════════════════════════════╝

FEATURES
  • Manual entry (5-tab form covering all CV sections)
  • Upload existing DOCX or PDF to auto-populate fields
  • One-click PDF export with Corporates Guide logo & branding
  • Built-in EXE packaging guide (PyInstaller)

EXACT COLOUR CODES extracted from the reference DOCX
  Header bg  : #1B3A6B
  Name text  : #FFFFFF  (26pt bold)
  Title text : #BDD7EE  (13pt)
  Tags text  : #8DB4D9  (9pt)
  Section hdr: #1B3A6B  (11pt bold)  + bottom border  #1B3A6B
  Job title  : #1B3A6B  (11pt bold)
  Company    : #2E6DB4  (bold) + #555555 (date italic)
  Body text  : #333333
  Bullets    : #333333
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, sys, threading, re, shutil
from pathlib import Path

# ── PDF ───────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether,
    PageBreak,
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas as rl_canvas

from PIL import Image as PILImage, ImageTk

# ── DOCX parser ───────────────────────────────────────────────────────────────
try:
    from docx import Document as DocxDoc

    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ── PDF text parser ───────────────────────────────────────────────────────────
try:
    import pdfplumber

    PDFPLUMBER_OK = True
except ImportError:
    PDFPLUMBER_OK = False


# ─────────────────────────────────────────────────────────────────────────────
#  PATHS
# ─────────────────────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(BASE, "CG Logo.jpg")


# ─────────────────────────────────────────────────────────────────────────────
#  EXACT COLOURS (from DOCX analysis)
# ─────────────────────────────────────────────────────────────────────────────
# PDF colours
C_HEADER_BG = colors.white
C_WHITE = colors.white
C_NAME = colors.HexColor("#1B3A6B")
C_TITLE_SUB = colors.HexColor("#2E6DB4")
C_TAG = colors.HexColor("#555555")
C_SEC_HDR = colors.HexColor("#1B3A6B")  # Section heading text
C_SEC_BORDER = colors.HexColor("#1B3A6B")  # Section heading border
C_JOB_TITLE = colors.HexColor("#1B3A6B")  # Job title
C_COMPANY = colors.HexColor("#2E6DB4")  # Company name
C_DATE = colors.HexColor("#555555")  # Date / separator
C_BODY = colors.HexColor("#333333")  # Body text
C_DIVIDER = colors.HexColor("#D0D7DE")  # Light divider
C_FOOTER = colors.HexColor("#555555")

# GUI colours
G_BG = "#0d1117"
G_CARD = "#161b22"
G_BORDER = "#30363d"
G_TOPBAR = "#1B3A6B"
G_ACCENT = "#2E6DB4"
G_RED = "#e94560"
G_TEXT = "#e6edf3"
G_MUTED = "#8b949e"
G_INPUT = "#0d1117"
G_SUCCESS = "#3fb950"
G_WARN = "#d29922"


# ─────────────────────────────────────────────────────────────────────────────
#  CUSTOM FLOWABLE  — Section Heading (matches DOCX style exactly)
#  Bold navy text + full-width navy bottom border line
# ─────────────────────────────────────────────────────────────────────────────
class SectionHeading(Flowable):
    """Replicates the DOCX section heading: bold navy text + navy underline."""

    H = 20

    def __init__(self, text, avail_width=0):
        super().__init__()
        self.text = text.upper()
        self._w = avail_width
        self.height = self.H

    def draw(self):
        c = self.canv
        # Text
        c.setFillColor(C_SEC_HDR)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(0, 5, self.text)
        # Bottom border line (matches DOCX 1B3A6B border)
        c.setStrokeColor(C_SEC_BORDER)
        c.setLineWidth(1.5)
        c.line(0, 1, self._w, 1)

    def wrap(self, aW, aH):
        self._w = aW
        return aW, self.H


# ─────────────────────────────────────────────────────────────────────────────
#  PDF GENERATOR  (exact replica of reference CV)
# ─────────────────────────────────────────────────────────────────────────────
class CVGenerator:
    PAGE_W, PAGE_H = A4
    L_MAR = R_MAR = 1.8 * cm
    T_MAR = 2.6 * cm  # leaves room for the navy header band
    B_MAR = 1.3 * cm
    HEADER_H = 2.0 * cm

    def __init__(self, data: dict, outfile: str, logo_path: str = LOGO):
        self.d = data
        self.out = outfile
        self.logo = logo_path if (logo_path and os.path.exists(logo_path)) else None
        self.CW = self.PAGE_W - self.L_MAR - self.R_MAR
        self._mk_styles()

    # ── Styles ────────────────────────────────────────────────────────────────
    def _mk_styles(self):
        base = getSampleStyleSheet()

        def S(n, **k):
            return ParagraphStyle(n, parent=base["Normal"], **k)

        self.st = {
            # ── Header band ──
            "name": S(
                "Name",
                fontSize=26,
                fontName="Helvetica-Bold",
                textColor=C_NAME,
                leading=30,
                spaceAfter=2,
            ),
            "title": S(
                "Title",
                fontSize=13,
                fontName="Helvetica",
                textColor=C_TITLE_SUB,
                leading=16,
                spaceAfter=3,
            ),
            "tagline": S(
                "Tag", fontSize=9, fontName="Helvetica", textColor=C_TAG, leading=13
            ),
            # ── Body ──
            "body": S("Body", fontSize=9, leading=14, textColor=C_BODY, spaceAfter=2),
            "bodyJ": S(
                "BodyJ",
                fontSize=9,
                leading=14,
                textColor=C_BODY,
                alignment=TA_JUSTIFY,
                spaceAfter=3,
            ),
            # ── Job title (bold navy, 11pt, matches DOCX sz=22 = 11pt) ──
            "jobtitle": S(
                "JT",
                fontSize=11,
                fontName="Helvetica-Bold",
                textColor=C_JOB_TITLE,
                leading=15,
                spaceBefore=8,
                spaceAfter=2,
            ),
            # ── Company name (bold blue) + date (italic grey) ──
            "company": S(
                "Co",
                fontSize=9,
                fontName="Helvetica",
                textColor=C_COMPANY,
                leading=14,
                spaceAfter=2,
            ),
            # ── Bullet items ──
            "bullet": S(
                "Bul",
                fontSize=9,
                leading=13.5,
                textColor=C_BODY,
                leftIndent=14,
                firstLineIndent=-10,
                spaceBefore=2,
                spaceAfter=2,
            ),
            # ── Competency lines (bold category + normal items) ──
            "compet": S("Cmp", fontSize=9, leading=14, textColor=C_BODY, spaceAfter=3),
        }

    # ── Helpers ───────────────────────────────────────────────────────────────
    def sp(self, h=4):
        return Spacer(1, h)

    def sec(self, title):
        """Section heading + small spacer."""
        return [self.sp(10), SectionHeading(title, self.CW), self.sp(5)]

    def bul(self, text):
        t = re.sub(r"^[\-\•\*]\s*", "", text.strip())
        return Paragraph(f"<bullet>\u2022</bullet> {t}", self.st["bullet"])

    # ── Page canvas (draws the navy header band + logo on every page) ─────────
    def _draw_page(self, c, doc, first=True):
        W, H = self.PAGE_W, self.PAGE_H
        c.saveState()

        if first:
            hh = self.HEADER_H
            # Navy band
            c.setFillColor(C_HEADER_BG)
            c.rect(0, H - hh, W, hh, fill=1, stroke=0)

            # Logo in right portion of header
            if self.logo:
                try:
                    lw, lh = 4.0 * cm, 4.55 * cm
                    c.drawImage(
                        self.logo,
                        W - lw - 1.0 * cm,
                        H - hh + (hh - lh) / 2,
                        width=lw,
                        height=lh,
                        preserveAspectRatio=True,
                        mask="auto",
                    )
                except Exception:
                    pass

            # Thin accent line below header
            c.setFillColor(colors.HexColor("#2E6DB4"))
            c.rect(0, H - hh - 2, W, 2, fill=1, stroke=0)
        else:
            # Compact header for page 2+
            hh = 1.1 * cm
            c.setFillColor(C_HEADER_BG)
            c.rect(0, H - hh, W, hh, fill=1, stroke=0)
            c.setFillColor(colors.HexColor("#2E6DB4"))
            c.rect(0, H - hh - 2, W, 2, fill=1, stroke=0)
            # Candidate name in compact header
            name = self.d.get("full_name", "")
            if name:
                c.setFont("Helvetica-Bold", 9)
                c.setFillColor(C_WHITE)
                c.drawString(self.L_MAR, H - hh + 4, name)

        # Footer
        c.setFont("Helvetica", 7)
        c.setFillColor(C_FOOTER)
        c.drawCentredString(
            W / 2,
            0.5 * cm,
            "Corporates Guide  •  Your Guide to Success  •  CV generated with CV Maker",
        )
        c.restoreState()

    def _page1(self, c, doc):
        self._draw_page(c, doc, first=True)

    def _page_n(self, c, doc):
        self._draw_page(c, doc, first=False)

    # ── Section builders ──────────────────────────────────────────────────────
    def _hdr_text(self):
        """Text lines that sit inside the navy header band."""
        out = []
        out.append(Paragraph(self.d.get("full_name", "Your Name"), self.st["name"]))
        if v := self.d.get("job_title", "").strip():
            out.append(Paragraph(v, self.st["title"]))
        if v := self.d.get("tagline", "").strip():
            out.append(Paragraph(v, self.st["tagline"]))
        return out

    def _contact(self):
        """Two-column contact row."""
        pairs = [
            ("Email", self.d.get("email", "")),
            ("Phone", self.d.get("phone", "")),
            ("Location", self.d.get("location", "")),
            ("LinkedIn", self.d.get("linkedin", "")),
            ("Website", self.d.get("website", "")),
        ]
        items = [(k, v) for k, v in pairs if v.strip()]
        if not items:
            return []
        out = self.sec("Contact Information")
        rows, row = [], []
        for k, v in items:
            row.append(
                Paragraph(
                    f'<font color="#1B3A6B"><b>{k}:</b></font>'
                    f'<font color="#333333">  {v}</font>',
                    self.st["body"],
                )
            )
            if len(row) == 2:
                rows.append(row)
                row = []
        if row:
            row.append(Paragraph("", self.st["body"]))
            rows.append(row)
        t = Table(rows, colWidths=[self.CW / 2] * 2)
        t.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )
        out.append(t)
        return out

    def _summary(self):
        txt = self.d.get("summary", "").strip()
        if not txt:
            return []
        return self.sec("Professional Summary") + [Paragraph(txt, self.st["bodyJ"])]

    def _competencies(self):
        """Each line: bold navy category + normal grey items."""
        raw = self.d.get("competencies", "").strip()
        if not raw:
            return []
        out = self.sec("Core Competencies")
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            if ":" in line:
                cat, _, rest = line.partition(":")
                # matches DOCX: category bold #1B3A6B, items #333333
                out.append(
                    Paragraph(
                        f'<font color="#1B3A6B"><b>{cat.strip()}:</b></font>'
                        f'<font color="#333333">  {rest.strip()}</font>',
                        self.st["compet"],
                    )
                )
            else:
                out.append(Paragraph(line, self.st["body"]))
        return out

    def _experience(self):
        raw = self.d.get("experience", "").strip()
        if not raw:
            return []
        out = self.sec("Work Experience")
        # Split on blank lines into job blocks
        blocks = [b.strip() for b in re.split(r"\n{2,}", raw) if b.strip()]
        for block in blocks:
            lines = [l.strip() for l in block.splitlines() if l.strip()]
            items, got_title, got_company = [], False, False
            for line in lines:
                if re.match(r"^[-•*]", line):
                    items.append(self.bul(line))
                elif not got_title:
                    # Job title: bold #1B3A6B 11pt
                    items.append(Paragraph(line, self.st["jobtitle"]))
                    got_title = True
                elif not got_company:
                    # Company  |  Date — replicate DOCX formatting exactly
                    if "|" in line:
                        co, _, dt = line.partition("|")
                        items.append(
                            Paragraph(
                                f'<font color="#2E6DB4"><b>{co.strip()}</b></font>'
                                f'<font color="#555555">  |  <i>{dt.strip()}</i></font>',
                                self.st["company"],
                            )
                        )
                    else:
                        items.append(
                            Paragraph(
                                f'<font color="#2E6DB4"><b>{line}</b></font>',
                                self.st["company"],
                            )
                        )
                    got_company = True
                else:
                    items.append(Paragraph(line, self.st["body"]))
            items.append(self.sp(4))
            out.extend(items)
        return out

    def _education(self):
        raw = self.d.get("education", "").strip()
        if not raw:
            return []
        out = self.sec("Education")
        for block in [b.strip() for b in re.split(r"\n{2,}", raw) if b.strip()]:
            lines = [l.strip() for l in block.splitlines() if l.strip()]
            if not lines:
                continue
            out.append(Paragraph(lines[0], self.st["jobtitle"]))
            for l in lines[1:]:
                out.append(
                    Paragraph(
                        f'<font color="#2E6DB4"><b>{l.split("|")[0].strip()}</b></font>'
                        + (
                            f'<font color="#555555">  |  <i>{l.split("|")[1].strip()}</i></font>'
                            if "|" in l
                            else ""
                        ),
                        self.st["company"],
                    )
                )
            out.append(self.sp(4))
        return out

    def _training(self):
        raw = self.d.get("training", "").strip()
        if not raw:
            return []
        out = self.sec("Key Training Programmes Delivered")
        for l in raw.splitlines():
            if l.strip():
                out.append(self.bul(l))
        return out

    def _achievements(self):
        raw = self.d.get("achievements", "").strip()
        if not raw:
            return []
        out = self.sec("Achievements & Highlights")
        for l in raw.splitlines():
            if l.strip():
                out.append(self.bul(l))
        return out

    def _skills(self):
        raw = self.d.get("skills", "").strip()
        if not raw:
            return []
        out = self.sec("Skills")
        items = [s.strip() for s in re.split(r"[,\n]", raw) if s.strip()]
        # 3-column grid
        rows = []
        for i in range(0, len(items), 3):
            chunk = items[i : i + 3]
            while len(chunk) < 3:
                chunk.append("")
            rows.append(
                [
                    (
                        Paragraph(
                            f'<font color="#1B3A6B">•</font>'
                            f'<font color="#333333"> {s}</font>',
                            self.st["body"],
                        )
                        if s
                        else Paragraph("", self.st["body"])
                    )
                    for s in chunk
                ]
            )
        t = Table(rows, colWidths=[self.CW / 3] * 3)
        t.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )
        out.append(t)
        return out

    def _certifications(self):
        raw = self.d.get("certifications", "").strip()
        if not raw:
            return []
        out = self.sec("Certifications")
        for l in raw.splitlines():
            if l.strip():
                out.append(self.bul(l))
        return out

    def _languages(self):
        raw = self.d.get("languages", "").strip()
        if not raw:
            return []
        return self.sec("Languages") + [
            Paragraph(raw.replace("\n", "  •  "), self.st["body"])
        ]

    def _notes(self):
        raw = self.d.get("additional_notes", "").strip()
        if not raw:
            return []
        return self.sec("Additional Information") + [Paragraph(raw, self.st["bodyJ"])]

    # ── Generate ─────────────────────────────────────────────────────────────
    def generate(self):
        doc = SimpleDocTemplate(
            self.out,
            pagesize=A4,
            leftMargin=self.L_MAR,
            rightMargin=self.R_MAR,
            topMargin=self.T_MAR,
            bottomMargin=self.B_MAR,
        )
        story = []
        # Header text (sits inside the navy band thanks to topMargin offset)
        story += self._hdr_text()
        story.append(self.sp(8))
        # Body sections
        story += self._contact()
        story += self._summary()
        story += self._competencies()
        story += self._experience()
        story += self._education()
        story += self._training()
        story += self._achievements()
        story += self._skills()
        story += self._certifications()
        story += self._languages()
        story += self._notes()

        doc.build(story, onFirstPage=self._page1, onLaterPages=self._page_n)


# ─────────────────────────────────────────────────────────────────────────────
#  CV FILE PARSER  (DOCX + PDF → data dict)
# ─────────────────────────────────────────────────────────────────────────────
class CVParser:
    # Maps lowercase keyword → field key
    MAP = {
        "professional summary": "summary",
        "summary": "summary",
        "core competencies": "competencies",
        "competencies": "competencies",
        "work experience": "experience",
        "experience": "experience",
        "employment": "experience",
        "education": "education",
        "qualification": "education",
        "key training": "training",
        "training programme": "training",
        "training program": "training",
        "achievements": "achievements",
        "highlights": "achievements",
        "skills": "skills",
        "certification": "certifications",
        "languages": "languages",
    }
    # Identify section header paragraphs in DOCX by style + formatting
    SEC_STYLES = {"Heading 1", "Heading 2", "Heading 3"}

    @classmethod
    def from_docx(cls, path: str) -> dict:
        if not DOCX_OK:
            raise RuntimeError(
                "python-docx not installed. Run: pip install python-docx"
            )
        doc = DocxDoc(path)
        data = {}

        # ── Extract name/title from the header table ──────────────────────────
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    txt = cell.text.strip()
                    if txt and len(txt) > 10:
                        lines = [l.strip() for l in txt.split("\n") if l.strip()]
                        if lines:
                            data["full_name"] = lines[0]
                        if len(lines) > 1:
                            data["job_title"] = lines[1]
                        if len(lines) > 2:
                            data["tagline"] = "  •  ".join(lines[2:])
                        break

        # ── Walk paragraphs ───────────────────────────────────────────────────
        cur_key = None
        buf = []

        def _flush():
            if cur_key and buf:
                existing = data.get(cur_key, "")
                chunk = "\n".join(buf).strip()
                data[cur_key] = (
                    existing + ("\n\n" if existing else "") + chunk
                ).strip()
            buf.clear()

        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                if buf:
                    buf.append("")
                continue

            # Detect section heading: bold + dark-navy colour OR known ALLCAPS text
            is_sec_hdr = False
            lo = txt.lower().strip()
            for kw, key in cls.MAP.items():
                if kw in lo and len(lo) < 70 and lo[: len(kw) + 5].count(" ") <= 5:
                    is_sec_hdr = True
                    _flush()
                    cur_key = key
                    break

            if is_sec_hdr:
                continue

            # Classify paragraph by style
            if p.style.name == "List Paragraph":
                buf.append(f"- {txt}")
            else:
                buf.append(txt)

        _flush()
        return data

    @classmethod
    def from_pdf(cls, path: str) -> dict:
        if not PDFPLUMBER_OK:
            raise RuntimeError("pdfplumber not installed. Run: pip install pdfplumber")
        lines = []
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages:
                raw = pg.extract_text() or ""
                lines += raw.splitlines()
        return cls._parse_lines([l.strip() for l in lines if l.strip()])

    @classmethod
    def _parse_lines(cls, lines: list) -> dict:
        data, cur_key, buf = {}, None, []

        def _flush():
            if cur_key and buf:
                chunk = "\n".join(buf).strip()
                existing = data.get(cur_key, "")
                data[cur_key] = (
                    existing + ("\n\n" if existing else "") + chunk
                ).strip()
            buf.clear()

        for i, line in enumerate(lines):
            lo = line.lower()
            matched = None
            for kw, key in cls.MAP.items():
                if kw in lo and len(lo) < 70:
                    matched = key
                    break
            if matched:
                _flush()
                cur_key = matched
            else:
                if i == 0 and not data.get("full_name"):
                    data["full_name"] = line
                buf.append(line)

        _flush()
        return data


# ─────────────────────────────────────────────────────────────────────────────
#  TOOLTIP HELPER
# ─────────────────────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, widget, text):
        self.w = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        x = self.w.winfo_rootx() + 20
        y = self.w.winfo_rooty() + self.w.winfo_height() + 4
        self.tip = tk.Toplevel(self.w)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip,
            text=self.text,
            background="#1B3A6B",
            foreground="white",
            font=("Segoe UI", 8),
            padx=8,
            pady=4,
            relief="flat",
        ).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN GUI APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class CVMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CV Maker  —  Corporates Guide")
        self.configure(bg=G_BG)
        self.geometry("1160x820")
        self.minsize(960, 680)
        self._logo_path = LOGO
        self._fields: dict[str, tk.Widget] = {}
        self._status_var = tk.StringVar(value="Ready")
        self._build_ui()

    # ══════════════════════════════════════════════════════════════════════════
    #  TOP-LEVEL LAYOUT
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        self._topbar()

        outer = tk.Frame(self, bg=G_BG)
        outer.pack(fill="both", expand=True, padx=12, pady=10)

        # Left (notebook)
        left = tk.Frame(outer, bg=G_BG)
        left.pack(side="left", fill="both", expand=True)

        # Right (sidebar)
        right = tk.Frame(
            outer, bg=G_CARD, highlightthickness=1, highlightbackground=G_BORDER
        )
        right.pack(side="right", fill="y", padx=(10, 0))
        right.config(width=250)
        right.pack_propagate(False)

        self._notebook(left)
        self._sidebar(right)

    # ── Top bar ────────────────────────────────────────────────────────────────
    def _topbar(self):
        bar = tk.Frame(self, bg=G_TOPBAR, height=60)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # Logo
        if os.path.exists(LOGO):
            try:
                img = PILImage.open(LOGO).resize((110, 54), PILImage.LANCZOS)
                self._tl_img = ImageTk.PhotoImage(img)
                tk.Label(bar, image=self._tl_img, bg=G_TOPBAR, cursor="hand2").pack(
                    side="left", padx=14, pady=3
                )
            except Exception:
                pass

        tk.Label(
            bar, text="CV Maker", font=("Segoe UI", 16, "bold"), fg=G_TEXT, bg=G_TOPBAR
        ).pack(side="left", padx=2)
        tk.Label(
            bar,
            text="— Professional CV Generator",
            font=("Segoe UI", 10),
            fg="#BDD7EE",
            bg=G_TOPBAR,
        ).pack(side="left", padx=4)

        # Status chip
        chip = tk.Frame(bar, bg=G_TOPBAR)
        chip.pack(side="right", padx=16)
        tk.Label(
            chip, text="STATUS:", font=("Segoe UI", 8), fg=G_MUTED, bg=G_TOPBAR
        ).pack(side="left")
        self._status_lbl = tk.Label(
            chip,
            textvariable=self._status_var,
            font=("Segoe UI", 8, "bold"),
            fg=G_SUCCESS,
            bg=G_TOPBAR,
        )
        self._status_lbl.pack(side="left", padx=(3, 0))

    # ── Notebook ───────────────────────────────────────────────────────────────
    def _notebook(self, parent):
        st = ttk.Style()
        st.theme_use("default")
        st.configure(
            "CG.TNotebook", background=G_BG, borderwidth=0, tabmargins=[0, 0, 0, 0]
        )
        st.configure(
            "CG.TNotebook.Tab",
            background=G_CARD,
            foreground=G_MUTED,
            padding=[16, 7],
            font=("Segoe UI", 9),
            borderwidth=0,
            focuscolor="",
        )
        st.map(
            "CG.TNotebook.Tab",
            background=[("selected", G_TOPBAR), ("active", "#253b5e")],
            foreground=[("selected", G_TEXT), ("active", G_TEXT)],
        )

        nb = ttk.Notebook(parent, style="CG.TNotebook")
        nb.pack(fill="both", expand=True)

        tabs_spec = [
            ("👤  Personal", self._tab_personal),
            ("📝  Summary", self._tab_summary),
            ("💼  Experience", self._tab_experience),
            ("🎓  Education", self._tab_education),
            ("⭐  Extras", self._tab_extras),
        ]
        for label, builder in tabs_spec:
            frame = tk.Frame(nb, bg=G_CARD, padx=18, pady=14)
            nb.add(frame, text=label)
            builder(frame)

    # ════════════════════════════════════════════════════════════════════════
    #  TABS
    # ════════════════════════════════════════════════════════════════════════

    # ── Personal ──────────────────────────────────────────────────────────────
    def _tab_personal(self, p):
        p.columnconfigure(0, weight=1)
        FIELDS = [
            ("full_name", "Full Name *", "e.g.  SHREYA"),
            (
                "job_title",
                "Professional Title *",
                "e.g.  Senior CRO & Experimentation Trainer",
            ),
            (
                "tagline",
                "Tag Line / Keywords (shown under title in header)",
                "e.g.  8+ Years of Industry Experience   |   A/B Testing  •  VWO Platform  •  CRO Strategy",
            ),
            ("email", "Email", "e.g.  shreya@corporatesguide.com"),
            ("phone", "Phone", "e.g.  +91 98765 43210"),
            ("location", "Location", "e.g.  Mumbai, India"),
            ("linkedin", "LinkedIn", "e.g.  linkedin.com/in/shreya-cro"),
            ("website", "Website", "e.g.  corporatesguide.com"),
        ]
        for row_i, (key, label, hint) in enumerate(FIELDS):
            self._lbl(p, label).grid(
                row=row_i * 2, column=0, sticky="w", pady=(10 if row_i == 0 else 7, 0)
            )
            e = self._entry(p)
            e.grid(row=row_i * 2 + 1, column=0, sticky="ew", pady=(2, 0), ipady=5)
            self._placeholder(e, hint)
            self._fields[key] = e

    # ── Summary & Competencies ─────────────────────────────────────────────────
    def _tab_summary(self, p):
        self._lbl(p, "Professional Summary").pack(anchor="w")
        self._hint(p, "Paste your professional summary paragraph.").pack(anchor="w")
        self._fields["summary"] = self._textarea(p, h=7)
        self._fields["summary"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Core Competencies").pack(anchor="w")
        self._hint(
            p,
            "One category per line, format:  Category Name: item1   •   item2   •   item3",
        ).pack(anchor="w")
        self._fields["competencies"] = self._textarea(p, h=13)
        self._fields["competencies"].pack(fill="x", pady=(3, 0))

    # ── Work Experience ────────────────────────────────────────────────────────
    def _tab_experience(self, p):
        self._hint(
            p,
            "Each job block separated by a blank line.\n"
            "  Line 1 : Job Title\n"
            "  Line 2 : Company Name  |  Start Date – End Date\n"
            "  Lines 3+: - bullet point  (start with  -  or  •)",
        ).pack(anchor="w", pady=(0, 8))
        self._fields["experience"] = self._textarea(p, h=27)
        self._fields["experience"].pack(fill="both", expand=True)

    # ── Education & Skills ─────────────────────────────────────────────────────
    def _tab_education(self, p):
        self._lbl(p, "Education").pack(anchor="w")
        self._hint(
            p,
            "Blank line between entries.\n"
            "  Line 1: Degree / Qualification\n"
            "  Line 2: Institution  |  Year Range",
        ).pack(anchor="w")
        self._fields["education"] = self._textarea(p, h=7)
        self._fields["education"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Skills  (comma-separated)").pack(anchor="w")
        self._fields["skills"] = self._textarea(p, h=4)
        self._fields["skills"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Certifications  (one per line)").pack(anchor="w")
        self._fields["certifications"] = self._textarea(p, h=5)
        self._fields["certifications"].pack(fill="x", pady=(3, 0))

    # ── Extras ────────────────────────────────────────────────────────────────
    def _tab_extras(self, p):
        self._lbl(p, "Key Training Programmes Delivered  (one per line)").pack(
            anchor="w"
        )
        self._fields["training"] = self._textarea(p, h=7)
        self._fields["training"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Achievements & Highlights  (one per line)").pack(anchor="w")
        self._fields["achievements"] = self._textarea(p, h=7)
        self._fields["achievements"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Languages").pack(anchor="w")
        self._fields["languages"] = self._textarea(p, h=2)
        self._fields["languages"].pack(fill="x", pady=(3, 12))

        self._lbl(p, "Additional Notes").pack(anchor="w")
        self._fields["additional_notes"] = self._textarea(p, h=3)
        self._fields["additional_notes"].pack(fill="x", pady=(3, 0))

    # ════════════════════════════════════════════════════════════════════════
    #  SIDEBAR
    # ════════════════════════════════════════════════════════════════════════
    def _sidebar(self, p):
        def section(title):
            tk.Frame(p, bg=G_BORDER, height=1).pack(fill="x", pady=(8, 0))
            tk.Label(
                p, text=title, font=("Segoe UI", 8, "bold"), fg=G_MUTED, bg=G_CARD
            ).pack(anchor="w", padx=14, pady=(6, 3))

        # Title
        tk.Label(
            p, text="Actions", font=("Segoe UI", 13, "bold"), fg=G_TEXT, bg=G_CARD
        ).pack(pady=(18, 2))
        tk.Label(
            p,
            text="Corporates Guide CV Maker",
            font=("Segoe UI", 8),
            fg=G_MUTED,
            bg=G_CARD,
        ).pack()

        # ── Import ──────────────────────────────────────────────────────────
        section("IMPORT EXISTING CV")
        b_up = self._btn(p, "📂  Upload DOCX / PDF", self._on_upload, G_ACCENT)
        b_up.pack(fill="x", padx=14, pady=4)
        Tooltip(b_up, "Import a .docx or .pdf CV to auto-fill all fields")

        # ── Logo ─────────────────────────────────────────────────────────────

        # ── Generate PDF ──────────────────────────────────────────────────────
        section("GENERATE")
        gen_btn = self._btn(
            p,
            "⬇️  Generate & Save PDF",
            self._on_generate,
            bg="#1B3A6B",
            fg="white",
            bold=True,
        )
        gen_btn.pack(fill="x", padx=14, pady=6)
        Tooltip(gen_btn, "Generate a professional PDF CV")

        # ── Utils ─────────────────────────────────────────────────────────────
        section("UTILITIES")
        self._btn(p, "🔄  Clear All Fields", self._on_clear, "#2a1a1a").pack(
            fill="x", padx=14, pady=4
        )
        self._btn(p, "💾  Save Draft (JSON)", self._on_save_draft, "#1f2d40").pack(
            fill="x", padx=14, pady=2
        )
        self._btn(p, "📂  Load Draft (JSON)", self._on_load_draft, "#1f2d40").pack(
            fill="x", padx=14, pady=2
        )

        # Status message
        self._msg_lbl = tk.Label(
            p,
            text="",
            fg=G_SUCCESS,
            bg=G_CARD,
            font=("Segoe UI", 8),
            wraplength=218,
            justify="center",
        )
        self._msg_lbl.pack(pady=14, padx=8)

        # Footer
        tk.Frame(p, bg=G_BORDER, height=1).pack(fill="x", side="bottom", pady=(0, 8))
        tk.Label(
            p,
            text="Corporates Guide  •  Your Guide to Success",
            font=("Segoe UI", 7),
            fg=G_MUTED,
            bg=G_CARD,
        ).pack(side="bottom", pady=(0, 6))

    # ════════════════════════════════════════════════════════════════════════
    #  WIDGET FACTORIES
    # ════════════════════════════════════════════════════════════════════════
    def _lbl(self, p, text):
        return tk.Label(
            p, text=text, font=("Segoe UI", 9, "bold"), fg=G_TEXT, bg=G_CARD
        )

    def _hint(self, p, text):
        return tk.Label(
            p, text=text, font=("Segoe UI", 8), fg=G_MUTED, bg=G_CARD, justify="left"
        )

    def _entry(self, parent):
        return tk.Entry(
            parent,
            bg=G_INPUT,
            fg=G_MUTED,
            insertbackground=G_TEXT,
            relief="flat",
            font=("Segoe UI", 9),
            highlightthickness=1,
            highlightbackground=G_BORDER,
            highlightcolor=G_ACCENT,
        )

    def _textarea(self, parent, h=5) -> tk.Frame:
        wrap = tk.Frame(parent, bg=G_BORDER, bd=1)
        t = tk.Text(
            wrap,
            bg=G_INPUT,
            fg=G_TEXT,
            insertbackground=G_TEXT,
            relief="flat",
            font=("Segoe UI", 9),
            height=h,
            wrap="word",
            highlightthickness=0,
            padx=7,
            pady=5,
        )
        sc = tk.Scrollbar(
            wrap, command=t.yview, bg=G_CARD, troughcolor=G_BG, bd=0, width=10
        )
        t.config(yscrollcommand=sc.set)
        t.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")
        wrap.text_widget = t  # store reference
        return wrap

    def _btn(self, parent, text, cmd, bg=G_ACCENT, fg=G_TEXT, bold=False) -> tk.Button:
        font = ("Segoe UI", 9, "bold") if bold else ("Segoe UI", 9)
        b = tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg=fg,
            activebackground="#2E6DB4",
            activeforeground="white",
            relief="flat",
            font=font,
            cursor="hand2",
            pady=8,
        )
        b.bind("<Enter>", lambda e: b.config(bg="#2E6DB4"))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _placeholder(self, entry: tk.Entry, hint: str):
        """Placeholder text that disappears on focus."""
        entry.insert(0, hint)
        entry.config(fg=G_MUTED)

        def _in(ev):
            if entry.cget("fg") == G_MUTED:
                entry.delete(0, "end")
                entry.config(fg=G_TEXT)

        def _out(ev):
            if not entry.get().strip():
                entry.insert(0, hint)
                entry.config(fg=G_MUTED)

        entry.bind("<FocusIn>", _in)
        entry.bind("<FocusOut>", _out)

    # ════════════════════════════════════════════════════════════════════════
    #  DATA HELPERS
    # ════════════════════════════════════════════════════════════════════════
    def _get(self, key: str) -> str:
        w = self._fields.get(key)
        if w is None:
            return ""
        if isinstance(w, tk.Entry):
            v = w.get().strip()
            return "" if w.cget("fg") == G_MUTED else v
        if isinstance(w, tk.Frame) and hasattr(w, "text_widget"):
            return w.text_widget.get("1.0", "end").strip()
        return ""

    def _set(self, key: str, value: str):
        if not value:
            return
        w = self._fields.get(key)
        if w is None:
            return
        if isinstance(w, tk.Entry):
            w.config(fg=G_TEXT)
            w.delete(0, "end")
            w.insert(0, value)
        elif isinstance(w, tk.Frame) and hasattr(w, "text_widget"):
            w.text_widget.delete("1.0", "end")
            w.text_widget.insert("1.0", value)

    def _collect(self) -> dict:
        return {
            k: self._get(k)
            for k in [
                "full_name",
                "job_title",
                "tagline",
                "email",
                "phone",
                "location",
                "linkedin",
                "website",
                "summary",
                "competencies",
                "experience",
                "education",
                "skills",
                "certifications",
                "training",
                "achievements",
                "languages",
                "additional_notes",
            ]
        }

    def _populate(self, data: dict):
        for k, v in data.items():
            self._set(k, v)

    def _msg(self, text: str, color=G_SUCCESS):
        self._msg_lbl.config(text=text, fg=color)
        self._status_var.set(text[:40])
        self._status_lbl.config(fg=color)
        self.update_idletasks()

    # ════════════════════════════════════════════════════════════════════════
    #  ACTIONS
    # ════════════════════════════════════════════════════════════════════════
    def _on_upload(self):
        path = filedialog.askopenfilename(
            title="Open Existing CV",
            filetypes=[
                ("CV files", "*.docx *.pdf"),
                ("Word Document", "*.docx"),
                ("PDF", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self._msg("Parsing file…", G_WARN)

        def _run():
            try:
                ext = Path(path).suffix.lower()
                if ext == ".docx":
                    data = CVParser.from_docx(path)
                elif ext == ".pdf":
                    data = CVParser.from_pdf(path)
                else:
                    raise ValueError("Unsupported file type. Use .docx or .pdf")
                if data:
                    self.after(0, lambda: self._populate(data))
                    self.after(
                        0,
                        lambda: self._msg(f"✓  Imported  {Path(path).name}", G_SUCCESS),
                    )
                else:
                    self.after(
                        0,
                        lambda: self._msg(
                            "⚠  Could not detect sections.\nPlease fill fields manually.",
                            G_WARN,
                        ),
                    )
            except Exception as ex:
                self.after(
                    0,
                    lambda: (
                        self._msg("Error parsing file.", G_RED),
                        messagebox.showerror("Parse Error", str(ex)),
                    ),
                )

        threading.Thread(target=_run, daemon=True).start()

    def _on_generate(self):
        data = self._collect()
        if not data.get("full_name"):
            messagebox.showwarning(
                "Missing Field", "Please enter at least a Full Name before generating."
            )
            return

        default = data["full_name"].replace(" ", "_") + "_CV.pdf"
        out = filedialog.asksaveasfilename(
            title="Save CV as PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf"), ("All files", "*.*")],
            initialfile=default,
        )
        if not out:
            return

        self._msg("⏳  Generating PDF…", G_WARN)

        def _run():
            try:
                CVGenerator(data, out, self._logo_path).generate()
                self.after(
                    0,
                    lambda: (
                        self._msg(f"✓  PDF saved!\n{Path(out).name}", G_SUCCESS),
                        messagebox.showinfo(
                            "Success", f"CV saved successfully!\n\nLocation:\n{out}"
                        ),
                    ),
                )
            except Exception as ex:
                self.after(
                    0,
                    lambda: (
                        self._msg("Error generating PDF.", G_RED),
                        messagebox.showerror("PDF Error", str(ex)),
                    ),
                )

        threading.Thread(target=_run, daemon=True).start()

    def _on_clear(self):
        if not messagebox.askyesno(
            "Clear All", "This will erase all entered data. Continue?"
        ):
            return
        for key, w in self._fields.items():
            if isinstance(w, tk.Entry):
                w.delete(0, "end")
                w.config(fg=G_MUTED)
            elif isinstance(w, tk.Frame) and hasattr(w, "text_widget"):
                w.text_widget.delete("1.0", "end")
        self._msg("All fields cleared.", G_MUTED)

    def _on_save_draft(self):
        import json

        data = self._collect()
        out = filedialog.asksaveasfilename(
            title="Save Draft",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
            initialfile="cv_draft.json",
        )
        if not out:
            return
        with open(out, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        self._msg(f"✓  Draft saved:\n{Path(out).name}")

    def _on_load_draft(self):
        import json

        path = filedialog.askopenfilename(
            title="Load Draft", filetypes=[("JSON", "*.json"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._populate(data)
            self._msg(f"✓  Draft loaded:\n{Path(path).name}")
        except Exception as ex:
            messagebox.showerror("Load Error", str(ex))

    def _on_exe_guide(self):
        script = os.path.abspath(__file__)
        logo = self._logo_path

        # Platform-appropriate separator
        sep = ";" if sys.platform == "win32" else ":"

        cmd = (
            f"pyinstaller --onefile --windowed "
            f'--add-data "{logo}{sep}." '
            f'--name "CVMaker_CorporatesGuide" '
            f'--clean "{script}"'
        )

        win = tk.Toplevel(self)
        win.title("Build Standalone EXE")
        win.configure(bg=G_CARD)
        win.geometry("680x420")

        tk.Label(
            win,
            text="📦  Build a Standalone Desktop App",
            font=("Segoe UI", 13, "bold"),
            fg=G_TEXT,
            bg=G_CARD,
        ).pack(pady=(18, 4))
        tk.Label(
            win,
            text="Run the command below in your terminal. The EXE will be in the dist/ folder.",
            font=("Segoe UI", 9),
            fg=G_MUTED,
            bg=G_CARD,
        ).pack(pady=(0, 10))

        instructions = (
            "STEP 1 — Install PyInstaller\n"
            "    pip install pyinstaller\n\n"
            "STEP 2 — Run this command from the same folder as cv_maker.py\n\n"
            + cmd
            + "\n\n"
            "STEP 3 — Collect your EXE\n"
            "    dist/CVMaker_CorporatesGuide.exe  (Windows)\n"
            "    dist/CVMaker_CorporatesGuide      (macOS / Linux)\n\n"
            "NOTE: The logo.jpg must be in the same folder when running the command."
        )
        t = tk.Text(
            win,
            bg=G_INPUT,
            fg=G_TEXT,
            font=("Courier New", 9),
            relief="flat",
            padx=12,
            pady=10,
            wrap="word",
            height=14,
        )
        t.insert("1.0", instructions)
        t.config(state="disabled")
        t.pack(fill="both", expand=True, padx=16)

        btn_frame = tk.Frame(win, bg=G_CARD)
        btn_frame.pack(pady=10)

        def _copy():
            self.clipboard_clear()
            self.clipboard_append(cmd)
            messagebox.showinfo("Copied", "PyInstaller command copied to clipboard!")

        self._btn(btn_frame, "📋  Copy Command", _copy, G_TOPBAR, "white").pack(
            side="left", padx=6
        )
        self._btn(btn_frame, "✕  Close", win.destroy, "#2a2a2a").pack(
            side="left", padx=6
        )


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = CVMakerApp()
    app.mainloop()
