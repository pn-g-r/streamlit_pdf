"""Georgian medical document -> English PDF + Optional Excel.

Design Philosophy:
  - Readability and executive medical typography come first (clean fonts, sharp contrast,
    comfortable line-height, elegant table borders).
  - Faithful layout: mimics side-by-side header cards, patient metadata grids,
    and diagnostic tables without awkward cramming or micro-fonts.
  - Multi-row lab panels break naturally across pages with repeating headers (<thead>).
  - Zero Georgian characters in English translations (two-phase clinical translation).
  - Optional Excel extraction: skips the 30-45s structured schema inference by default,
    making PDF generation blazing fast (~5-10s).

Run as Streamlit app:   streamlit run pdf_html_pdf.py
Run as batch CLI:       python pdf_html_pdf.py --cli file1.pdf [file2.pdf ...] [--out DIR] [--excel]
"""
import base64
import concurrent.futures as cf
import io
import json
import os
import re
import statistics
import sys
import time

from dotenv import load_dotenv

load_dotenv()

# Windows-only MSYS2 DLL registration (ignored on Linux/Streamlit Cloud)
if sys.platform == "win32":
    current_path = os.environ.get("PATH", "")
    cleaned_path = [p for p in current_path.split(os.pathsep) if "tesseract" not in p.lower()]
    candidate_paths = [
        r"C:\msys64\ucrt64\bin",
        r"C:\msys64\mingw64\bin",
        r"C:\Program Files\GTK3-Runtime Win64\bin",
    ]
    valid_dll_dir = None
    for path in candidate_paths:
        if os.path.isdir(path) and os.path.exists(os.path.join(path, "libpango-1.0-0.dll")):
            valid_dll_dir = path
            break
    if valid_dll_dir:
        os.add_dll_directory(valid_dll_dir)
        cleaned_path.insert(0, valid_dll_dir)
    os.environ["PATH"] = os.pathsep.join(cleaned_path)

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from pydantic import BaseModel, Field
import pymupdf
from google import genai
from google.genai import types
from weasyprint import HTML

MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
PT_TO_MM = 25.4 / 72


def _probe_find_tables():
    """Fallback if find_tables encounters DLL issues on Windows."""
    try:
        import importlib.util
        if importlib.util.find_spec("onnxruntime") is None:
            return False
        import onnxruntime  # noqa: F401
        return False
    except Exception:
        return True


_FIND_TABLES_BROKEN = _probe_find_tables()


# ---------------------------------------------------------------------------
# Schemas (Excel extraction)
# ---------------------------------------------------------------------------
class UniversalField(BaseModel):
    section_or_item_number: str = Field(description="Point #, section name, or category")
    field_label_ge: str = Field(description="Original Georgian label verbatim")
    field_label_en: str = Field(description="English translation of label")
    field_value_ge: str = Field(default="", description="Filled Georgian text verbatim")
    field_value_en: str = Field(default="", description="English translation or exact numbers")


class UniversalTableRow(BaseModel):
    table_name: str = Field(description="Table name")
    row_number: int = Field(description="Sequential row index")
    group_or_date: str = Field(default="", description="Date or subgroup header")
    code: str = Field(default="", description="Code or test abbreviation")
    name_ge: str = Field(default="", description="Name in Georgian")
    name_en: str = Field(default="", description="Name in English")
    result_or_qty: str = Field(default="", description="Result or quantity")
    unit_or_reference_range: str = Field(default="", description="Unit or reference range")
    notes_or_status: str = Field(default="", description="Status flag or notes")


class UniversalDocumentData(BaseModel):
    document_title: str = Field(description="Document title in English")
    fields: list[UniversalField] = Field(description="All metadata fields and numbered points")
    tables: list[UniversalTableRow] = Field(description="All table rows")


# ---------------------------------------------------------------------------
# 1. Layout & Vector Analysis (PyMuPDF)
# ---------------------------------------------------------------------------
def _image_to_png(doc, page, xref, bbox):
    try:
        if doc.xref_get_key(xref, "ImageMask")[1] == "true":
            raise ValueError("stencil")
        if doc.xref_get_key(xref, "Decode")[0] != "null":
            raise ValueError("decode array")
        pix = pymupdf.Pixmap(doc, xref)
        smask = doc.xref_get_key(xref, "SMask")
        if pix.colorspace and pix.colorspace.n not in (1, 3):
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        if smask[0] == "xref":
            mask = pymupdf.Pixmap(doc, int(smask[1].split()[0]))
            if pix.alpha:
                pix = pymupdf.Pixmap(pix, 0)
            pix = pymupdf.Pixmap(pix, mask)
        return pix.tobytes("png")
    except Exception:
        clip = pymupdf.Rect(bbox) & page.rect
        pix = page.get_pixmap(clip=clip, dpi=300, alpha=True, colorspace=pymupdf.csRGB)
        return pix.tobytes("png")


def _grid_from_drawings(page):
    W = page.rect.width
    v, h = [], []
    try:
        drawings = page.get_drawings()
    except Exception:
        return []
    for d in drawings:
        for it in d["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.x - p2.x) < 1.5 and abs(p1.y - p2.y) > 5:
                    v.append((p1.x, min(p1.y, p2.y), max(p1.y, p2.y)))
                elif abs(p1.y - p2.y) < 1.5 and abs(p1.x - p2.x) > 5:
                    h.append((p1.y, min(p1.x, p2.x), max(p1.x, p2.x)))
            elif it[0] == "re":
                r = it[1]
                if r.width < 2 and r.height > 5:
                    v.append((r.x0, r.y0, r.y1))
                elif r.height < 2 and r.width > 5:
                    h.append((r.y0, r.x0, r.x1))
                elif r.width > 5 and r.height > 5 and d.get("fill") is None:
                    v += [(r.x0, r.y0, r.y1), (r.x1, r.y0, r.y1)]
                    h += [(r.y0, r.x0, r.x1), (r.y1, r.x0, r.x1)]
    if len(h) < 2 or len(v) < 2:
        return []

    # 1. Merge collinear horizontal segments on the same Y
    from collections import defaultdict
    by_y = defaultdict(list)
    for y, x0, x1 in h:
        matched_y = None
        for ky in by_y:
            if abs(y - ky) <= 2:
                matched_y = ky
                break
        if matched_y is None:
            matched_y = y
        by_y[matched_y].append((x0, x1))

    merged_h = []
    for y in sorted(by_y.keys()):
        spans = sorted(by_y[y])
        cur_x0, cur_x1 = spans[0]
        for x0, x1 in spans[1:]:
            if x0 <= cur_x1 + 3:
                cur_x1 = max(cur_x1, x1)
            else:
                merged_h.append((y, cur_x0, cur_x1))
                cur_x0, cur_x1 = x0, x1
        merged_h.append((y, cur_x0, cur_x1))

    if not merged_h:
        return []

    # 2. Group into distinct tables: requires y_gap <= 35 AND matching x0/x1 bounds (within 35pt)
    groups = []
    cur = [merged_h[0]]
    for line in merged_h[1:]:
        prev = cur[-1]
        y_gap = line[0] - prev[0]
        x0_diff = abs(line[1] - cur[0][1])
        x1_diff = abs(line[2] - cur[0][2])
        if y_gap <= 35 and x0_diff <= 35 and x1_diff <= 35:
            cur.append(line)
        else:
            if len(cur) >= 3:
                groups.append(cur)
            cur = [line]
    if len(cur) >= 3:
        groups.append(cur)

    tables = []
    for g in groups:
        y0, y1 = g[0][0], g[-1][0]
        rows = len({round(l[0]) for l in g}) - 1
        if rows < 1 or y1 - y0 < 12:
            continue
        x0_bound = min(l[1] for l in g) - 3
        x1_bound = max(l[2] for l in g) + 3
        xs = sorted({round(x) for x, a, b in v if a <= y1 + 3 and b >= y0 - 3 and x0_bound <= x <= x1_bound})
        merged = []
        for x in xs:
            if not merged or x - merged[-1] > 3:
                merged.append(x)
        if len(merged) < 2:
            continue
        tw = merged[-1] - merged[0]
        ys = sorted({round(l[0]) for l in g})
        tables.append({
            "bbox": [merged[0], round(y0), merged[-1], round(y1)],
            "rows": rows, "cols": len(merged) - 1,
            "col_width_pct": [round(100 * (b - a) / tw) for a, b in zip(merged, merged[1:])],
            "row_heights_pt": [b - a for a, b in zip(ys, ys[1:])],
            "width_pct_of_page": round(100 * tw / W),
        })
    return tables


def _filled_areas(page):
    out = []
    W, H = page.rect.width, page.rect.height
    try:
        drawings = page.get_drawings()
    except Exception:
        return out
    for d in drawings:
        fill = d.get("fill")
        r = d.get("rect")
        if not fill or r is None or r.width < 20 or r.height < 6:
            continue
        if min(fill) > 0.97 or (r.width > 0.95 * W and r.height > 0.95 * H):
            continue
        if fill[0] < 0.03 and fill[1] < 0.03 and fill[2] < 0.03 and r.height < 3:
            continue
        hexc = "#%02x%02x%02x" % tuple(int(255 * c) for c in fill[:3])
        out.append((round(r.y0), round(r.x0), round(r.x1), round(r.y1), hexc))
    out.sort()
    merged = []
    for a in out:
        if merged and merged[-1][4] == a[4] and abs(merged[-1][0] - a[0]) < 2 and abs(merged[-1][3] - a[3]) < 2:
            m = merged[-1]
            merged[-1] = (m[0], min(m[1], a[1]), max(m[2], a[2]), m[3], m[4])
        else:
            merged.append(a)
    return merged[:40]


def analyze_pdf(pdf_bytes):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    pages, all_sizes = [], []
    page_w, page_h = doc[0].rect.width, doc[0].rect.height

    for pno, page in enumerate(doc):
        W, H = page.rect.width, page.rect.height
        lines = []
        raw_text_parts = []
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                spans = [s for s in line["spans"] if s["text"].strip()]
                if not spans:
                    continue
                text = " ".join(s["text"].strip() for s in spans)
                raw_text_parts.append(text)
                size = max(s["size"] for s in spans)
                bold = any(("bold" in s["font"].lower()) or (s["flags"] & 16) for s in spans)
                color = spans[0]["color"]
                x0, y0, x1, y1 = line["bbox"]
                if y1 < 0 or y0 > H:
                    continue
                lines.append((round(y0), round(x0), round(x1), round(size, 1), bold, color, text))
                for s in spans:
                    all_sizes.extend([round(s["size"], 1)] * max(1, len(s["text"].strip())))
        lines.sort()

        tables = []
        try:
            if _FIND_TABLES_BROKEN:
                raise RuntimeError("find_tables unavailable")
            for tb in page.find_tables().tables:
                x0, y0, x1, y1 = tb.bbox
                cols = None
                if tb.rows:
                    best = max(tb.rows, key=lambda r: sum(c is not None for c in r.cells))
                    cells = [c for c in best.cells if c is not None]
                    tw = (x1 - x0) or 1
                    cols = [round(100 * (c[2] - c[0]) / tw) for c in cells]
                tables.append({
                    "bbox": [round(v) for v in tb.bbox],
                    "rows": tb.row_count, "cols": tb.col_count,
                    "col_width_pct": cols,
                    "width_pct_of_page": round(100 * (x1 - x0) / W),
                })
        except Exception:
            pass
        if not tables:
            tables = _grid_from_drawings(page)

        images, seen = [], set()
        for info in page.get_image_info(xrefs=True):
            bbox = pymupdf.Rect(info["bbox"])
            vis = bbox & page.rect
            if vis.is_empty or vis.width < 12 or vis.height < 12:
                continue
            key = (info["xref"], tuple(round(v) for v in bbox))
            if key in seen:
                continue
            seen.add(key)
            if vis.width > 0.9 * W and vis.height > 0.9 * H:
                continue
            img_id = f"p{pno + 1}_{len(images)}"
            if vis.y1 < H * 0.22:
                role = "header logo"
            elif vis.y0 > H * 0.5:
                role = "stamp/signature graphic"
            else:
                role = "inline graphic"
            images.append({
                "id": img_id,
                "x": round(vis.x0), "y": round(vis.y0),
                "w": round(vis.width), "h": round(vis.height),
                "role": role,
                "png": _image_to_png(doc, page, info["xref"], bbox) if info["xref"] else
                       page.get_pixmap(clip=vis, dpi=300, alpha=True).tobytes("png"),
            })

        preview = page.get_pixmap(dpi=150, colorspace=pymupdf.csRGB).tobytes("png")
        single = pymupdf.open()
        single.insert_pdf(doc, from_page=pno, to_page=pno)
        page_pdf = single.tobytes()
        single.close()

        raw_page_text = page.get_text() or "\n".join(raw_text_parts)

        pages.append({
            "number": pno + 1, "width": round(W), "height": round(H),
            "lines": lines, "tables": tables, "images": images,
            "fills": _filled_areas(page),
            "preview_png": preview, "pdf": page_pdf,
            "raw_text": raw_page_text,
            "has_text": bool(lines or raw_page_text.strip()),
        })
    doc.close()

    # Adaptively choose font size and margins:
    # Dense single-page invoices stay comfortably on 1 page (8.5pt, compact margins).
    # Multi-page documents and Form #100 get generous 9.2pt executive styling.
    is_dense_single = (len(pages) == 1 and (len(all_sizes) > 300 or len(pages[0]["lines"]) > 35 or any(t.get("rows", 0) > 8 for t in pages[0]["tables"])))
    base_size = 8.5 if is_dense_single else 9.2
    margins = (10.0, 12.0, 10.0, 12.0) if is_dense_single else (13.0, 14.0, 13.0, 14.0)
    profile = {
        "page_w_mm": page_w * PT_TO_MM,
        "page_h_mm": page_h * PT_TO_MM,
        "base_pt": base_size,
        "margins_mm": margins,
    }
    return pages, profile


def layout_map_text(page):
    out = [f"PAGE {page['number']} — size {page['width']}x{page['height']} pt."]
    if page["tables"]:
        out.append("DETECTED TABLE GRIDS:")
        for i, t in enumerate(page["tables"]):
            rh = t.get("row_heights_pt")
            out.append(f"  table{i + 1}: rows={t['rows']} cols={t['cols']} "
                       f"col_width_pct={t['col_width_pct']} width={t['width_pct_of_page']}%"
                       + (f" row_heights_pt={rh}" if rh else ""))
    if page.get("fills"):
        out.append("FILLED BACKGROUND COLORS [y0, x0-x1, y1, hex]:")
        for y0, x0, x1, y1, c in page["fills"]:
            out.append(f"  [{y0}, {x0}-{x1}, {y1}, {c}]")
    if page["images"]:
        out.append("GRAPHIC ASSETS (insert placeholder where the image belongs):")
        for im in page["images"]:
            out.append(f"  <!-- IMG:{im['id']} --> at x={im['x']} y={im['y']} ({im['role']})")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 2. Executive Medical Translation & Style-First Layout Generation
# ---------------------------------------------------------------------------
SYSTEM_PROMPT_TRANSLATION = """You are a senior medical translator specializing in clinical documentation.
Translate the provided Georgian medical document page text completely and faithfully into professional English.

MANDATORY RULES:
1. COMPLETE TRANSLATION: Translate every single diagnosis, clinical anamnesis paragraph, symptom,
   doctor observation, laboratory test name, measurement unit, reference interval, and administrative note.
   Zero omissions.
2. TRANSLITERATE PROPER NAMES:
   - Patient names, doctor names, and clinic names into standard Latin characters (e.g. ტატიანა ალიევა -> Tatiana Alieva).
3. PRESERVE ALL CLINICAL DATA:
   - Exact numbers, dates, test units, and ICD codes must be preserved verbatim.
4. Return ONLY the complete English clinical translation."""

SYSTEM_PROMPT_HTML_BUILDER = """You are a master medical document layout architect and healthcare document designer.
You receive:
1. High-resolution visual image of the original Georgian medical document page.
2. The complete, verified English clinical translation of all text on this page.
3. Detected visual elements (header logo, stamp/signature anchors, table grids).

Your mission is to perform visual reconstruction: look at the original document image and recreate its visual layout, structure, alignment, and formatting in clean, beautiful HTML using the verified English translation.
Focus on executive styling, balance, and readability above raw copy-pasting.

MANDATORY VISUAL LAYOUT & FORMATTING RULES:
1. HEADER & ISSUER BLOCK (Form #100 & Clinical Reports):
   - Carefully inspect the top area of the original image:
     * Clinic stamp/issuer card (top-left):
       `<div class="clinic-card" style="border: 1pt solid #94a3b8; background-color: #f8fafc; border-radius: 4pt; padding: 6pt 10pt; min-width: 220pt; max-width: 280pt; display: inline-block; box-sizing: border-box;">
          <div style="font-weight: 600; color: #0f172a; font-size: 9.5pt; line-height: 1.35;">Clinic / Hospital Name</div>
          <div style="font-size: 9pt; color: #334155;">Identification Code / Tax ID</div>
          <div style="font-size: 9pt; color: #334155;">Certificate # Number</div>
        </div>`
       CRITICAL: Keep the clinic name on 1-2 natural lines. Do NOT shrink this box or wrap words individually.
     * Legal decree / approval note (top-right):
       `<div class="decree-block" style="max-width: 190pt; margin-left: auto; text-align: right; font-size: 8.5pt; color: #475569; line-height: 1.35;">
          Approved by Order No. 230/N<br>of the Minister of Labor, Health<br>and Social Affairs of Georgia,<br>dated 15.10.2008
        </div>`
       CRITICAL: Constrain its width (`max-width: 190pt;`) and let it wrap into a compact 3-4 line block. Never let it stretch horizontally across the page.
     * Place top-left clinic card and top-right approval side-by-side using:
       `<table class="layout" style="width: 100%; margin-bottom: 12pt; table-layout: fixed;">
          <tr>
            <td style="width: 58%; vertical-align: top;">...clinic card...</td>
            <td style="width: 42%; vertical-align: top; text-align: right;">...decree block...</td>
          </tr>
        </table>`
     * If there is a hospital logo, insert `<!-- IMG:id -->` exactly where it appears.
2. DOCUMENT TITLES (CENTERED & PROMINENT):
   - Medical document titles (such as 'Medical Documentation Form No. IV-100/a', 'CERTIFICATE', 'REGARDING HEALTH STATUS', or 'MEDICAL INVOICE') MUST be prominently centered:
     `<div style="text-align: center; margin: 10pt 0 14pt 0;">
        <div style="font-size: 11pt; font-weight: bold; color: #334155;">Medical Documentation Form No. IV-100/a</div>
        <div style="font-size: 18pt; font-weight: bold; letter-spacing: 1px; margin: 4pt 0; text-transform: uppercase;">CERTIFICATE</div>
        <div style="font-size: 14pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase;">REGARDING HEALTH STATUS</div>
      </div>`
3. NUMBERED QUESTIONNAIRE & FORM ITEMS (Form 100 style):
   - COMFORTABLE VERTICAL RHYTHM & BREATHING ROOM: Form #100 items are official questionnaire fields. Give them generous, uniform vertical spacing so they breathe elegantly and fill the page naturally:
     * Wrap every item in `<div class="form-item" style="margin-bottom: 12pt; line-height: 1.5;">...</div>`.
     * For items where value is below the label:
       `<div class="form-item" style="margin-bottom: 12pt; line-height: 1.5;">
          <div style="color: #334155; font-size: 9.5pt; margin-bottom: 3pt;">1. Name of the issuing institution...:</div>
          <div style="font-weight: 600; color: #0f172a; font-size: 10pt; margin-left: 8pt;">LLC Batumi Republican Clinical Hospital</div>
        </div>`
     * For concise items: `<div class="form-item" style="margin-bottom: 12pt; line-height: 1.5;"><span style="color: #334155;">3. Patient's Name and Surname:</span> <b style="color: #0f172a; margin-left: 6pt;">Anastasia Bolgova</b></div>`
     * Sub-items (like 8. Dates: a, b, c, d): indent with `<div style="margin-left: 14pt; line-height: 1.45; margin-bottom: 4pt;">...</div>`.
     * Recommendations (Item 17): wrap recommendation text in `<div class="box" style="margin-left: 8pt; border: 1px solid #cbd5e1; padding: 6pt 10pt; background: #f8fafc; border-radius: 4pt;">...</div>`.
4. CLINICAL DATA TABLES & FINANCIAL SUMMARY GRIDS:
   - Invoices, CBC laboratory panels, and procedural tests MUST use `<table class="data-table">`.
   - Headers (`<th>`): bold text with professional background.
   - Rows: alternating clean zebra rows with proper borders.
   - Category / Date dividers: `<tr class="group-row"><td colspan="N">...</td></tr>`.
   - FINANCIAL SUMMARY GRIDS (Total Cost, Discounts, Net Total, Insurance Share, Patient Copay / Balance Due):
     * MUST be rendered as a bordered table: `<table class="summary-table">` (or `<table class="data-table">`).
     * The monetary values MUST NOT wrap (e.g., `<td class="nowrap" style="text-align: right; font-weight: bold;">4,850.00 GEL</td>`).
     * Style the final balance due row with `<tr class="total-row">`.
5. SIGNATURES & STAMPS (BOTTOM OF DOCUMENT):
   - Signature blocks (e.g., Billing Manager, Chief Accountant, Doctor, Stamp area) MUST be placed side-by-side using `<table class="layout" style="width: 100%; margin-top: 15pt;"><tr>...</tr></table>`.
   - Stamp placeholders should be centered: `<div style="border: 1px dashed #94a3b8; width: 70pt; height: 50pt; line-height: 50pt; text-align: center; color: #64748b; font-size: 8pt; margin: 0 auto;">STAMP</div>`.
6. NO PAGE NUMBERS & NO RUNNING FOOTERS:
   - NEVER output page numbers (do NOT write '1/3', '2/3', '3/3', '1/6', '2/6', 'Page 1 of 3', etc.). Layout pagination is handled by the PDF compiler.
   - NEVER output software or certification watermarks (do NOT output 'Powered by TCPDF', 'Official Medical Document', or 'Certified translation').
7. NO ARBITRARY GIANT SPACERS & ZERO GEORGIAN SCRIPT:
   - Do NOT insert artificial empty gap divs (no `<div style="height: 100pt"></div>`).
   - Do NOT leave huge unnatural empty white spaces. Keep content flowing naturally.
   - Output 100% English translation. Zero Georgian script.
8. RETURN ONLY CLEAN HTML:
   - Output ONLY the HTML fragment for the page body. No <html>, <head>, or <body> tags, no markdown code blocks."""


def _clean_fragment(text):
    text = text.strip()
    text = re.sub(r"^```(?:html)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    text = re.sub(r"(?is)<!DOCTYPE[^>]*>|</?html[^>]*>|<head.*?</head>|</?body[^>]*>|<style.*?</style>", "", text)
    text = re.sub(r"(?i)position\s*:\s*absolute;?", "", text)

    # 1. Strip HTML tags whose sole content is a page number (e.g. 1/3, 2/6, Page 1 of 6, etc.)
    pat_page_num_tag = re.compile(
        r'<(?P<tag>[a-z0-9]+)[^>]*>\s*(?:<b>\s*)?(?:Page\s+)?\d+\s*(?:/|of)\s*\d+(?:\s*</b>)?\s*</(?P=tag)>',
        re.IGNORECASE
    )
    text = pat_page_num_tag.sub('', text)

    # 2. Strip tags containing certification boilerplate, running footers, or TCPDF watermarks
    pat_footer_tag = re.compile(
        r'<(?P<tag>[a-z0-9]+)[^>]*>[^<]*?(?:Certified English translation|Official Medical Document|English Certified Translation|Powered by TCPDF)[^<]*?</(?P=tag)>',
        re.IGNORECASE
    )
    text = pat_footer_tag.sub('', text)

    # 3. Strip standalone lines with just page numbers e.g. '1/3' or '2/6'
    text = re.sub(r'(?m)^\s*(?:<b>\s*)?(?:Page\s+)?\d+\s*(?:/|of)\s*\d+(?:\s*</b>)?\s*$', '', text)

    # 4. Strip standalone TCPDF lines or certification lines
    text = re.sub(r'(?m)^\s*Powered by TCPDF.*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(?m)^\s*(?:Official Medical Document|Certified English translation|English Certified Translation).*$', '', text, flags=re.IGNORECASE)

    # 5. Format top-right decree note into compact right-aligned block
    def _fix_decree(match):
        prefix = match.group(1)
        inner = re.sub(r'<[^>]+>', ' ', match.group(2))
        inner = ' '.join(inner.split())
        suffix = match.group(3)
        return (f'{prefix}<div class="decree-block" style="max-width: 190pt; margin-left: auto; '
                f'text-align: right; font-size: 8.5pt; color: #475569; line-height: 1.35;">'
                f'{inner}</div>{suffix}')

    text = re.sub(
        r'(<td[^>]*>)\s*(?:<div[^>]*>)?\s*(Approved\s+by\s+(?:Order|Decree)[^<]+?(?:Minister|Georgia|2008|230)[^<]*?)\s*(?:</div>)?\s*(</td>)',
        _fix_decree,
        text,
        flags=re.IGNORECASE | re.DOTALL
    )

    # 6. Ensure clinic header box does not shrink or wrap awkwardly
    text = re.sub(
        r'(<div[^>]*class="[^"]*(?:box|clinic-card|issuer-card)[^"]*"[^>]*style=")([^"]*display:\s*inline-block;?[^"]*)(")',
        lambda m: m.group(1) + m.group(2).rstrip(';') + ('; min-width: 220pt; max-width: 280pt; box-sizing: border-box;' if 'min-width' not in m.group(2) else '') + m.group(3),
        text,
        flags=re.IGNORECASE
    )

    # 5. Remove running footer containers
    text = re.sub(r'<div[^>]*class="[^"]*footer[^"]*"[^>]*>.*?</div>', '', text, flags=re.IGNORECASE | re.DOTALL)

    # 6. Clamp or remove arbitrary spacer divs
    def _clamp_spacer(m):
        h = float(m.group(1))
        clamped = min(10.0, max(2.0, h * 0.1))
        return f'<div style="height:{clamped:.1f}pt;"></div>'
    text = re.sub(r'<div style="height:\s*([\d\.]+)pt;?"></div>', _clamp_spacer, text)

    return text.strip()


def _generate(client, contents, config, retries=3):
    last = None
    for attempt in range(retries):
        try:
            resp = client.models.generate_content(model=MODEL, contents=contents, config=config)
            if resp.text:
                return resp.text
            last = RuntimeError("empty response")
        except Exception as exc:
            last = exc
        time.sleep(2 * (attempt + 1))
    raise last


def translate_page(client, page, total_pages):
    raw_text = page.get("raw_text", "").strip()
    if raw_text and any(ord(c) >= 0x10A0 and ord(c) <= 0x10FF for c in raw_text):
        trans_contents = [
            types.Part.from_bytes(data=page["preview_png"], mime_type="image/png"),
            f"Georgian page text:\n{raw_text}\n\nTranslate this entire medical page into professional clinical English.",
        ]
        trans_config = types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT_TRANSLATION,
            temperature=0.0,
            max_output_tokens=32768,
        )
        english_translation = _generate(client, trans_contents, trans_config)
    else:
        english_translation = "All text on this page is already in Latin/numerical characters or from image."

    layout_info = layout_map_text(page)
    builder_prompt = (
        f"This is page {page['number']} of {total_pages}.\n\n"
        f"VERIFIED ENGLISH CLINICAL TRANSLATION OF THIS PAGE:\n{english_translation}\n\n"
        f"DETECTED ASSETS & LAYOUT MAP:\n{layout_info}\n\n"
        "Look carefully at the attached page image to understand its exact visual layout. "
        "Recreate this page faithfully in clean print-ready HTML matching the original layout and structure: "
        "side-by-side header cards, centered uppercase titles, clean labeled form items with bold filled values, "
        "and properly structured tables. Use ONLY the English translation—zero Georgian script in output."
    )

    single_pdf = page.get("pdf")
    html_contents = [
        types.Part.from_bytes(data=page["preview_png"], mime_type="image/png"),
    ]
    if single_pdf:
        html_contents.append(types.Part.from_bytes(data=single_pdf, mime_type="application/pdf"))
    html_contents.append(builder_prompt)
    html_config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT_HTML_BUILDER,
        temperature=0.0,
        max_output_tokens=32768,
    )
    fragment = _generate(client, html_contents, html_config)
    return _clean_fragment(fragment)


def inject_images(html, images):
    for im in images:
        b64 = base64.b64encode(im["png"]).decode("ascii")
        tag = (f'<img class="doc-img" src="data:image/png;base64,{b64}" '
               f'style="max-width:160pt; max-height:85pt; object-fit:contain;" alt="" />')
        pattern = re.compile(r"<!--\s*IMG\s*:\s*" + re.escape(im["id"]) + r"\s*-->")
        if pattern.search(html):
            html = pattern.sub(lambda _m: tag, html, count=1)
            html = pattern.sub("", html)
        else:
            align = "left" if im["x"] < 200 else ("right" if im["x"] > 330 else "center")
            block = f'<div style="text-align:{align}; margin:4pt 0;">{tag}</div>'
            html = block + html if im["role"] == "header logo" else html + block
    return re.sub(r"<!--\s*(IMG\s*:[^>]*|LOGO)\s*-->", "", html)


def build_html(fragments, profile):
    t, r, b, l = profile["margins_mm"]
    base = profile["base_pt"]
    body = "\n".join(f'<section class="page">\n{f}\n</section>' for f in fragments)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <style>
    @media print {{
      @page {{
        size: {profile['page_w_mm']:.1f}mm {profile['page_h_mm']:.1f}mm;
        margin: {t:.1f}mm {r:.1f}mm {b:.1f}mm {l:.1f}mm;
      }}
      body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        font-size: {base}pt; line-height: 1.35; color: #0f172a; margin: 0; padding: 0; background: transparent;
      }}
      section.page {{
        margin: 0;
        padding: 0;
        background: transparent;
        box-shadow: none;
        border-radius: 0;
        width: auto;
        min-height: auto;
        break-after: auto;
      }}
    }}
    
    @media screen {{
      body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        font-size: {base}pt; line-height: 1.35; color: #0f172a; margin: 0;
        background-color: #f1f5f9;
        padding: 24px 12px;
        display: flex;
        flex-direction: column;
        align-items: center;
      }}
      section.page {{
        background-color: #ffffff;
        width: {profile['page_w_mm']:.1f}mm;
        max-width: 100%;
        min-height: auto;
        padding: {t:.1f}mm {r:.1f}mm {b:.1f}mm {l:.1f}mm;
        margin: 0 auto 24px auto;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08), 0 1px 3px rgba(0, 0, 0, 0.05);
        box-sizing: border-box;
        border-radius: 3px;
      }}
    }}
    
    html {{ font-size: {base}pt; }}
    
    /* Clean, modern headings */
    h1, h2, h3, h4 {{ color: #0f172a; margin: 6pt 0 4pt 0; font-weight: 600; line-height: 1.25; }}
    h1 {{ font-size: 1.45em; text-align: center; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1.5pt solid #cbd5e1; padding-bottom: 4pt; }}
    h2 {{ font-size: 1.2em; color: #1e3a8a; border-bottom: 0.75pt solid #e2e8f0; padding-bottom: 2pt; margin-top: 10pt; }}
    h3, h4 {{ font-size: 1.05em; color: #1e293b; }}
    
    p {{ margin: 3pt 0 5pt 0; }}
    ul, ol {{ margin: 3pt 0 6pt 0; padding-left: 16pt; }}
    li {{ margin-bottom: 2pt; }}
    
    /* Form items & Questionnaires (Form 100) */
    .form-item {{
      break-inside: auto;
      margin-bottom: 12pt;
      line-height: 1.5;
    }}
    .form-item > span:first-child, .form-item > div:first-child {{
      color: #334155;
      font-size: 9.5pt;
      margin-bottom: 2pt;
    }}
    .signature-block, .box, .diag-block {{
      break-inside: avoid;
    }}
    
    /* Layout table for header cards & metadata */
    table {{ width: 100%; border-collapse: collapse; margin: 6pt 0; }}
    td, th {{ vertical-align: top; word-wrap: break-word; }}
    table.layout {{ width: 100%; table-layout: fixed; }}
    table.layout td, table.layout th {{ border: none; padding: 2pt 4pt 2pt 0; vertical-align: top; }}
    
    /* Dedicated header classes for Form 100 & clinical reports */
    .clinic-card, .issuer-card, table.layout td .box {{
      border: 1pt solid #94a3b8;
      background-color: #f8fafc;
      border-radius: 4pt;
      padding: 6pt 10pt;
      min-width: 220pt;
      max-width: 280pt;
      display: inline-block;
      line-height: 1.35;
      box-sizing: border-box;
    }}
    .decree-block, .approval-note {{
      display: inline-block;
      max-width: 190pt;
      margin-left: auto;
      text-align: right;
      font-size: 8.5pt;
      color: #475569;
      line-height: 1.35;
    }}
    
    /* Executive styled clinical data tables & financial summary grids */
    table.data-table, table.summary-table, table.summary-grid {{
      width: 100%;
      border-collapse: collapse;
      margin: 6pt 0;
      border: 1pt solid #cbd5e1;
      font-size: 8.5pt;
    }}
    table.data-table thead, table.summary-table thead, table.summary-grid thead {{ display: table-header-group; }}
    table.data-table th, table.summary-table th, table.summary-grid th {{
      background-color: #f1f5f9;
      color: #0f172a;
      font-weight: 600;
      border: 0.75pt solid #cbd5e1;
      padding: 3.5pt 5pt;
      text-align: left;
    }}
    table.data-table td, table.summary-table td, table.summary-grid td {{
      border: 0.75pt solid #e2e8f0;
      padding: 3pt 5pt;
      color: #1e293b;
    }}
    table.data-table tr:nth-child(even) td, table.summary-table tr:nth-child(even) td, table.summary-grid tr:nth-child(even) td {{
      background-color: #f8fafc;
    }}
    table.data-table td:empty::after, table.data-table th:empty::after,
    table.summary-table td:empty::after, table.summary-grid td:empty::after {{ content: "\\00a0"; }}
    table.data-table tr.group-row td {{
      background-color: #e0e7ff;
      color: #1e3a8a;
      font-weight: bold;
      text-align: center;
      border-top: 1pt solid #cbd5e1;
      border-bottom: 1pt solid #cbd5e1;
    }}
    
    /* Summary rows / Grand totals */
    tr.total-row td, tr.grand-total td, .summary-table tr:last-child td, .summary-grid tr:last-child td {{
      font-weight: bold;
      background-color: #eff6ff !important;
      color: #1e3a8a;
      border-top: 1.25pt solid #94a3b8;
    }}

    /* Prevent awkward wrapping on numbers and currency */
    td.nowrap, .nowrap, td.amount, td.price, td:last-child {{
      white-space: nowrap;
    }}

    /* Any unclassified table inside a box or summary container gets borders */
    .box table td, .summary-block table td {{
      border: 0.5pt solid #cbd5e1;
      padding: 2.5pt 4pt;
    }}
    
    tr {{ break-inside: avoid; }}
    .box {{
      border: 1pt solid #cbd5e1;
      background-color: #f8fafc;
      border-radius: 3pt;
      padding: 5pt 8pt;
      margin: 6pt 0;
    }}
    .fill {{ border-bottom: 0.75pt solid #475569; padding: 0 4pt; font-weight: 500; }}
    img.doc-img {{ vertical-align: middle; }}
  </style>
</head>
<body>
{body}
</body>
</html>"""


def html_to_pdf(full_html):
    buf = io.BytesIO()
    HTML(string=full_html).write_pdf(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 3. Optional Structured Data Digitization (Excel)
# ---------------------------------------------------------------------------
SYSTEM_PROMPT_EXCEL = (
    "You are a universal medical document parser.\n"
    "Extract EVERY piece of data from this document into two collections:\n"
    "1. 'fields': Every single numbered point, header metadata item, patient detail, or key-value pair. "
    "Split inline rows into separate fields.\n"
    "2. 'tables': Every structured grid, laboratory test row, or procedural line without skipping any rows.\n"
    "Keep all numbers, dates, IDs, and codes verbatim."
)


def extract_excel(client, pdf_bytes):
    text = _generate(
        client,
        [types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
         "Extract all fields and tables universally into structured JSON."],
        types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT_EXCEL,
            temperature=0.0,
            max_output_tokens=65536,
            response_mime_type="application/json",
            response_schema=UniversalDocumentData,
        ),
    )
    data = json.loads(text)

    df_fields = pd.DataFrame([
        {
            "Point / Section (# / Section)": f.get("section_or_item_number", ""),
            "Field Label (Georgian)": f.get("field_label_ge", ""),
            "Field Label (English)": f.get("field_label_en", ""),
            "Value (Georgian)": f.get("field_value_ge", ""),
            "Value (English)": f.get("field_value_en", ""),
        }
        for f in data.get("fields", [])
    ])
    df_tables = pd.DataFrame([
        {
            "Table": r.get("table_name", ""),
            "№": r.get("row_number", ""),
            "Date / Subgroup": r.get("group_or_date", ""),
            "Code": r.get("code", ""),
            "Description (Georgian)": r.get("name_ge", ""),
            "Description (English)": r.get("name_en", ""),
            "Result / Qty": r.get("result_or_qty", ""),
            "Ref Range / Unit": r.get("unit_or_reference_range", ""),
            "Status / Notes": r.get("notes_or_status", ""),
        }
        for r in data.get("tables", [])
    ])

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_fields.to_excel(writer, sheet_name="Document_Fields", index=False)
        if not df_tables.empty:
            df_tables.to_excel(writer, sheet_name="Table_Data", index=False)

    excel_buffer.seek(0)
    wb = openpyxl.load_workbook(excel_buffer)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    side = Side(style="thin", color="D9D9D9")
    thin_border = Border(left=side, right=side, top=side, bottom=side)

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 26
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column), start=2):
            for cell in row:
                if row_idx % 2 == 0:
                    cell.fill = zebra_fill
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        for col in ws.columns:
            max_len = max((len(str(c.value or "").split("\n")[0]) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 4, 14), 45)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), df_fields, df_tables


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def translate_to_pdf(client, pdf_bytes, progress=None):
    pages, profile = analyze_pdf(pdf_bytes)
    fragments = [None] * len(pages)
    
    ctx = None
    try:
        from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx, add_script_run_ctx
        ctx = get_script_run_ctx()
    except Exception:
        add_script_run_ctx = None

    def _worker(i, p):
        if ctx and add_script_run_ctx:
            add_script_run_ctx(ctx=ctx)
        return translate_page(client, p, len(pages))

    with cf.ThreadPoolExecutor(max_workers=min(4, len(pages))) as ex:
        futs = {ex.submit(_worker, i, p): i for i, p in enumerate(pages)}
        for done, fut in enumerate(cf.as_completed(futs), 1):
            i = futs[fut]
            fragments[i] = inject_images(fut.result(), pages[i]["images"])
            if progress:
                try:
                    progress(done / len(pages))
                except Exception:
                    pass
    full_html = build_html(fragments, profile)
    return html_to_pdf(full_html), full_html


def process_document(pdf_bytes, progress=None, with_excel=False):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        try:
            import streamlit as st
            api_key = st.secrets.get("GEMINI_API_KEY")
            if api_key:
                os.environ["GEMINI_API_KEY"] = api_key
        except Exception:
            pass
    client_kwargs = {"http_options": {"api_version": "v1beta"}}
    if api_key:
        client_kwargs["api_key"] = api_key
    client = genai.Client(**client_kwargs)
    ctx = None
    try:
        from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx, add_script_run_ctx
        ctx = get_script_run_ctx()
    except Exception:
        add_script_run_ctx = None

    def _pdf_worker():
        if ctx and add_script_run_ctx:
            add_script_run_ctx(ctx=ctx)
        return translate_to_pdf(client, pdf_bytes, progress)

    def _xls_worker():
        if ctx and add_script_run_ctx:
            add_script_run_ctx(ctx=ctx)
        return extract_excel(client, pdf_bytes)

    if with_excel:
        with cf.ThreadPoolExecutor(max_workers=2) as ex:
            f_pdf = ex.submit(_pdf_worker)
            f_xls = ex.submit(_xls_worker)
            pdf_data, full_html = f_pdf.result()
            excel_data, df_fields, df_tables = f_xls.result()
    else:
        pdf_data, full_html = _pdf_worker()
        excel_data, df_fields, df_tables = None, pd.DataFrame(), pd.DataFrame()

    return pdf_data, full_html, excel_data, df_fields, df_tables


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------
def run_cli(argv):
    out_dir = "translated_out"
    if "--out" in argv:
        i = argv.index("--out")
        out_dir = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    with_excel = "--excel" in argv
    files = [a for a in argv if a.lower().endswith(".pdf")]
    os.makedirs(out_dir, exist_ok=True)
    for path in files:
        stem = os.path.splitext(os.path.basename(path))[0]
        print(f"Processing {ascii(path)} (Excel={with_excel}) ...", flush=True)
        t0 = time.time()
        with open(path, "rb") as fh:
            pdf_data, html, xls, _, _ = process_document(fh.read(), with_excel=with_excel)
        with open(os.path.join(out_dir, f"{stem}_EN.pdf"), "wb") as fh:
            fh.write(pdf_data)
        with open(os.path.join(out_dir, f"{stem}_EN.html"), "w", encoding="utf-8") as fh:
            fh.write(html)
        if xls:
            with open(os.path.join(out_dir, f"{stem}_data.xlsx"), "wb") as fh:
                fh.write(xls)
        print(f"  done in {time.time() - t0:.1f}s", flush=True)


def run_streamlit():
    import streamlit as st
    import streamlit.components.v1 as components

    st.set_page_config(page_title="Medical Document Translator", page_icon="🏥", layout="wide", initial_sidebar_state="expanded")

    saved_key = os.environ.get("GEMINI_API_KEY", "")
    if not saved_key:
        try:
            saved_key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            pass
    if saved_key:
        os.environ["GEMINI_API_KEY"] = saved_key

    # Sidebar: API Key and Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        user_key = st.text_input("Gemini API Key", value=saved_key, type="password", key="side_key",
                                 help="Leave blank if set in environment or Streamlit Secrets.")
        if user_key:
            os.environ["GEMINI_API_KEY"] = user_key

        st.markdown("---")
        st.markdown("### 📋 Sample Documents")
        sample_dir = os.path.join(os.path.dirname(__file__), "sample_documents")
        sample_files = []
        if os.path.exists(sample_dir):
            sample_files = [f for f in os.listdir(sample_dir) if f.lower().endswith(".pdf")]
        
        selected_sample = None
        if sample_files:
            chosen = st.selectbox("Quick Test with Sample PDF:", ["None (Upload my own)"] + sample_files, key="side_sample")
            if chosen != "None (Upload my own)":
                selected_sample = os.path.join(sample_dir, chosen)
                st.caption(f"Loaded: `{chosen}`")

        st.markdown("---")
        st.info("🇬🇪 **Georgian ➔ 🇬🇧 English**\n\nTransforms complex medical invoices, Form #100 certificates, and laboratory panels into certified-grade English PDFs.")
        st.markdown("<div style='text-align: center; color: #64748b; font-size: 0.85em; margin-top: 15px;'>Made with ❤️ by Gigi</div>", unsafe_allow_html=True)

    st.title("🏥 Georgian Medical Document Translation Pipeline")
    st.write("Convert complex Georgian clinical documents into professional English PDFs with executive medical styling.")

    def _render_footer():
        st.markdown("<br><hr style='border: none; border-top: 1px solid #e2e8f0; margin: 30px 0 10px 0;'><div style='text-align: center; color: #64748b; font-size: 0.9em; font-weight: 500;'>Made with ❤️ by Gigi</div>", unsafe_allow_html=True)

    # Input selection: upload or sample
    pdf_bytes = None
    file_name = "translated_document"

    if selected_sample and os.path.exists(selected_sample):
        with open(selected_sample, "rb") as sf:
            pdf_bytes = sf.read()
        file_name = os.path.basename(selected_sample)
        st.success(f"Selected sample document: **{file_name}**")
    else:
        uploaded_file = st.file_uploader("Upload a Medical PDF (Invoice, Form #100, Lab Report)", type=["pdf"])
        if uploaded_file is not None:
            pdf_bytes = uploaded_file.read()
            file_name = uploaded_file.name

    if not pdf_bytes:
        st.info("👆 Please upload a medical PDF above or select a sample document from the sidebar to begin.")
        _render_footer()
        return

    col_opt1, col_opt2 = st.columns([1, 3])
    with col_opt1:
        include_excel = st.checkbox("Also extract to Excel (.xlsx)", value=False,
                                    help="Unchecked by default for fast translation (~5-10s). Check if you need structured Excel tables.")

    if not st.button("🚀 Translate & Generate English Document", type="primary", use_container_width=True):
        _render_footer()
        return

    if not os.environ.get("GEMINI_API_KEY"):
        st.error("Missing Gemini API Key! Please enter your key in the sidebar or set it in your environment/secrets.")
        return

    bar = st.progress(0.0, text="Translating medical text and applying executive styling...")
    t0 = time.time()
    try:
        with st.spinner("Processing document layout and clinical terminology..."):
            pdf_data, full_html, excel_data, df_fields, df_tables = process_document(
                pdf_bytes,
                progress=lambda f: bar.progress(f, text=f"Processed {f:.0%} of pages"),
                with_excel=include_excel
            )
    except Exception as exc:
        bar.empty()
        st.error(f"❌ Error during document processing: {str(exc)}")
        import traceback
        st.code(traceback.format_exc(), language="python")
        return
    bar.empty()
    duration = time.time() - t0
    st.success(f"✅ Translation completed in {duration:.1f} seconds!")

    stem = os.path.splitext(file_name)[0]
    safe_stem = re.sub(r'[\s]+', '_', stem.strip())

    # Action buttons
    if include_excel and excel_data:
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("📥 Download English PDF", data=pdf_data, file_name=f"{safe_stem}_EN.pdf",
                               mime="application/pdf", use_container_width=True)
        with col_dl2:
            st.download_button("📊 Download Structured Excel (.xlsx)", data=excel_data, file_name=f"{safe_stem}_data.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
    else:
        st.download_button("📥 Download English PDF", data=pdf_data, file_name=f"{safe_stem}_EN.pdf",
                           mime="application/pdf", use_container_width=True)

    # Mobile Safari direct Blob viewer (solves iOS Safari download manager issues)
    b64_pdf = base64.b64encode(pdf_data).decode("ascii")
    components.html(
        f"""
        <div style="margin: 4px 0 10px 0; text-align: center;">
            <a id="iosOpenBtn" href="#" target="_blank" style="
                display: block;
                background: linear-gradient(135deg, #1e3a8a, #2563eb);
                color: #ffffff;
                text-align: center;
                padding: 13px 20px;
                border-radius: 8px;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                font-size: 15px;
                font-weight: 600;
                text-decoration: none;
                box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            ">📱 Tap Here to Open PDF in New Tab (iPhone / Safari)</a>
        </div>
        <script>
        try {{
            const b64 = "{b64_pdf}";
            const byteChars = atob(b64);
            const byteNums = new Array(byteChars.length);
            for (let i = 0; i < byteChars.length; i++) {{
                byteNums[i] = byteChars.charCodeAt(i);
            }}
            const byteArray = new Uint8Array(byteNums);
            const blob = new Blob([byteArray], {{ type: 'application/pdf' }});
            const blobUrl = URL.createObjectURL(blob);
            const btn = document.getElementById("iosOpenBtn");
            btn.href = blobUrl;
        }} catch(e) {{
            console.error("Blob URL generation error:", e);
        }}
        </script>
        """,
        height=65
    )

    # Document Preview Tabs
    st.markdown("### 📄 Document Preview")
    tabs = st.tabs(["Interactive Document Preview", "HTML Source"] + (["Structured Data"] if (include_excel and not df_fields.empty) else []))
    
    with tabs[0]:
        components.html(full_html, height=800, scrolling=True)

    with tabs[1]:
        st.code(full_html, language="html")

    if include_excel and not df_fields.empty and len(tabs) > 2:
        with tabs[2]:
            subtab1, subtab2 = st.tabs(["Document Fields", "Table Data"])
            with subtab1:
                st.dataframe(df_fields, use_container_width=True)
            with subtab2:
                if not df_tables.empty:
                    st.dataframe(df_tables, use_container_width=True)
                else:
                    st.info("No multi-row tables detected in this document.")

    _render_footer()


if "--cli" in sys.argv:
    run_cli(sys.argv[1:])
else:
    run_streamlit()
