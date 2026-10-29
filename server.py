"""
FileShift - File Converter Backend
High-fidelity conversion engine — preserves tables, styles, formatting.

Strategy per route:
  DOCX → PDF   : docx2pdf (drives MS Word on Windows) → falls back to html-chain
  DOCX → MD/TXT: python-docx (structure-aware extraction)
  MD   → PDF   : MD→HTML→PDF via weasyprint (or pymupdf html rendering)
  MD   → DOCX  : python-docx with proper heading/table/code mapping
  PDF  → *     : pymupdf (best-in-class text/layout extraction)
  HTML → *     : beautifulsoup4 structure-aware extraction
  TXT  → *     : plain text, no loss possible

Run: python server.py  (port 8080)
"""

import html as html_mod
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

# pyrefly: ignore [missing-import]
import fitz  # pymupdf
# pyrefly: ignore [missing-import]
import markdown as md_lib
# pyrefly: ignore [missing-import]
from bs4 import BeautifulSoup
# pyrefly: ignore [missing-import]
from docx import Document
# pyrefly: ignore [missing-import]
from docx.shared import Pt, RGBColor, Inches
# pyrefly: ignore [missing-import]
from docx.oxml.ns import qn
# pyrefly: ignore [missing-import]
from docx.oxml import OxmlElement
# pyrefly: ignore [missing-import]
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
# pyrefly: ignore [missing-import]
from fastapi.middleware.cors import CORSMiddleware
# pyrefly: ignore [missing-import]
from fastapi.responses import FileResponse, HTMLResponse
# pyrefly: ignore [missing-import]
import uvicorn

app = FastAPI(title="FileShift")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

UPLOAD_DIR = Path(tempfile.gettempdir()) / "fileshift"
UPLOAD_DIR.mkdir(exist_ok=True)

CONVERSIONS = {
    "pdf":  ["md", "docx", "txt"],
    "md":   ["pdf", "docx", "txt", "html"],
    "docx": ["pdf", "md", "txt"],
    "txt":  ["pdf", "md", "docx"],
    "html": ["pdf", "md", "docx"],
}

# ─── CSS for HTML→PDF rendering ───────────────────────────────────────────────
# NOTE: pymupdf Story has limited CSS support. Rules that work reliably:
#   - font-family, font-size, font-weight, font-style, color
#   - margin, padding, line-height
#   - border (simple), background-color ON BLOCK elements only
#   - width on table/td/th
# Rules that BREAK pymupdf rendering:
#   - thead tr { background } / tbody tr:nth-child → applies to char spans → black boxes
#   - @import (google fonts) → silently ignored
#   - complex selectors
# Solution: inject inline styles on th/td directly in HTML post-processing

PDF_CSS = """
body { font-family: Helvetica, Arial, sans-serif; font-size: 10.5pt; line-height: 1.65; color: #1a1a1a; }
h1 { font-size: 20pt; font-weight: bold; color: #111111; margin-top: 24pt; margin-bottom: 8pt; }
h2 { font-size: 15pt; font-weight: bold; color: #c0392b; margin-top: 18pt; margin-bottom: 6pt; }
h3 { font-size: 12pt; font-weight: bold; color: #2c3e50; margin-top: 14pt; margin-bottom: 4pt; }
h4 { font-size: 11pt; font-weight: bold; color: #2c3e50; margin-top: 10pt; margin-bottom: 3pt; }
h5, h6 { font-size: 10pt; font-weight: bold; margin-top: 8pt; margin-bottom: 2pt; }
p { margin: 0 0 7pt; }
ul { margin: 4pt 0 6pt; padding-left: 18pt; }
ol { margin: 4pt 0 6pt; padding-left: 18pt; }
li { margin-bottom: 3pt; line-height: 1.55; }
/* NOTE: background-color for code/pre/blockquote is intentionally NOT set here.
   pymupdf's Story renderer applies stylesheet-level background-color to individual
   character spans instead of the element block, painting solid dark bars over text.
   These elements get their background-color injected as an INLINE style instead —
   see _apply_table_inline_styles(). Only non-background properties are safe to keep
   at the stylesheet level, as a font/spacing fallback for any element the inliner misses. */
code { font-family: Courier, monospace; font-size: 9pt; color: #c0392b; padding: 1pt 3pt; }
pre { font-family: Courier, monospace; font-size: 8.5pt; color: #333333; padding: 10pt 12pt; margin: 8pt 0; white-space: pre-wrap; }
pre code { color: #333333; padding: 0; }
blockquote { border-left: 3pt solid #c0392b; margin: 6pt 0; padding: 4pt 10pt; color: #555555; font-style: italic; }
strong { font-weight: bold; }
em { font-style: italic; }
hr { border-top: 1pt solid #cccccc; margin: 14pt 0; }
a { color: #2563eb; }
"""

WEASYPRINT_AVAILABLE = False
try:
    
    # pyrefly: ignore [missing-import]
    from weasyprint import HTML as WP_HTML, CSS as WP_CSS
    import warnings as _warnings
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")
        _test = io.BytesIO()
        WP_HTML(string='<p>test</p>').write_pdf(_test)
        if _test.tell() > 0:
            WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

# ─── Helpers ──────────────────────────────────────────────────────────────────

def style_name(para) -> str:
    try:
        return (para.style.name or '').lower()
    except Exception:
        return ''

def _apply_table_inline_styles(html: str) -> str:
    """
    Inject inline styles on every element that carries a background-color.

    Root cause of the "black box" bug: pymupdf's Story renderer does not apply
    CSS selector rules (thead tr{background}, code{background}, blockquote{background}, etc.)
    to the element block — it applies them to each individual character span instead,
    which paints a filled rectangle behind every glyph. Depending on stacking/order
    this shows up as solid dark bars over the text (headers, inline code, code
    blocks, and blockquotes were all affected — anything with a stylesheet-level
    background-color).

    Fix: never rely on the <style> stylesheet for background-color. Inject the
    color as an inline `style="..."` attribute directly on each element instead —
    pymupdf (and weasyprint) both apply inline styles correctly as a single block
    fill. Done with BeautifulSoup rather than string replace so nested elements
    (e.g. <code> inside <pre>) are handled correctly and cannot double-fill.
    """
    soup = BeautifulSoup(html, 'html.parser')

    for table in soup.find_all('table'):
        table['style'] = 'width:100%;border-collapse:collapse;margin:10pt 0 14pt;font-size:9.5pt;'
    for thead in soup.find_all('thead'):
        thead['style'] = 'background-color:#2c3e50;color:#ffffff;'
    for th in soup.find_all('th'):
        th['style'] = ('background-color:#2c3e50;color:#ffffff;padding:7pt 10pt;'
                        'text-align:left;font-weight:bold;border:1px solid #2c3e50;')
    # Alternate row coloring must be done via explicit style since nth-child is unsupported
    for td in soup.find_all('td'):
        td['style'] = 'padding:6pt 10pt;border:1px solid #dee2e6;vertical-align:top;'

    # Code blocks (<pre><code>...</code></pre>) — style the <pre>, and make the
    # inner <code> transparent so it never paints a second, misaligned fill.
    for pre in soup.find_all('pre'):
        pre['style'] = ('font-family:Courier, monospace;font-size:8.5pt;background-color:#f4f4f5;'
                         'color:#333333;padding:10pt 12pt;margin:8pt 0;white-space:pre-wrap;')
        for code in pre.find_all('code'):
            code['style'] = 'background-color:transparent;color:#333333;padding:0;font-family:Courier, monospace;'

    # Inline code (not inside a <pre>)
    for code in soup.find_all('code'):
        if code.find_parent('pre') is None:
            code['style'] = ('font-family:Courier, monospace;font-size:9pt;background-color:#f4f4f5;'
                              'color:#c0392b;padding:1pt 3pt;')

    # Blockquotes
    for bq in soup.find_all('blockquote'):
        bq['style'] = ('border-left:3pt solid #c0392b;margin:6pt 0;padding:4pt 10pt;'
                        'background-color:#fdf2f2;color:#555555;font-style:italic;')

    return str(soup)

def md_to_html_string(md_text: str) -> str:
    """Convert markdown to a full HTML document with PDF-safe CSS."""
    # Do NOT use codehilite (adds dark div wrappers that break pymupdf)
    # Do NOT use nl2br (inserts <br> inside table cells, breaking layout)
    # Do NOT use toc (adds anchor IDs that clutter output)
    body = md_lib.markdown(
        md_text,
        extensions=['tables', 'fenced_code', 'sane_lists']
    )
    body = _apply_table_inline_styles(body)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<style>{PDF_CSS}</style>
</head>
<body>{body}</body>
</html>"""

def html_file_to_pdf_weasyprint(html_path: Path, out: Path):
    html_content = html_path.read_text(encoding='utf-8', errors='replace')
    if '<style>' not in html_content:
        html_content = html_content.replace('</head>', f'<style>{PDF_CSS}</style></head>', 1)
    html_content = _apply_table_inline_styles(html_content)
    WP_HTML(string=html_content, base_url=str(html_path.parent)).write_pdf(
        str(out),
        stylesheets=[WP_CSS(string='@page { size: A4; margin: 20mm 22mm; }')]
    )

def html_string_to_pdf(html_str: str, out: Path, base_url: str = None):
    if WEASYPRINT_AVAILABLE:
        WP_HTML(string=html_str, base_url=base_url).write_pdf(
            str(out),
            stylesheets=[WP_CSS(string='@page { size: A4; margin: 20mm 22mm; }')]
        )
    else:
        # Fallback: pymupdf Story (fitz) — good table support
        _html_to_pdf_pymupdf(html_str, out)

def _html_to_pdf_pymupdf(html_str: str, out: Path):
    """Pymupdf Story renderer — decent HTML→PDF without system libs."""
    try:
        # pyrefly: ignore [missing-import]
        import fitz
        story = fitz.Story(html=html_str, user_css=PDF_CSS)
        writer = fitz.DocumentWriter(str(out))
        mediabox = fitz.paper_rect("a4")
        where = mediabox + (36, 36, -36, -36)  # margins
        more = True
        while more:
            dev = writer.begin_page(mediabox)
            more, _ = story.place(where)
            story.draw(dev)
            writer.end_page()
        writer.close()
    except Exception as e:
        raise RuntimeError(f"pymupdf Story rendering failed: {e}")

def docx_to_pdf_word(src: Path, out: Path) -> bool:
    """Try to use MS Word via docx2pdf (Windows/Mac only)."""
    try:
        # pyrefly: ignore [missing-import]
        from docx2pdf import convert
        convert(str(src), str(out))
        return out.exists() and out.stat().st_size > 0
    except Exception:
        return False

def docx_to_pdf_libreoffice(src: Path, out: Path) -> bool:
    """Try LibreOffice headless (Linux/Mac)."""
    try:
        lo_bins = ['libreoffice', 'soffice', '/usr/bin/libreoffice', '/usr/bin/soffice']
        lo = next((b for b in lo_bins if shutil.which(b)), None)
        if not lo:
            return False
        result = subprocess.run(
            [lo, '--headless', '--convert-to', 'pdf', '--outdir', str(out.parent), str(src)],
            capture_output=True, timeout=60
        )
        # LibreOffice outputs <stem>.pdf
        expected = out.parent / (src.stem + '.pdf')
        if expected.exists():
            shutil.move(str(expected), str(out))
        return out.exists() and out.stat().st_size > 0
    except Exception:
        return False

def docx_to_html_string(src: Path) -> str:
    """Convert docx to HTML preserving tables, headings, lists, formatting."""
    doc = Document(str(src))
    parts = ['<!DOCTYPE html><html><head><meta charset="utf-8"/>',
             f'<style>{PDF_CSS}</style></head><body>']

    def run_html(run):
        text = html_mod.escape(run.text or '')
        if not text:
            return ''
        if run.bold and run.italic:
            text = f'<strong><em>{text}</em></strong>'
        elif run.bold:
            text = f'<strong>{text}</strong>'
        elif run.italic:
            text = f'<em>{text}</em>'
        if run.underline:
            text = f'<u>{text}</u>'
        try:
            if run.font.name and 'mono' in run.font.name.lower() or \
               run.font.name and 'courier' in run.font.name.lower() or \
               run.font.name and 'code' in run.font.name.lower():
                text = f'<code>{text}</code>'
        except Exception:
            pass
        return text

    # Track table processing
    processed_tables = set()

    def para_html(para):
        sn = style_name(para)
        runs_html = ''.join(run_html(r) for r in para.runs)
        if not runs_html.strip() and not para.text.strip():
            return '<p>&nbsp;</p>'
        if 'heading 1' in sn: return f'<h1>{runs_html}</h1>'
        if 'heading 2' in sn: return f'<h2>{runs_html}</h2>'
        if 'heading 3' in sn: return f'<h3>{runs_html}</h3>'
        if 'heading 4' in sn: return f'<h4>{runs_html}</h4>'
        if 'heading 5' in sn: return f'<h5>{runs_html}</h5>'
        if 'list bullet' in sn: return f'<li>{runs_html}</li>'
        if 'list number' in sn: return f'<li>{runs_html}</li>'
        return f'<p>{runs_html}</p>'

    def table_html(table):
        rows = table.rows
        if not rows:
            return ''
        out = ['<table>']
        for i, row in enumerate(rows):
            out.append('<tr>')
            for cell in row.cells:
                tag = 'th' if i == 0 else 'td'
                cell_text = html_mod.escape(cell.text or '')
                out.append(f'<{tag}>{cell_text}</{tag}>')
            out.append('</tr>')
        out.append('</table>')
        return '\n'.join(out)

    # Walk body elements in order
    # pyrefly: ignore [missing-import]
    from docx.oxml.ns import qn as _qn
    body = doc.element.body
    in_ul = False
    in_ol = False

    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            # Find matching paragraph object
            para = None
            for p in doc.paragraphs:
                if p._element is child:
                    para = p
                    break
            if para is None:
                continue
            sn = style_name(para)
            is_bullet = 'list bullet' in sn
            is_number = 'list number' in sn

            if is_bullet:
                if not in_ul:
                    if in_ol: parts.append('</ol>'); in_ol = False
                    parts.append('<ul>'); in_ul = True
            elif is_number:
                if not in_ol:
                    if in_ul: parts.append('</ul>'); in_ul = False
                    parts.append('<ol>'); in_ol = True
            else:
                if in_ul: parts.append('</ul>'); in_ul = False
                if in_ol: parts.append('</ol>'); in_ol = False

            parts.append(para_html(para))

        elif tag == 'tbl':
            if in_ul: parts.append('</ul>'); in_ul = False
            if in_ol: parts.append('</ol>'); in_ol = False
            # Find matching table
            for tbl in doc.tables:
                if tbl._element is child:
                    parts.append(table_html(tbl))
                    break

    if in_ul: parts.append('</ul>')
    if in_ol: parts.append('</ol>')
    parts.append('</body></html>')
    return '\n'.join(parts)


# ─── CONVERSION FUNCTIONS ──────────────────────────────────────────────────────

# ── PDF → * ───────────────────────────────────────────────────────────────────
def pdf_to_md(src: Path, out: Path):
    doc = fitz.open(str(src))
    pages = []
    for page in doc:
        # Use dict extraction for better structure
        blocks = page.get_text("blocks")
        lines = []
        for b in sorted(blocks, key=lambda b: (b[1], b[0])):  # sort by y then x
            text = b[4].strip()
            if text:
                lines.append(text)
        pages.append('\n\n'.join(lines))
    doc.close()
    out.write_text('\n\n---\n\n'.join(pages), encoding='utf-8')

def pdf_to_txt(src: Path, out: Path):
    doc = fitz.open(str(src))
    pages = [page.get_text('text') for page in doc]
    doc.close()
    out.write_text('\n\n'.join(pages), encoding='utf-8')

def pdf_to_docx(src: Path, out: Path):
    # pyrefly: ignore [missing-import]
    from pdf2docx import Converter
    cv = Converter(str(src))
    cv.convert(str(out))
    cv.close()


# ── MD → * ────────────────────────────────────────────────────────────────────
def md_to_pdf(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    html_str = md_to_html_string(md_text)
    html_string_to_pdf(html_str, out)

def md_to_docx(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    # Convert MD→HTML→parse→docx for proper table and formatting support
    body_html = md_lib.markdown(md_text, extensions=['tables', 'fenced_code', 'sane_lists', 'nl2br'])
    soup = BeautifulSoup(body_html, 'html.parser')
    doc = Document()

    # Set default font

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def add_run_formatted(para, text, bold=False, italic=False, code=False):
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        if code:
            run.font.name = 'Courier New'
            run.font.size = Pt(9)

    def process_inline(para, element):
        """Process inline elements (bold, italic, code, text)."""
        if isinstance(element, str):
            if element:
                para.add_run(element)
        elif element.name == 'strong':
            run = para.add_run(element.get_text())
            run.bold = True
        elif element.name == 'em':
            run = para.add_run(element.get_text())
            run.italic = True
        elif element.name == 'code':
            run = para.add_run(element.get_text())
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
        elif element.name == 'a':
            para.add_run(element.get_text())
        else:
            for child in element.children:
                process_inline(para, child)

    def process_element(el):
        if isinstance(el, str):
            return

        if el.name in ('h1','h2','h3','h4','h5','h6'):
            level = int(el.name[1])
            doc.add_heading(el.get_text(), level=level)

        elif el.name == 'p':
            para = doc.add_paragraph()
            for child in el.children:
                process_inline(para, child)

        elif el.name == 'ul':
            for li in el.find_all('li', recursive=False):
                para = doc.add_paragraph(style='List Bullet')
                for child in li.children:
                    if hasattr(child, 'name') and child.name in ('ul','ol'):
                        continue  # skip nested for now
                    process_inline(para, child)

        elif el.name == 'ol':
            for li in el.find_all('li', recursive=False):
                para = doc.add_paragraph(style='List Number')
                for child in li.children:
                    if hasattr(child, 'name') and child.name in ('ul','ol'):
                        continue
                    process_inline(para, child)

        elif el.name == 'pre':
            code_text = el.get_text()
            para = doc.add_paragraph()
            run = para.add_run(code_text)
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
            pPr = para._p.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'F4F4F5')
            pPr.append(shd)

        elif el.name == 'blockquote':
            para = doc.add_paragraph(el.get_text().strip())
            para.style = doc.styles['Normal']
            run = para.runs[0] if para.runs else para.add_run('')
            run.italic = True

        elif el.name == 'hr':
            para = doc.add_paragraph()
            pPr = para._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'AAAAAA')
            pBdr.append(bottom)
            pPr.append(pBdr)

        elif el.name == 'table':
            rows = el.find_all('tr')
            if not rows:
                return
            max_cols = max(len(r.find_all(['td','th'])) for r in rows)
            if max_cols == 0:
                return
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.style = 'Table Grid'
            for r_idx, row in enumerate(rows):
                cells = row.find_all(['td','th'])
                for c_idx, cell in enumerate(cells):
                    if c_idx >= max_cols:
                        break
                    tc = table.rows[r_idx].cells[c_idx]
                    tc.text = cell.get_text(strip=True)
                    # Bold header row
                    if cell.name == 'th' or r_idx == 0:
                        for run in tc.paragraphs[0].runs:
                            run.bold = True
                        # pyrefly: ignore [missing-import]
                        from docx.oxml import OxmlElement as OE
                        tcPr = tc._tc.get_or_add_tcPr()
                        shd = OE('w:shd')
                        shd.set(qn('w:val'), 'clear')
                        shd.set(qn('w:color'), 'auto')
                        shd.set(qn('w:fill'), '2C3E50')
                        tcPr.append(shd)
                        for para in tc.paragraphs:
                            for run in para.runs:
                                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for element in soup.children:
        process_element(element)

    doc.save(str(out))

def md_to_html(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    out.write_text(md_to_html_string(md_text), encoding='utf-8')

def md_to_txt(src: Path, out: Path):
    out.write_text(src.read_text(encoding='utf-8', errors='replace'), encoding='utf-8')


# ── DOCX → * ──────────────────────────────────────────────────────────────────
def docx_to_pdf(src: Path, out: Path):
    # Try MS Word first (Windows) → LibreOffice → HTML chain fallback
    if docx_to_pdf_word(src, out):
        return
    if docx_to_pdf_libreoffice(src, out):
        return
    # HTML chain fallback
    html_str = docx_to_html_string(src)
    html_string_to_pdf(html_str, out)

def docx_to_md(src: Path, out: Path):
    doc = Document(str(src))
    lines = []
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = next((p for p in doc.paragraphs if p._element is child), None)
            if para is None: continue
            sn = style_name(para)
            text = para.text or ''
            if not text.strip(): lines.append(''); continue
            if 'heading 1' in sn: lines.append(f'# {text}')
            elif 'heading 2' in sn: lines.append(f'## {text}')
            elif 'heading 3' in sn: lines.append(f'### {text}')
            elif 'heading 4' in sn: lines.append(f'#### {text}')
            elif 'list bullet' in sn: lines.append(f'- {text}')
            elif 'list number' in sn: lines.append(f'1. {text}')
            else: lines.append(text)
        elif tag == 'tbl':
            tbl = next((t for t in doc.tables if t._element is child), None)
            if tbl is None: continue
            rows = tbl.rows
            if not rows: continue
            header = '| ' + ' | '.join(html_mod.escape(c.text) for c in rows[0].cells) + ' |'
            sep    = '| ' + ' | '.join('---' for _ in rows[0].cells) + ' |'
            lines.append(header)
            lines.append(sep)
            for row in rows[1:]:
                lines.append('| ' + ' | '.join(html_mod.escape(c.text) for c in row.cells) + ' |')
            lines.append('')
    out.write_text('\n'.join(lines), encoding='utf-8')

def docx_to_txt(src: Path, out: Path):
    doc = Document(str(src))
    lines = []
    for para in doc.paragraphs:
        lines.append(para.text)
    # Also extract table text
    for table in doc.tables:
        for row in table.rows:
            lines.append('\t'.join(c.text for c in row.cells))
    out.write_text('\n'.join(lines), encoding='utf-8')


# ── TXT → * ───────────────────────────────────────────────────────────────────
def txt_to_pdf(src: Path, out: Path):
    text = src.read_text(encoding='utf-8', errors='replace')
    # Wrap as preformatted HTML for clean rendering
    html_str = f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<style>{PDF_CSS} body{{white-space:pre-wrap;font-family:'JetBrains Mono','Courier New',monospace;font-size:9.5pt;}}</style>
</head><body>{html_mod.escape(text)}</body></html>"""
    html_string_to_pdf(html_str, out)

def txt_to_md(src: Path, out: Path):
    out.write_text(src.read_text(encoding='utf-8', errors='replace'), encoding='utf-8')

def txt_to_docx(src: Path, out: Path):
    text = src.read_text(encoding='utf-8', errors='replace')
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    doc.save(str(out))


# ── HTML → * ──────────────────────────────────────────────────────────────────
def html_to_pdf(src: Path, out: Path):
    if WEASYPRINT_AVAILABLE:
        html_file_to_pdf_weasyprint(src, out)
    else:
        _html_to_pdf_pymupdf(src.read_text(encoding='utf-8', errors='replace'), out)

def html_to_md(src: Path, out: Path):
    soup = BeautifulSoup(src.read_text(encoding='utf-8', errors='replace'), 'html.parser')
    lines = []
    for el in soup.find_all(['h1','h2','h3','h4','h5','h6','p','li','pre','hr','blockquote','table']):
        n = el.name
        text = el.get_text(' ', strip=True)
        if   n == 'h1': lines += [f'# {text}', '']
        elif n == 'h2': lines += [f'## {text}', '']
        elif n == 'h3': lines += [f'### {text}', '']
        elif n == 'h4': lines += [f'#### {text}', '']
        elif n == 'h5': lines += [f'##### {text}', '']
        elif n == 'li': lines.append(f'- {text}')
        elif n == 'pre': lines += [f'```\n{el.get_text()}\n```', '']
        elif n == 'blockquote': lines += [f'> {text}', '']
        elif n == 'hr': lines += ['---', '']
        elif n == 'table':
            rows = el.find_all('tr')
            if rows:
                header = '| ' + ' | '.join(c.get_text(strip=True) for c in rows[0].find_all(['th','td'])) + ' |'
                sep    = '| ' + ' | '.join('---' for _ in rows[0].find_all(['th','td'])) + ' |'
                lines += [header, sep]
                for row in rows[1:]:
                    lines.append('| ' + ' | '.join(c.get_text(strip=True) for c in row.find_all('td')) + ' |')
                lines.append('')
        elif text: lines += [text, '']
    out.write_text('\n'.join(lines), encoding='utf-8')

def html_to_docx(src: Path, out: Path):
    # Reuse md_to_docx logic via html soup
    html_content = src.read_text(encoding='utf-8', errors='replace')
    soup = BeautifulSoup(html_content, 'html.parser')
    body = soup.find('body') or soup
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    for el in body.children:
        if not hasattr(el, 'name') or el.name is None:
            continue
        n = el.name
        text = el.get_text(strip=True)
        if   n == 'h1': doc.add_heading(text, 1)
        elif n == 'h2': doc.add_heading(text, 2)
        elif n == 'h3': doc.add_heading(text, 3)
        elif n == 'h4': doc.add_heading(text, 4)
        elif n == 'p':  doc.add_paragraph(text)
        elif n == 'ul':
            for li in el.find_all('li', recursive=False):
                doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
        elif n == 'ol':
            for li in el.find_all('li', recursive=False):
                doc.add_paragraph(li.get_text(strip=True), style='List Number')
        elif n == 'pre':
            p = doc.add_paragraph()
            run = p.add_run(el.get_text())
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
        elif n == 'table':
            rows = el.find_all('tr')
            if not rows: continue
            max_cols = max(len(r.find_all(['td','th'])) for r in rows)
            if max_cols == 0: continue
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.style = 'Table Grid'
            for r_i, row in enumerate(rows):
                cells = row.find_all(['td','th'])
                for c_i, cell in enumerate(cells):
                    if c_i >= max_cols: break
                    table.rows[r_i].cells[c_i].text = cell.get_text(strip=True)
    doc.save(str(out))


# ─── DISPATCH TABLE ────────────────────────────────────────────────────────────
DISPATCH = {
    ('pdf',  'md'):   pdf_to_md,
    ('pdf',  'txt'):  pdf_to_txt,
    ('pdf',  'docx'): pdf_to_docx,
    ('md',   'pdf'):  md_to_pdf,
    ('md',   'docx'): md_to_docx,
    ('md',   'html'): md_to_html,
    ('md',   'txt'):  md_to_txt,
    ('docx', 'pdf'):  docx_to_pdf,
    ('docx', 'md'):   docx_to_md,
    ('docx', 'txt'):  docx_to_txt,
    ('txt',  'pdf'):  txt_to_pdf,
    ('txt',  'md'):   txt_to_md,
    ('txt',  'docx'): txt_to_docx,
    ('html', 'pdf'):  html_to_pdf,
    ('html', 'md'):   html_to_md,
    ('html', 'docx'): html_to_docx,
}


# ─── API ───────────────────────────────────────────────────────────────────────
@app.get('/formats')
def get_formats():
    return CONVERSIONS

@app.post('/convert')
async def convert(file: UploadFile = File(...), target_format: str = Form(...)):
    src_ext = Path(file.filename).suffix.lstrip('.').lower()
    if src_ext not in CONVERSIONS:
        raise HTTPException(400, f"Format '.{src_ext}' not supported")
    if target_format not in CONVERSIONS[src_ext]:
        raise HTTPException(400, f"Cannot convert .{src_ext} → .{target_format}")
    fn = DISPATCH.get((src_ext, target_format))
    if not fn:
        raise HTTPException(400, "Conversion not available")

    job = UPLOAD_DIR / uuid.uuid4().hex
    job.mkdir()
    src_path = job / file.filename
    out_path = job / f"{Path(file.filename).stem}.{target_format}"

    try:
        with open(src_path, 'wb') as f:
            shutil.copyfileobj(file.file, f)
        fn(src_path, out_path)
    except Exception as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {e}")

    if not out_path.exists():
        raise HTTPException(500, "Output file was not created")

    return FileResponse(str(out_path), filename=out_path.name,
                        media_type='application/octet-stream')

@app.get('/', response_class=HTMLResponse)
def serve_ui():
    return (Path(__file__).parent / 'index.html').read_text(encoding='utf-8')

if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))