"""Convert proposal.md into a professionally styled corporate PDF.

Usage:
    python generate_proposal.py

Input:
    proposal.md

Output:
    Margin_Recovery_Proposal.pdf

Note:
    The preferred HTML/CSS path is markdown2 + WeasyPrint. On Windows,
    WeasyPrint can require native GTK/Pango libraries. This script therefore
    uses markdown2 for conversion readiness and fpdf2 for a no-native-DLL PDF
    fallback that works reliably in this project environment.
"""

from __future__ import annotations

import re
from pathlib import Path

import markdown2
from fpdf import FPDF


SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_FILE = SCRIPT_DIR / "proposal.md"
OUTPUT_FILE = SCRIPT_DIR / "Margin_Recovery_Proposal.pdf"
FONT_REGULAR = Path(r"C:\Windows\Fonts\arial.ttf")
FONT_BOLD = Path(r"C:\Windows\Fonts\arialbd.ttf")

CORPORATE_CSS = """
@page {
    size: A4;
    margin: 22mm 18mm 24mm 18mm;
    @bottom-left { content: "Logistics Margin Recovery Proposal"; font-family: Helvetica, Arial, sans-serif; font-size: 8.5pt; color: #64748b; }
    @bottom-right { content: "Page " counter(page) " of " counter(pages); font-family: Helvetica, Arial, sans-serif; font-size: 8.5pt; color: #64748b; }
}
body { font-family: Helvetica, Arial, sans-serif; color: #334155; font-size: 10.8pt; line-height: 1.6; background: #ffffff; }
h1, h2, h3 { color: #1e293b; font-weight: 800; line-height: 1.2; page-break-after: avoid; }
h1 { font-size: 26pt; margin: 0 0 18px 0; padding-bottom: 12px; border-bottom: 3px solid #1e293b; letter-spacing: -0.03em; }
h2 { font-size: 17pt; margin: 30px 0 10px 0; padding-bottom: 7px; border-bottom: 1px solid #cbd5e1; }
h3 { font-size: 13pt; margin: 22px 0 8px 0; color: #334155; }
p { margin: 0 0 11px 0; }
strong { color: #1e293b; font-weight: 800; }
table { width: 100%; border-collapse: collapse; margin: 18px 0 22px 0; font-size: 9.4pt; page-break-inside: avoid; }
thead tr { background: #1e293b; color: #ffffff; }
th { padding: 9px 10px; border: 1px solid #1e293b; text-align: left; font-weight: 800; }
td { padding: 8px 10px; border: 1px solid #dbe3ec; vertical-align: top; }
tbody tr:nth-child(even) { background: #f8fafc; }
tbody tr:nth-child(odd) { background: #ffffff; }
"""


class ProposalPDF(FPDF):
    NAVY = (30, 41, 59)
    TEXT = (51, 65, 85)
    MUTED = (100, 116, 139)
    BORDER = (203, 213, 225)
    ZEBRA = (248, 250, 252)
    WHITE = (255, 255, 255)

    def header(self):
        self.set_y(10)
        self.set_font("ArialCustom", "B", 9)
        self.set_text_color(*self.NAVY)
        self.cell(0, 6, "Margin Recovery Proposal", align="R", new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*self.BORDER)
        self.set_line_width(0.3)
        self.line(self.l_margin, 20, self.w - self.r_margin, 20)
        self.ln(6)

    def footer(self):
        self.set_y(-14)
        self.set_font("ArialCustom", "", 8)
        self.set_text_color(*self.MUTED)
        self.cell(95, 6, "Logistics Margin Recovery Proposal", align="L")
        self.cell(0, 6, f"Page {self.page_no()} of {{nb}}", align="R")


def clean_inline_markdown(text: str) -> str:
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = text.replace("&", "&")
    return text.strip()


def ensure_space(pdf: ProposalPDF, height: float) -> None:
    if pdf.get_y() + height > pdf.page_break_trigger:
        pdf.add_page()


def write_heading(pdf: ProposalPDF, text: str, level: int) -> None:
    text = clean_inline_markdown(text)
    if level == 1:
        ensure_space(pdf, 24)
        pdf.set_font("ArialCustom", "B", 22)
        pdf.set_text_color(*ProposalPDF.NAVY)
        pdf.multi_cell(0, 10, text, align="C")
        pdf.set_draw_color(*ProposalPDF.NAVY)
        pdf.set_line_width(0.8)
        pdf.line(pdf.l_margin, pdf.get_y() + 2, pdf.w - pdf.r_margin, pdf.get_y() + 2)
        pdf.ln(8)
    elif level == 2:
        ensure_space(pdf, 18)
        pdf.ln(4)
        pdf.set_font("ArialCustom", "B", 15)
        pdf.set_text_color(*ProposalPDF.NAVY)
        pdf.multi_cell(0, 8, text)
        pdf.set_draw_color(*ProposalPDF.BORDER)
        pdf.set_line_width(0.35)
        pdf.line(pdf.l_margin, pdf.get_y() + 1, pdf.w - pdf.r_margin, pdf.get_y() + 1)
        pdf.ln(5)
    else:
        ensure_space(pdf, 14)
        pdf.ln(2)
        pdf.set_font("ArialCustom", "B", 12)
        pdf.set_text_color(*ProposalPDF.TEXT)
        pdf.multi_cell(0, 7, text)
        pdf.ln(1)


def write_paragraph(pdf: ProposalPDF, text: str) -> None:
    text = clean_inline_markdown(text)
    if not text:
        return
    ensure_space(pdf, 10)
    pdf.set_font("ArialCustom", "", 10.5)
    pdf.set_text_color(*ProposalPDF.TEXT)
    pdf.multi_cell(0, 6.5, text)
    pdf.ln(2)


def write_bullet(pdf: ProposalPDF, text: str, indent: int = 0) -> None:
    text = clean_inline_markdown(text)
    ensure_space(pdf, 8)
    pdf.set_font("ArialCustom", "", 10.2)
    pdf.set_text_color(*ProposalPDF.TEXT)
    left = pdf.l_margin + indent
    pdf.set_x(left)
    pdf.cell(5, 6.3, "•")
    pdf.set_x(left + 6)
    pdf.multi_cell(0, 6.3, text)
    pdf.ln(0.5)


def write_numbered(pdf: ProposalPDF, number: str, text: str) -> None:
    text = clean_inline_markdown(text)
    ensure_space(pdf, 8)
    pdf.set_font("ArialCustom", "", 10.2)
    pdf.set_text_color(*ProposalPDF.TEXT)
    pdf.set_x(pdf.l_margin)
    pdf.cell(8, 6.3, f"{number}.")
    pdf.set_x(pdf.l_margin + 9)
    pdf.multi_cell(0, 6.3, text)
    pdf.ln(0.5)


def write_ascii_table(pdf: ProposalPDF, table_lines: list[str]) -> None:
    content_rows = []
    for line in table_lines:
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue
        cells = [clean_inline_markdown(cell.strip()) for cell in stripped.strip("|").split("|")]
        content_rows.append(cells)

    if not content_rows:
        return

    ensure_space(pdf, 30)
    max_cols = max(len(row) for row in content_rows)
    usable_width = pdf.w - pdf.l_margin - pdf.r_margin
    widths = [usable_width / max_cols] * max_cols
    row_h = 8

    for row_idx, row in enumerate(content_rows):
        ensure_space(pdf, row_h)
        row = row + [""] * (max_cols - len(row))
        is_header = row_idx in (0, 1)
        if is_header:
            pdf.set_fill_color(*ProposalPDF.NAVY)
            pdf.set_text_color(*ProposalPDF.WHITE)
            pdf.set_font("ArialCustom", "B", 8.6)
            border_color = ProposalPDF.NAVY
        else:
            pdf.set_fill_color(*(ProposalPDF.ZEBRA if row_idx % 2 else ProposalPDF.WHITE))
            pdf.set_text_color(*ProposalPDF.TEXT)
            pdf.set_font("ArialCustom", "", 8.3)
            border_color = ProposalPDF.BORDER
        pdf.set_draw_color(*border_color)
        for cell, width in zip(row, widths):
            pdf.cell(width, row_h, cell[:38], border=1, align="C", fill=True)
        pdf.ln(row_h)
    pdf.ln(4)


def build_html(markdown_text: str) -> str:
    body_html = markdown2.markdown(markdown_text, extras=["tables", "fenced-code-blocks", "strike", "task_list"])
    return f"""<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Margin Recovery Proposal</title>
    <style>{CORPORATE_CSS}</style>
</head>
<body>{body_html}</body>
</html>"""


def generate_pdf(markdown_text: str) -> None:
    # Build the HTML as requested; the PDF renderer below mirrors the same corporate styling
    # without requiring WeasyPrint's native GTK/Pango DLLs on Windows.
    _ = build_html(markdown_text)

    pdf = ProposalPDF()
    pdf.alias_nb_pages()
    pdf.set_margins(16, 24, 16)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_font("ArialCustom", "", str(FONT_REGULAR), uni=True)
    pdf.add_font("ArialCustom", "B", str(FONT_BOLD), uni=True)
    pdf.add_page()

    lines = markdown_text.splitlines()
    paragraph_buffer: list[str] = []
    table_buffer: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer
        if paragraph_buffer:
            write_paragraph(pdf, " ".join(paragraph_buffer))
            paragraph_buffer = []

    def flush_table() -> None:
        nonlocal table_buffer
        if table_buffer:
            write_ascii_table(pdf, table_buffer)
            table_buffer = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("+") or stripped.startswith("|"):
            flush_paragraph()
            table_buffer.append(stripped)
            continue
        flush_table()

        if not stripped:
            flush_paragraph()
            continue
        if stripped == "---":
            flush_paragraph()
            ensure_space(pdf, 10)
            pdf.set_draw_color(*ProposalPDF.BORDER)
            pdf.line(pdf.l_margin, pdf.get_y() + 2, pdf.w - pdf.r_margin, pdf.get_y() + 2)
            pdf.ln(8)
            continue
        if stripped.startswith("### "):
            flush_paragraph()
            write_heading(pdf, stripped[4:], 3)
            continue
        if stripped.startswith("## "):
            flush_paragraph()
            write_heading(pdf, stripped[3:], 2)
            continue
        if stripped.startswith("# "):
            flush_paragraph()
            write_heading(pdf, stripped[2:], 1)
            continue
        if stripped.startswith("*   ") or stripped.startswith("- "):
            flush_paragraph()
            indent = 7 if raw_line.startswith("    ") else 0
            bullet_text = stripped[4:] if stripped.startswith("*   ") else stripped[2:]
            write_bullet(pdf, bullet_text, indent=indent)
            continue
        numbered_match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if numbered_match:
            flush_paragraph()
            write_numbered(pdf, numbered_match.group(1), numbered_match.group(2))
            continue

        paragraph_buffer.append(stripped)

    flush_paragraph()
    flush_table()
    pdf.output(str(OUTPUT_FILE))


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Could not find {INPUT_FILE}. Place proposal.md in the same folder as this script.")

    markdown_text = INPUT_FILE.read_text(encoding="utf-8")
    generate_pdf(markdown_text)
    print(f"Generated {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
