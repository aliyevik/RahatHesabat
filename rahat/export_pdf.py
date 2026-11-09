"""Hesabatı PDF-ə (albom formatında) yazır."""
from __future__ import annotations

import os
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle, KeepTogether)
from reportlab.lib.styles import ParagraphStyle

from . import theme
from .layout import Grid, NOTE
from .model import fmt_money

_FONTS_OK = False


def resource_path(*parts) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, *parts)


def _register_fonts():
    global _FONTS_OK
    if _FONTS_OK:
        return
    pdfmetrics.registerFont(TTFont("DejaVu", resource_path("assets", "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", resource_path("assets", "DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold")
    _FONTS_OK = True


def _c(h):
    return colors.HexColor(h)


def _fmt(v):
    if isinstance(v, (int, float)):
        return "–" if abs(v) < 0.005 else fmt_money(v)
    return "" if v is None else str(v)


def _fit_font(grid: Grid, widths) -> float:
    """Ən uzun rəqəm öz sütununa sığacaq şrift ölçüsü."""
    need = [0.0] * grid.ncols
    for cells in grid.rows:
        for ci, cell in enumerate(cells):
            if cell is not None and isinstance(cell.value, (int, float)) and cell.colspan == 1:
                need[ci] = max(need[ci], pdfmetrics.stringWidth(_fmt(cell.value), "DejaVu-Bold", 1))
    size = 8.5
    for w, n in zip(widths, need):
        if n:
            size = min(size, (w - 4.5) / n)
    return size


def save_grid_pdf(grid: Grid, path: str, paper: str = "auto", footer_text: str = "") -> None:
    _register_fonts()
    margin = 8 * mm
    total_w = sum(grid.col_widths)
    for size_name, ps in (("A4", A4), ("A3", A3)):
        pagesize = landscape(ps)
        usable = pagesize[0] - 2 * margin
        widths = [usable * w / total_w for w in grid.col_widths]
        fsize = _fit_font(grid, widths)
        if paper == size_name or (paper == "auto" and fsize >= 5.2):
            break
    fsize = max(4.0, fsize)

    cell_style = ParagraphStyle("c", fontName="DejaVu", fontSize=fsize, leading=fsize * 1.2)
    head_style = ParagraphStyle("h", parent=cell_style, fontName="DejaVu-Bold",
                                textColor=colors.white, alignment=1)
    sub_size = fsize * 0.9
    sub_style = ParagraphStyle("hs", parent=head_style, fontSize=sub_size, leading=sub_size * 1.15)
    num_w = min(widths[2:]) if len(widths) > 2 else widths[-1]
    short = pdfmetrics.stringWidth("Hesablanma", "DejaVu-Bold", sub_size) > num_w - 4

    data, styles = [], []
    all_rows = grid.head_rows + grid.rows
    for ri, cells in enumerate(all_rows):
        if cells and cells[0] is not None and cells[0].style == NOTE:
            data.append([""] * grid.ncols)
            styles += [("LINEBELOW", (0, ri), (-1, ri), 0, colors.white),
                       ("LINEABOVE", (0, ri), (-1, ri), 0, colors.white),
                       ("LINEBEFORE", (0, ri), (-1, ri), 0, colors.white),
                       ("LINEAFTER", (0, ri), (-1, ri), 0, colors.white),
                       ("BACKGROUND", (0, ri), (-1, ri), colors.white),
                       ("TOPPADDING", (0, ri), (-1, ri), 1),
                       ("BOTTOMPADDING", (0, ri), (-1, ri), 1)]
            continue
        line = []
        for ci, cell in enumerate(cells):
            if cell is None:
                line.append("")
                continue
            if cell.style in ("head", "sub"):
                txt = (cell.value or "")
                if cell.style == "sub" and short:
                    txt = {"Hesablanma": "Hesabl.", "Ödəmə": "Ödəmə", "Qalıq": "Qalıq"}.get(txt, txt)
                txt = txt.replace("\n", "<br/>")
                line.append(Paragraph(txt, sub_style if cell.style == "sub" else head_style))
            elif cell.align == "left" and isinstance(cell.value, str):
                st = cell_style if cell.style not in theme.STYLE_BOLD else \
                    ParagraphStyle("b", parent=cell_style, fontName="DejaVu-Bold")
                line.append(Paragraph(cell.value, st))
            else:
                line.append(_fmt(cell.value))
            bg = theme.STYLE_BG.get(cell.style)
            end = (ci + cell.colspan - 1, ri + cell.rowspan - 1)
            if cell.colspan > 1 or cell.rowspan > 1:
                styles.append(("SPAN", (ci, ri), end))
            if bg:
                styles.append(("BACKGROUND", (ci, ri), end, _c(bg)))
            if cell.style in theme.STYLE_BOLD:
                styles.append(("FONTNAME", (ci, ri), end, "DejaVu-Bold"))
            styles.append(("ALIGN", (ci, ri), end,
                           {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}[cell.align]))
        data.append(line)

    base = [
        ("FONTNAME", (0, 0), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 0), (-1, -1), fsize),
        ("LEADING", (0, 0), (-1, -1), fsize * 1.2),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, _c(theme.BORDER)),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]
    # Zebra (yalnız adi sətirlər)
    for ri, cells in enumerate(all_rows):
        if ri >= len(grid.head_rows) and cells and cells[0] is not None \
                and cells[0].style == "item" and ri % 2 == 0:
            base.append(("BACKGROUND", (0, ri), (-1, ri), _c(theme.BLUE_PALE)))
    table = Table(data, colWidths=widths, repeatRows=len(grid.head_rows))
    table.setStyle(TableStyle(base + styles))

    title_st = ParagraphStyle("t", fontName="DejaVu-Bold", fontSize=15, textColor=_c(theme.NAVY),
                              alignment=1, leading=19)
    sub_st = ParagraphStyle("s", fontName="DejaVu", fontSize=9.5, textColor=_c(theme.MUTED),
                            alignment=1, leading=12)
    info_st = ParagraphStyle("i", fontName="DejaVu", fontSize=9, textColor=_c(theme.TEXT), leading=12)
    note_st = ParagraphStyle("n", fontName="DejaVu", fontSize=7, textColor=_c(theme.MUTED), leading=9)
    sig_st = ParagraphStyle("g", fontName="DejaVu-Bold", fontSize=9.5, textColor=_c(theme.NAVY),
                            leading=13)

    story = [Paragraph(grid.title, title_st), Paragraph(grid.subtitle, sub_st), Spacer(1, 4 * mm)]
    for label, val in grid.top_lines:
        v = fmt_money(val) + " AZN" if isinstance(val, (int, float)) else str(val)
        story.append(Paragraph(f"{label} <b>{v}</b>", info_st))
    story += [Spacer(1, 3 * mm), table, Spacer(1, 2 * mm)]
    if grid.footnote:
        story.append(Paragraph(grid.footnote, note_st))
    if grid.signature:
        sig = [Spacer(1, 8 * mm)]
        for s in grid.signature:
            sig += [Paragraph(s, sig_st), Spacer(1, 5 * mm)]
        story.append(KeepTogether(sig))

    def on_page(canvas, doc):
        canvas.saveState()
        w, h = pagesize
        canvas.setFillColor(_c(theme.NAVY))
        canvas.rect(0, h - 4 * mm, w, 4 * mm, stroke=0, fill=1)
        canvas.setFillColor(_c(theme.BLUE_MID))
        canvas.rect(w * 0.72, h - 4 * mm, w * 0.28, 4 * mm, stroke=0, fill=1)
        canvas.setFont("DejaVu", 7)
        canvas.setFillColor(_c(theme.MUTED))
        canvas.drawString(margin, 5 * mm, footer_text)
        canvas.drawRightString(w - margin, 5 * mm, f"Səhifə {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(path, pagesize=pagesize, leftMargin=margin, rightMargin=margin,
                            topMargin=10 * mm, bottomMargin=11 * mm,
                            title=grid.title, author="Rahat Hesabat — FMC")
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
