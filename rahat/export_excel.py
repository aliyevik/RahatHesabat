"""Hesabatı Excel-ə (albom formatında) yazır."""
from __future__ import annotations

import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from . import theme
from .layout import Grid, NOTE

MONEY = '#,##0.00;[Red]-#,##0.00;"-"'


def _hex(c):
    return c.lstrip("#").upper()


def save_grid_xlsx(grid: Grid, path: str, sheet_title: str = "Hesabat",
                   footer_text: str = "") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.sheet_view.showGridLines = False

    thin = Side(style="thin", color=_hex(theme.BORDER))
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    last_col = get_column_letter(grid.ncols)

    # Başlıq
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = grid.title
    ws["A1"].font = Font(size=16, bold=True, color=_hex(theme.NAVY))
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = grid.subtitle
    ws["A2"].font = Font(size=11, italic=True, color=_hex(theme.MUTED))
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 24

    row = 4
    for label, val in grid.top_lines:
        ws.cell(row, 1, label).font = Font(bold=True, color=_hex(theme.NAVY))
        c = ws.cell(row, 2, val)
        c.font = Font(bold=True)
        if isinstance(val, (int, float)):
            c.number_format = MONEY
        row += 1
    row += 1

    header_start = row
    all_rows = [(r, True) for r in grid.head_rows] + [(r, False) for r in grid.rows]
    for cells, is_head in all_rows:
        is_blank = cells and cells[0] is not None and cells[0].style == NOTE
        if is_blank:
            row += 1
            continue
        for ci, cell in enumerate(cells, start=1):
            if cell is None:
                continue
            c = ws.cell(row, ci, cell.value)
            bg = theme.STYLE_BG.get(cell.style)
            if bg:
                c.fill = PatternFill("solid", fgColor=_hex(bg))
            c.font = Font(bold=cell.style in theme.STYLE_BOLD,
                          color=_hex(theme.STYLE_FG.get(cell.style, theme.TEXT)),
                          size=9 if cell.style == "sub" else 10)
            c.alignment = Alignment(horizontal=cell.align, vertical="center", wrap_text=True)
            if isinstance(cell.value, (int, float)):
                c.number_format = MONEY
            if cell.colspan > 1 or cell.rowspan > 1:
                ws.merge_cells(start_row=row, start_column=ci,
                               end_row=row + cell.rowspan - 1, end_column=ci + cell.colspan - 1)
            for rr in range(row, row + cell.rowspan):
                for cc in range(ci, ci + cell.colspan):
                    ws.cell(rr, cc).border = border
                    if bg and (rr, cc) != (row, ci):
                        ws.cell(rr, cc).fill = PatternFill("solid", fgColor=_hex(bg))
        if cells and cells[0] is not None and cells[0].style == "head":
            ws.row_dimensions[row].height = 30
        row += 1

    # Sütun enləri
    for i, w in enumerate(grid.col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, w * 10)

    ws.freeze_panes = f"{get_column_letter(2 if grid.col_widths[0] > 2 else 3)}{header_start + len(grid.head_rows)}"

    row += 1
    if grid.footnote:
        ws.cell(row, 1, grid.footnote).font = Font(italic=True, size=8, color=_hex(theme.MUTED))
        row += 2
    for line in grid.signature:
        ws.cell(row, 1, line).font = Font(bold=True, color=_hex(theme.NAVY))
        row += 2

    # Çap: albom, eni bir səhifəyə
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.5)
    ws.print_title_rows = f"{header_start}:{header_start + len(grid.head_rows) - 1}"
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = footer_text
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = "Səhifə &P / &N"
    ws.oddFooter.right.size = 8

    wb.properties.creator = "Rahat Hesabat — FMC"
    wb.properties.created = dt.datetime.now()
    wb.save(path)
