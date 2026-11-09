"""Hesabatı ümumi cədvəl şəklinə salır (Excel, PDF və ekranda göstərmək üçün)."""
from __future__ import annotations

from dataclasses import dataclass, field

from .model import PLReport, DebtReport, month_label, fmt_date

# Hüceyrə stilləri
TITLE, HEAD, SUB, SECTION, ITEM, TOTAL, RESULT, CASH, NOTE = (
    "title", "head", "sub", "section", "item", "total", "result", "cash", "note")


@dataclass
class Cell:
    value: object = None        # str | float | None
    style: str = ITEM
    colspan: int = 1
    rowspan: int = 1
    align: str = "right"        # left | center | right


@dataclass
class Grid:
    title: str
    subtitle: str
    ncols: int
    head_rows: list[list[Cell]] = field(default_factory=list)
    rows: list[list[Cell]] = field(default_factory=list)
    col_widths: list[float] = field(default_factory=list)   # nisbi en
    top_lines: list[tuple[str, str]] = field(default_factory=list)
    signature: list[str] = field(default_factory=list)
    footnote: str = ""


def _num(v, style=ITEM):
    return Cell(round(v, 2), style)


def _hdr(rows_spec):
    return rows_spec


def pl_grid(r: PLReport, signature: list[str]) -> Grid:
    M = r.months
    ncols = 2 + 2 * len(M) + 3
    op_label = fmt_date(r.opening_date)
    g = Grid(
        title=r.building or "Bina",
        subtitle=f"{fmt_date(r.start)} – {fmt_date(r.end)} tarix aralığı üzrə gəlir və xərclər hesabatı",
        ncols=ncols,
        col_widths=[3.4, 1.25] + [1.1] * (2 * len(M)) + [1.2, 1.2, 1.25],
        signature=signature,
    )
    g.top_lines = [
        (f"Pul vəsaitinin qalığı ({op_label})", None),
    ]

    def header(section_name, open_name, close_name, total_name):
        h1 = [Cell(section_name, HEAD, rowspan=2, align="left"),
              Cell(f"{op_label}\n{open_name}", HEAD, rowspan=2, align="center")]
        for m in M:
            h1.append(Cell(month_label(m), HEAD, colspan=2, align="center"))
            h1.append(None)
        h1 += [Cell(total_name, HEAD, colspan=2, align="center"), None,
               Cell(f"{fmt_date(r.end)}\n{close_name}", HEAD, rowspan=2, align="center")]
        h2 = [None, None]
        for _ in M:
            h2 += [Cell("Hesablanma", SUB, align="center"), Cell("Ödəmə", SUB, align="center")]
        h2 += [Cell("Hesablanma", SUB, align="center"), Cell("Ödəmə", SUB, align="center"), None]
        return [h1, h2]

    def body(rows, sumfn, label):
        out = []
        for x in rows:
            line = [Cell(x.item, ITEM, align="left"), _num(x.opening)]
            for m in M:
                line += [_num(x.h.get(m, 0.0)), _num(x.o.get(m, 0.0))]
            line += [_num(x.total_h()), _num(x.total_o()), _num(x.closing())]
            out.append(line)
        tot = [Cell(label, TOTAL, align="left"), _num(sumfn("opening"), TOTAL)]
        for m in M:
            tot += [_num(sumfn("h", m), TOTAL), _num(sumfn("o", m), TOTAL)]
        tot += [_num(sumfn("h"), TOTAL), _num(sumfn("o"), TOTAL), _num(sumfn("closing"), TOTAL)]
        out.append(tot)
        return out

    # Gəlirlər
    g.head_rows = header("Gəlirlər", "Alınacaqlar", "Dövr sonuna alınacaqlar", "Cəmi gəlirlər")
    g.rows += body(r.gelir, r.gelir_sum, "Cəmi gəlirlər:")
    g.rows.append([Cell(None, NOTE)] + [None] * (ncols - 1))
    # Xərclər (öz başlığı ilə)
    for hr in header("Xərclər", "Ödəniləcəklər", "Dövr sonuna ödəniləcəklər", "Cəmi xərclər"):
        g.rows.append(hr)
    g.rows += body(r.gider, r.gider_sum, "Cəmi xərclər:")
    g.rows.append([Cell(None, NOTE)] + [None] * (ncols - 1))

    # Nəticə: Gəlir − Xərc
    res = [Cell("Nəticə: mənfəət (+) / zərər (−)", RESULT, align="left"), Cell(None, RESULT)]
    for m in M:
        res += [_num(r.result("h", m), RESULT), _num(r.result("o", m), RESULT)]
    res += [_num(r.result("h"), RESULT), _num(r.result("o"), RESULT), Cell(None, RESULT)]
    g.rows.append(res)

    # Pul vəsaitinin qalığı
    cash = [Cell("Pul vəsaitinin qalığı (ay sonuna)", CASH, align="left"), _num(r.opening_cash, CASH)]
    for m in M:
        cash += [Cell(None, CASH), _num(r.cash_after(m), CASH)]
    cash += [Cell(None, CASH), _num(r.closing_cash, CASH), Cell(None, CASH)]
    g.rows.append(cash)

    g.top_lines = [(f"Pul vəsaitinin qalığı ({op_label}):", r.opening_cash),
                   (f"Pul vəsaitinin qalığı ({fmt_date(r.end)}):", r.closing_cash)]
    g.footnote = ("Hesablanma – tahakkuk (hesablama) əsaslı, Ödəmə – ödəniş əsaslı məbləğlər. "
                  "Mənbə: Apsiyon proqramı.")
    return g


def debt_grid(r: DebtReport, signature: list[str]) -> Grid:
    M = r.months
    ncols = 2 + 3 + 3 * len(M) + 3
    g = Grid(
        title=r.building or "Bina",
        subtitle=f"{fmt_date(r.start)} – {fmt_date(r.end)} tarix aralığı üzrə kontragentlərin borc hesabatı",
        ncols=ncols,
        col_widths=[0.8, 3.0] + [0.95] * 3 + [0.95] * (3 * len(M)) + [1.05, 1.05, 1.1],
        signature=signature,
    )
    h1 = [Cell("Mənzil", HEAD, rowspan=2, align="center"),
          Cell("Kontragent (borclu)", HEAD, rowspan=2, align="left"),
          Cell("İlk saldo (keçən borc)", HEAD, colspan=3, align="center"), None, None]
    for m in M:
        h1 += [Cell(month_label(m), HEAD, colspan=3, align="center"), None, None]
    h1 += [Cell("Dövr üzrə cəmi", HEAD, colspan=3, align="center"), None, None]
    h2 = [None, None]
    for _ in range(len(M) + 2):
        h2 += [Cell("Hesablanma", SUB, align="center"), Cell("Ödəmə", SUB, align="center"),
               Cell("Qalıq", SUB, align="center")]
    g.head_rows = [h1, h2]

    for x in r.rows:
        line = [Cell(x.daire, ITEM, align="center"), Cell(x.borclu, ITEM, align="left"),
                _num(x.op_h), _num(x.op_o), _num(x.op_q)]
        for m in M:
            line += [_num(x.h.get(m, 0.0)), _num(x.o.get(m, 0.0)), _num(x.q.get(m, 0.0))]
        line += [_num(x.op_h + x.total("h")), _num(x.op_o + x.total("o")), _num(x.closing())]
        g.rows.append(line)
    tot = [Cell("CƏMİ", TOTAL, colspan=2, align="left"), None,
           _num(r.col_sum("op_h"), TOTAL), _num(r.col_sum("op_o"), TOTAL), _num(r.col_sum("op_q"), TOTAL)]
    for m in M:
        tot += [_num(r.col_sum("h", m), TOTAL), _num(r.col_sum("o", m), TOTAL), _num(r.col_sum("q", m), TOTAL)]
    tot += [_num(r.col_sum("op_h") + r.col_sum("h"), TOTAL),
            _num(r.col_sum("op_o") + r.col_sum("o"), TOTAL),
            _num(r.col_sum("closing"), TOTAL)]
    g.rows.append(tot)
    g.top_lines = [("Kontragent sayı:", f"{len(r.rows)}"),
                   ("Dövrün sonuna qalıq borc:", r.col_sum("closing"))]
    g.footnote = ("Hesablanma – hesablanmış borc, Ödəmə – həmin borca edilmiş ödəniş, "
                  "Qalıq – ödənilməmiş hissə. Mənbə: Apsiyon «Detaylı Borç Listesi».")
    return g
