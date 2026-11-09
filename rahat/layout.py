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
    ref: tuple | None = None    # ətraflı baxış (kartoçka) üçün açar
    drag: tuple | None = None   # (növ, maddə) — sətri tutub dartmaq üçün


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
    stacked: bool = False       # alt-alta görünüş (yalnız ekran üçün)


def _num(v, style=ITEM, ref=None):
    return Cell(round(v, 2), style, ref=ref)


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

    def body(rows, sumfn, label, kind):
        out = []
        for x in rows:
            it = x.item
            line = [Cell(x.item, ITEM, align="left", drag=(kind, it)), _num(x.opening)]
            for m in M:
                line += [_num(x.h.get(m, 0.0), ref=("pl", kind, it, m, "h")),
                         _num(x.o.get(m, 0.0), ref=("pl", kind, it, m, "o"))]
            line += [_num(x.total_h(), ref=("pl", kind, it, None, "h")),
                     _num(x.total_o(), ref=("pl", kind, it, None, "o")), _num(x.closing())]
            out.append(line)
        tot = [Cell(label, TOTAL, align="left"), _num(sumfn("opening"), TOTAL)]
        for m in M:
            tot += [_num(sumfn("h", m), TOTAL, ("pl", kind, None, m, "h")),
                    _num(sumfn("o", m), TOTAL, ("pl", kind, None, m, "o"))]
        tot += [_num(sumfn("h"), TOTAL, ("pl", kind, None, None, "h")),
                _num(sumfn("o"), TOTAL, ("pl", kind, None, None, "o")),
                _num(sumfn("closing"), TOTAL)]
        out.append(tot)
        return out

    # Gəlirlər
    g.head_rows = header("Gəlirlər", "Alınacaqlar", "Dövr sonuna alınacaqlar", "Cəmi gəlirlər")
    g.rows += body(r.gelir, r.gelir_sum, "Cəmi gəlirlər:", "gelir")
    g.rows.append([Cell(None, NOTE)] + [None] * (ncols - 1))
    # Xərclər (öz başlığı ilə)
    for hr in header("Xərclər", "Ödəniləcəklər", "Dövr sonuna ödəniləcəklər", "Cəmi xərclər"):
        g.rows.append(hr)
    g.rows += body(r.gider, r.gider_sum, "Cəmi xərclər:", "gider")
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
        k = (x.daire, x.borclu)
        line = [Cell(x.daire, ITEM, align="center"), Cell(x.borclu, ITEM, align="left")]
        line += [_num(v, ref=("debt", *k, "op")) for v in (x.op_h, x.op_o, x.op_q)]
        for m in M:
            line += [_num(v, ref=("debt", *k, m)) for v in (x.h.get(m, 0.0), x.o.get(m, 0.0), x.q.get(m, 0.0))]
        line += [_num(v, ref=("debt", *k, None))
                 for v in (x.op_h + x.total("h"), x.op_o + x.total("o"), x.closing())]
        g.rows.append(line)
    tot = [Cell("CƏMİ", TOTAL, colspan=2, align="left"), None]
    tot += [_num(r.col_sum(a), TOTAL, ("debt", None, None, "op")) for a in ("op_h", "op_o", "op_q")]
    for m in M:
        tot += [_num(r.col_sum(a, m), TOTAL, ("debt", None, None, m)) for a in ("h", "o", "q")]
    tot += [_num(r.col_sum("op_h") + r.col_sum("h"), TOTAL, ("debt", None, None, None)),
            _num(r.col_sum("op_o") + r.col_sum("o"), TOTAL, ("debt", None, None, None)),
            _num(r.col_sum("closing"), TOTAL, ("debt", None, None, None))]
    g.rows.append(tot)
    g.top_lines = [("Kontragent sayı:", f"{len(r.rows)}"),
                   ("Dövrün sonuna qalıq borc:", r.col_sum("closing"))]
    g.footnote = ("Hesablanma – hesablanmış borc, Ödəmə – həmin borca edilmiş ödəniş, "
                  "Qalıq – ödənilməmiş hissə. Mənbə: Apsiyon «Detaylı Borç Listesi».")
    return g


# ---------------------------------------------------------------- Alt-alta görünüş (ekran üçün)
KIND_H, KIND_O, KIND_Q = "Hesablanma", "Ödəmə", "Qalıq"


def pl_grid_stacked(r: PLReport, signature: list[str]) -> Grid:
    """Hər ay bir sütun; hər maddədə Hesablanma və Ödəmə üst-üstə."""
    M = r.months
    ncols = 3 + len(M) + 2
    op_label = fmt_date(r.opening_date)
    g = Grid(title=r.building or "Bina",
             subtitle=f"{fmt_date(r.start)} – {fmt_date(r.end)}", ncols=ncols,
             col_widths=[3.4, 0.9, 1.2] + [1.1] * len(M) + [1.2, 1.25],
             signature=signature, stacked=True)

    def header(name, open_name, close_name, total_name):
        return [Cell(name, HEAD, align="left"), Cell("", HEAD),
                Cell(f"{op_label}\n{open_name}", HEAD, align="center")] + \
               [Cell(month_label(m), HEAD, align="center") for m in M] + \
               [Cell(total_name, HEAD, align="center"),
                Cell(f"{fmt_date(r.end)}\n{close_name}", HEAD, align="center")]

    def pair(label, style, opening, hv, ov, th, to, closing, kind, item, label_style=None):
        st2 = style if style != ITEM else "alt"
        a = [Cell(label, label_style or style, rowspan=2, align="left",
                  drag=(kind, item) if kind and item else None), Cell(KIND_H, style, align="left"),
             Cell(None if opening is None else round(opening, 2), style, rowspan=2)]
        a += [_num(hv(m), style, ("pl", kind, item, m, "h") if kind else None) for m in M]
        a += [_num(th, style, ("pl", kind, item, None, "h") if kind else None),
              Cell(None if closing is None else round(closing, 2), style, rowspan=2)]
        b = [None, Cell(KIND_O, st2, align="left"), None]
        b += [_num(ov(m), st2, ("pl", kind, item, m, "o") if kind else None) for m in M]
        b += [_num(to, st2, ("pl", kind, item, None, "o") if kind else None), None]
        return [a, b]

    g.head_rows = [header("Gəlirlər", "İlk qalıq", "Son qalıq", "Cəmi")]
    for kind, rows, sumfn, label in (("gelir", r.gelir, r.gelir_sum, "Cəmi gəlirlər:"),
                                     ("gider", r.gider, r.gider_sum, "Cəmi xərclər:")):
        if kind == "gider":
            g.rows.append([Cell(None, NOTE)] + [None] * (ncols - 1))
            g.rows.append(header("Xərclər", "İlk qalıq", "Son qalıq", "Cəmi"))
        for x in rows:
            g.rows += pair(x.item, ITEM, x.opening, lambda m, x=x: x.h.get(m, 0.0),
                           lambda m, x=x: x.o.get(m, 0.0), x.total_h(), x.total_o(), x.closing(),
                           kind, x.item)
        g.rows += pair(label, TOTAL, sumfn("opening"), lambda m: sumfn("h", m), lambda m: sumfn("o", m),
                       sumfn("h"), sumfn("o"), sumfn("closing"), kind, None)
    g.rows.append([Cell(None, NOTE)] + [None] * (ncols - 1))
    g.rows += pair("Nəticə: mənfəət (+) / zərər (−)", RESULT, None, lambda m: r.result("h", m),
                   lambda m: r.result("o", m), r.result("h"), r.result("o"), None, None, None)
    cash = [Cell("Pul vəsaitinin qalığı (ay sonuna)", CASH, colspan=2, align="left"), None,
            _num(r.opening_cash, CASH)]
    cash += [_num(r.cash_after(m), CASH) for m in M]
    cash += [_num(r.closing_cash, CASH), Cell(None, CASH)]
    g.rows.append(cash)
    return g


def debt_grid_stacked(r: DebtReport, signature: list[str]) -> Grid:
    """Hər ay bir sütun; hər kontragentdə Hesablanma, Ödəmə, Qalıq üst-üstə."""
    M = r.months
    ncols = 3 + 1 + len(M) + 1
    g = Grid(title=r.building or "Bina", subtitle=f"{fmt_date(r.start)} – {fmt_date(r.end)}",
             ncols=ncols, col_widths=[0.8, 3.0, 0.9, 1.1] + [1.05] * len(M) + [1.2],
             signature=signature, stacked=True)
    g.head_rows = [[Cell("Mənzil", HEAD, align="center"), Cell("Kontragent (borclu)", HEAD, align="left"),
                    Cell("", HEAD), Cell("İlk saldo", HEAD, align="center")] +
                   [Cell(month_label(m), HEAD, align="center") for m in M] +
                   [Cell("Dövr üzrə cəmi", HEAD, align="center")]]

    def triple(c0, c1, vals_op, vals_m, vals_tot, key, style):
        out = []
        for i, (kname, st) in enumerate(((KIND_H, style), (KIND_O, "alt" if style == ITEM else style),
                                         (KIND_Q, "alt2" if style == ITEM else style))):
            row = [Cell(c0, style, rowspan=3, align="center") if i == 0 else None,
                   Cell(c1, style, rowspan=3, align="left") if i == 0 else None,
                   Cell(kname, st, align="left"),
                   _num(vals_op[i], st, ("debt", *key, "op"))]
            row += [_num(vals_m(m)[i], st, ("debt", *key, m)) for m in M]
            row.append(_num(vals_tot[i], st, ("debt", *key, None)))
            out.append(row)
        return out

    for x in r.rows:
        g.rows += triple(x.daire, x.borclu, (x.op_h, x.op_o, x.op_q),
                         lambda m, x=x: (x.h.get(m, 0.0), x.o.get(m, 0.0), x.q.get(m, 0.0)),
                         (x.op_h + x.total("h"), x.op_o + x.total("o"), x.closing()),
                         (x.daire, x.borclu), ITEM)
    g.rows += triple("", "CƏMİ", (r.col_sum("op_h"), r.col_sum("op_o"), r.col_sum("op_q")),
                     lambda m: (r.col_sum("h", m), r.col_sum("o", m), r.col_sum("q", m)),
                     (r.col_sum("op_h") + r.col_sum("h"), r.col_sum("op_o") + r.col_sum("o"),
                      r.col_sum("closing")), (None, None), TOTAL)
    return g
